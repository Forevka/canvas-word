import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WS_MSG, type UserInfo } from "@cw/shared";
import { Broadcaster } from "./Broadcaster";

// Minimal fake `ws` socket: just enough surface for Broadcaster (readyState/OPEN,
// a send spy, and a close handler we can fire manually). Cast to the ws type the
// Broadcaster expects — it only ever touches these members.
interface FakeWs {
  readyState: number;
  OPEN: number;
  send: ReturnType<typeof vi.fn>;
  on: (ev: string, cb: () => void) => void;
  close: () => void;
}

function makeWs(): FakeWs {
  let closeHandler: (() => void) | undefined;
  const ws: FakeWs = {
    OPEN: 1,
    readyState: 1,
    send: vi.fn(),
    on: (ev, cb) => {
      if (ev === "close") closeHandler = cb;
    },
    close: () => closeHandler?.(),
  };
  return ws;
}

// Decode the JSON payloads a socket received, newest-last.
const sent = (ws: FakeWs): unknown[] => ws.send.mock.calls.map((c) => JSON.parse(c[0] as string));
const sentOfType = (ws: FakeWs, type: string): Record<string, unknown>[] =>
  sent(ws).filter((m): m is Record<string, unknown> => (m as { type?: string }).type === type);

// Broadcaster is typed against ws.WebSocket; our fake is structurally compatible
// for the members it uses, so this cast is the test seam.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const asWs = (w: FakeWs): any => w;

const alice: UserInfo = { id: "u1", firstName: "Al", lastName: "Ice" };
const bob: UserInfo = { id: "u2", firstName: "Bo", lastName: "B" };

describe("Broadcaster", () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });
  afterEach(() => {
    vi.useRealTimers();
  });

  describe("hello", () => {
    it("sends the joining client a roster of the existing members", () => {
      const b = new Broadcaster();
      const a = makeWs();
      const c = makeWs();
      b.join("doc", asWs(a));
      b.join("doc", asWs(c));
      // a identifies first so it appears in c's roster.
      b.hello(asWs(a), "siteA", alice);
      b.hello(asWs(c), "siteC", bob);

      const roster = sentOfType(c, WS_MSG.Roster);
      expect(roster).toHaveLength(1);
      expect(roster[0]!.entries).toEqual([{ siteId: "siteA", user: alice, selection: null }]);
    });

    it("announces the new client's presence to the rest of the room but not itself", () => {
      const b = new Broadcaster();
      const a = makeWs();
      const c = makeWs();
      b.join("doc", asWs(a));
      b.join("doc", asWs(c));
      b.hello(asWs(a), "siteA", alice);
      a.send.mockClear();
      c.send.mockClear();

      b.hello(asWs(c), "siteC", bob);

      // a receives c's presence; c does not receive its own.
      const presenceToA = sentOfType(a, WS_MSG.Presence);
      expect(presenceToA).toHaveLength(1);
      expect(presenceToA[0]).toMatchObject({ siteId: "siteC", user: bob, selection: null });
      expect(sentOfType(c, WS_MSG.Presence)).toHaveLength(0);
    });

    it("does nothing if the socket never joined", () => {
      const b = new Broadcaster();
      const a = makeWs();
      b.hello(asWs(a), "siteA", alice);
      expect(a.send).not.toHaveBeenCalled();
    });
  });

  describe("presence", () => {
    it("relays a caret move to the rest of the room, excluding the sender", () => {
      const b = new Broadcaster();
      const a = makeWs();
      const c = makeWs();
      b.join("doc", asWs(a));
      b.join("doc", asWs(c));
      b.hello(asWs(a), "siteA", alice);
      b.hello(asWs(c), "siteC", bob);
      a.send.mockClear();
      c.send.mockClear();

      const selection = { anchor: 3, focus: 5 };
      b.presence(asWs(a), selection);

      const toC = sentOfType(c, WS_MSG.Presence);
      expect(toC).toHaveLength(1);
      expect(toC[0]).toMatchObject({ siteId: "siteA", user: alice, selection });
      expect(sentOfType(a, WS_MSG.Presence)).toHaveLength(0);
    });

    it("ignores presence from a socket that never sent hello (no siteId)", () => {
      const b = new Broadcaster();
      const a = makeWs();
      const c = makeWs();
      b.join("doc", asWs(a));
      b.join("doc", asWs(c));
      c.send.mockClear();
      b.presence(asWs(a), { anchor: 1 });
      expect(c.send).not.toHaveBeenCalled();
    });
  });

  describe("publish", () => {
    it("sends an accepted change to the whole room including the sender", () => {
      const b = new Broadcaster();
      const a = makeWs();
      const c = makeWs();
      b.join("doc", asWs(a));
      b.join("doc", asWs(c));
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const change = { id: "ch1" } as any;

      b.publish("doc", change);

      expect(sentOfType(a, WS_MSG.Change)).toEqual([{ type: WS_MSG.Change, change }]);
      expect(sentOfType(c, WS_MSG.Change)).toEqual([{ type: WS_MSG.Change, change }]);
    });

    it("delivers nothing for an unknown room", () => {
      const b = new Broadcaster();
      const a = makeWs();
      b.join("doc", asWs(a));
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      b.publish("other-doc", { id: "x" } as any);
      expect(a.send).not.toHaveBeenCalled();
    });
  });

  describe("publishReview", () => {
    it("sends a review op to the room but excludes the sender", () => {
      const b = new Broadcaster();
      const a = makeWs();
      const c = makeWs();
      b.join("doc", asWs(a));
      b.join("doc", asWs(c));
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const env = { op: { type: "addThread" }, dependsOnSeq: 0 } as any;

      b.publishReview("doc", env, asWs(a));

      expect(sentOfType(a, WS_MSG.Review)).toHaveLength(0);
      expect(sentOfType(c, WS_MSG.Review)).toEqual([{ type: WS_MSG.Review, review: env }]);
    });
  });

  describe("close / leave", () => {
    it("broadcasts a Leave with the closer's siteId to the remaining room", () => {
      const b = new Broadcaster();
      const a = makeWs();
      const c = makeWs();
      b.join("doc", asWs(a));
      b.join("doc", asWs(c));
      b.hello(asWs(a), "siteA", alice);
      b.hello(asWs(c), "siteC", bob);
      a.send.mockClear();
      c.send.mockClear();

      a.close();

      const leave = sentOfType(c, WS_MSG.Leave);
      expect(leave).toEqual([{ type: WS_MSG.Leave, siteId: "siteA" }]);
    });

    it("does not send a Leave when the closing client never identified (no siteId)", () => {
      const b = new Broadcaster();
      const a = makeWs();
      const c = makeWs();
      b.join("doc", asWs(a));
      b.join("doc", asWs(c));
      c.send.mockClear();

      a.close();

      expect(sentOfType(c, WS_MSG.Leave)).toHaveLength(0);
    });
  });

  describe("debounced session end", () => {
    it("fires onSessionEnd after debounceMs once the room empties", () => {
      const onEnd = vi.fn();
      const b = new Broadcaster(onEnd, 100);
      const a = makeWs();
      b.join("doc", asWs(a));
      b.hello(asWs(a), "siteA", alice);

      a.close();
      expect(onEnd).not.toHaveBeenCalled();

      vi.advanceTimersByTime(99);
      expect(onEnd).not.toHaveBeenCalled();

      vi.advanceTimersByTime(1);
      expect(onEnd).toHaveBeenCalledTimes(1);
      expect(onEnd).toHaveBeenCalledWith("doc", { lastEditor: alice });
    });

    it("passes an empty ctx when the last connection had no identified user", () => {
      const onEnd = vi.fn();
      const b = new Broadcaster(onEnd, 50);
      const a = makeWs();
      b.join("doc", asWs(a));

      a.close();
      vi.advanceTimersByTime(50);

      expect(onEnd).toHaveBeenCalledWith("doc", {});
    });

    it("cancels the pending session end if someone rejoins before it elapses", () => {
      const onEnd = vi.fn();
      const b = new Broadcaster(onEnd, 100);
      const a = makeWs();
      b.join("doc", asWs(a));
      b.hello(asWs(a), "siteA", alice);

      a.close();
      vi.advanceTimersByTime(50);

      // Someone rejoins → the pending timer must be cleared.
      const c = makeWs();
      b.join("doc", asWs(c));

      vi.advanceTimersByTime(100);
      expect(onEnd).not.toHaveBeenCalled();
    });

    it("does not schedule anything when no onSessionEnd callback was provided", () => {
      const b = new Broadcaster(undefined, 10);
      const a = makeWs();
      b.join("doc", asWs(a));
      a.close();
      // Nothing to assert beyond "no throw"; advancing timers must be a no-op.
      expect(() => vi.advanceTimersByTime(50)).not.toThrow();
    });
  });
});
