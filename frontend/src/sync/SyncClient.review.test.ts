import { afterEach, describe, expect, it, vi } from "vitest";
import { SyncClient, type SyncEditor } from "./SyncClient";
import type { ReviewOp } from "@cw/shared";

// Minimal controllable WebSocket stub (jsdom/node have no WebSocket).
class FakeWS {
  static OPEN = 1;
  static last: FakeWS | null = null;
  readyState = 1;
  onopen: (() => void) | null = null;
  onmessage: ((ev: { data: string }) => void) | null = null;
  onclose: (() => void) | null = null;
  sent: string[] = [];
  constructor(public url: string) {
    FakeWS.last = this;
  }
  send(d: string): void {
    this.sent.push(d);
  }
  close(): void {
    this.readyState = 3;
    this.onclose?.();
  }
  open(): void {
    this.onopen?.();
  }
  deliver(obj: unknown): void {
    this.onmessage?.({ data: JSON.stringify(obj) });
  }
}

vi.stubGlobal("WebSocket", FakeWS);
afterEach(() => (FakeWS.last = null));

const insertOp = (id: string): ReviewOp => ({
  type: "addSuggestion",
  s: { id, kind: "insert", anchor: { start: { blockId: "p", offset: 0 }, end: { blockId: "p", offset: 3 } }, author: { id: "u", firstName: "A", lastName: "B" }, createdAt: 1 },
});

const makeClient = () => {
  const applied: ReviewOp[] = [];
  const editor: SyncEditor = {
    applyRemoteOps: () => {},
    applyRemoteReviewOp: (op) => applied.push(op),
    setPeerPresence: () => {},
    removePeer: () => {},
  };
  const s = new SyncClient({ wsUrl: "ws://x", docId: "d", editor, startVersion: 0 });
  s.connect();
  return { s, applied, ws: FakeWS.last! };
};

const reviewMsg = (op: ReviewOp, dependsOnSeq = 0) => ({ type: "review", review: { op, dependsOnSeq } });

describe("SyncClient — join-window review hold", () => {
  it("holds review ops until markReviewReady, then replays them", () => {
    const { s, applied, ws } = makeClient();
    s.holdReview();
    ws.open();
    ws.deliver(reviewMsg(insertOp("s1")));
    ws.deliver(reviewMsg(insertOp("s2")));
    expect(applied).toHaveLength(0); // held during the seed window
    s.markReviewReady();
    expect(applied.map((o) => (o.type === "addSuggestion" ? o.s.id : ""))).toEqual(["s1", "s2"]);
  });

  it("applies review ops immediately when not holding (freshly shared doc)", () => {
    const { applied, ws } = makeClient();
    ws.open();
    ws.deliver(reviewMsg(insertOp("s1")));
    expect(applied).toHaveLength(1);
  });

  it("whenOpen resolves once the socket opens", async () => {
    const { s, ws } = makeClient();
    let opened = false;
    const p = s.whenOpen().then(() => (opened = true));
    expect(opened).toBe(false);
    ws.open();
    await p;
    expect(opened).toBe(true);
  });
});
