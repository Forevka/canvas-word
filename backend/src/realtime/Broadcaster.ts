// Live change broadcast + presence. Holds the in-memory room state and the
// debounced "session ended" signal; transport-agnostic (works with any ws socket).
import type { WebSocket } from "ws";
import type { Change, ReviewOpEnvelope, UserInfo } from "@cw/shared";

interface Conn {
  docId: string;
  siteId?: string;
  user?: UserInfo | undefined;
  selection?: unknown;
}

export class Broadcaster {
  private readonly rooms = new Map<string, Set<WebSocket>>();
  private readonly meta = new Map<WebSocket, Conn>();
  // Pending session-end timers, keyed by doc — set when a room empties, cleared
  // if someone (re)joins before the debounce elapses.
  private readonly pendingEnd = new Map<string, ReturnType<typeof setTimeout>>();

  constructor(
    private readonly onSessionEnd?: (docId: string, ctx: { lastEditor?: UserInfo }) => void,
    private readonly debounceMs: number = Number(process.env.SESSION_END_DEBOUNCE_MS ?? 8000),
  ) {}

  join(docId: string, ws: WebSocket): void {
    // A (re)join means the editing session is still alive — cancel any pending
    // session-end webhook for this document.
    const pending = this.pendingEnd.get(docId);
    if (pending) {
      clearTimeout(pending);
      this.pendingEnd.delete(docId);
    }
    let room = this.rooms.get(docId);
    if (!room) this.rooms.set(docId, (room = new Set()));
    room.add(ws);
    this.meta.set(ws, { docId });
    ws.on("close", () => {
      room!.delete(ws);
      const m = this.meta.get(ws);
      this.meta.delete(ws);
      // Tell the room this collaborator left so their caret disappears.
      if (m?.siteId) this.toRoom(docId, { type: "leave", siteId: m.siteId }, ws);
      // The last editor just left → arm the debounced session-end signal.
      if (room!.size === 0) this.scheduleSessionEnd(docId, m);
    });
  }

  /** Arm (or re-arm) the debounced session-end for a now-empty room. The last
   *  connection's user is the best-effort "last editor"; the dispatcher falls
   *  back to the most recent author if it's absent. */
  private scheduleSessionEnd(docId: string, m: Conn | undefined): void {
    if (!this.onSessionEnd) return;
    const existing = this.pendingEnd.get(docId);
    if (existing) clearTimeout(existing);
    const lastEditor = m?.user;
    const timer = setTimeout(() => {
      this.pendingEnd.delete(docId);
      this.onSessionEnd!(docId, lastEditor ? { lastEditor } : {});
    }, this.debounceMs);
    this.pendingEnd.set(docId, timer);
  }

  /** A client identified itself: send it the current roster, announce it. */
  hello(ws: WebSocket, siteId: string, user?: UserInfo): void {
    const m = this.meta.get(ws);
    if (!m) return;
    m.siteId = siteId;
    m.user = user;
    const entries = this.roster(m.docId, ws);
    if (ws.readyState === ws.OPEN) ws.send(JSON.stringify({ type: "roster", entries }));
    this.toRoom(m.docId, { type: "presence", siteId, user, selection: m.selection ?? null }, ws);
  }

  /** A client's caret moved: relay to the rest of the room. */
  presence(ws: WebSocket, selection: unknown): void {
    const m = this.meta.get(ws);
    if (!m?.siteId) return;
    m.selection = selection;
    this.toRoom(m.docId, { type: "presence", siteId: m.siteId, user: m.user, selection }, ws);
  }

  /** Accepted change → everyone in the room (incl. sender, who treats it as ack). */
  publish(docId: string, change: Change): void {
    this.toRoom(docId, { type: "change", change });
  }

  /** A review op (suggestion/comment) → the rest of the room. Excludes the
   *  sender, who already applied it locally (apply is idempotent regardless). */
  publishReview(docId: string, env: ReviewOpEnvelope, sender: WebSocket): void {
    this.toRoom(docId, { type: "review", review: env }, sender);
  }

  private roster(
    docId: string,
    exclude: WebSocket,
  ): Array<{ siteId: string; user?: UserInfo | undefined; selection: unknown }> {
    const room = this.rooms.get(docId);
    const out: Array<{ siteId: string; user?: UserInfo | undefined; selection: unknown }> = [];
    if (!room) return out;
    for (const ws of room) {
      if (ws === exclude) continue;
      const m = this.meta.get(ws);
      if (m?.siteId) out.push({ siteId: m.siteId, user: m.user, selection: m.selection ?? null });
    }
    return out;
  }

  private toRoom(docId: string, msg: unknown, exclude?: WebSocket): void {
    const room = this.rooms.get(docId);
    if (!room) return;
    const s = JSON.stringify(msg);
    for (const ws of room) {
      if (ws === exclude) continue;
      if (ws.readyState === ws.OPEN) ws.send(s);
    }
  }
}
