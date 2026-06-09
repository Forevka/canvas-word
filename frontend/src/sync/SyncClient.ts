// SyncClient — live collaboration over WebSocket (Jupiter / ShareDB model).
//
// The server holds the canonical total order and assigns each change a seq. The
// client applies its own edits optimistically, then keeps converging with the
// server:
//
//   - One change is in flight (sent, awaiting ack) at a time; later local ops
//     queue in `buffer` and are sent once the in-flight one is acked.
//   - When a REMOTE change arrives, the client transforms it against its own
//     un-acked ops (inflight + buffer) and applies the result to the editor,
//     and rebases inflight + buffer against the remote.
//
// It also relays PRESENCE (ephemeral, never persisted): a "hello" with the local
// user on connect, the local caret on every move (throttled), and inbound
// roster/presence/leave → the editor (which owns + rebases peer caret positions)
// plus userEntered/userLeave/presence callbacks for the embedder.

import { currentSiteId, freshId, transformOps, type Change, type DocSelection, type Op, type UserInfo } from "@cw/shared";

export interface SyncEditor {
  applyRemoteOps(ops: Op[]): void;
  /** Upsert a peer's caret (the editor rebases it through edits + renders it). */
  setPeerPresence(siteId: string, user: UserInfo | undefined, selection: DocSelection | null): void;
  /** Remove a peer's caret (they left). */
  removePeer(siteId: string): void;
}

export interface SyncClientOptions {
  /** ws:// base, e.g. ws://localhost:8787 */
  wsUrl: string;
  docId: string;
  editor: SyncEditor;
  /** Server version (head) the initial document was loaded at. */
  startVersion: number;
  /** Local user identity (attribution + presence). */
  user?: UserInfo | undefined;
  /** Notified on every applied remote change (e.g. activity panel). */
  onRemote?: ((change: Change) => void) | undefined;
  onUserEntered?: ((siteId: string, user: UserInfo | undefined) => void) | undefined;
  onUserLeave?: ((siteId: string, user: UserInfo | undefined) => void) | undefined;
  onPresence?: ((participants: Array<{ siteId: string; user?: UserInfo | undefined }>) => void) | undefined;
}

interface PresenceMsg {
  type: string;
  change?: Change;
  siteId?: string;
  user?: UserInfo;
  selection?: DocSelection | null;
  entries?: Array<{ siteId: string; user?: UserInfo; selection?: DocSelection | null }>;
}

const PRESENCE_THROTTLE_MS = 80;

export class SyncClient {
  private ws: WebSocket | null = null;
  private version: number;
  private inflight: { id: string; ops: Op[] } | null = null;
  private buffer: Op[] = [];
  private readonly opts: SyncClientOptions;
  // Known peers (for entered/left dedup + the participants list).
  private readonly peers = new Map<string, UserInfo | undefined>();
  private presenceTimer: ReturnType<typeof setTimeout> | null = null;
  private lastSelection: DocSelection | null = null;

  constructor(opts: SyncClientOptions) {
    this.opts = opts;
    this.version = opts.startVersion;
  }

  connect(): void {
    const ws = new WebSocket(`${this.opts.wsUrl}/ws?doc=${encodeURIComponent(this.opts.docId)}`);
    this.ws = ws;
    ws.onopen = () => {
      ws.send(JSON.stringify({ type: "hello", siteId: currentSiteId(), user: this.opts.user }));
      this.flush();
    };
    ws.onmessage = (ev) => {
      const msg = JSON.parse(ev.data as string) as PresenceMsg;
      switch (msg.type) {
        case "change":
          if (msg.change) this.onServerChange(msg.change);
          break;
        case "roster":
          for (const e of msg.entries ?? []) this.applyPeer(e.siteId, e.user, e.selection ?? null);
          break;
        case "presence":
          if (msg.siteId) this.applyPeer(msg.siteId, msg.user, msg.selection ?? null);
          break;
        case "leave":
          if (msg.siteId) this.removePeer(msg.siteId);
          break;
      }
    };
    ws.onclose = () => {
      this.ws = null;
    };
  }

  /** Feed local edits (the editor's recorded ops) into the sync pipeline. */
  localEdit(ops: Op[]): void {
    if (ops.length === 0) return;
    this.buffer.push(...ops);
    this.flush();
  }

  /** Broadcast the local caret (throttled). */
  localPresence(selection: DocSelection | null): void {
    this.lastSelection = selection;
    if (this.presenceTimer) return;
    this.presenceTimer = setTimeout(() => {
      this.presenceTimer = null;
      if (this.ws && this.ws.readyState === WebSocket.OPEN) {
        this.ws.send(
          JSON.stringify({ type: "presence", siteId: currentSiteId(), user: this.opts.user, selection: this.lastSelection }),
        );
      }
    }, PRESENCE_THROTTLE_MS);
  }

  getVersion(): number {
    return this.version;
  }

  private flush(): void {
    if (this.inflight || this.buffer.length === 0) return;
    if (!this.ws || this.ws.readyState !== WebSocket.OPEN) return;
    const id = freshId();
    const ops = this.buffer;
    this.buffer = [];
    this.inflight = { id, ops };
    const change: Change = {
      id,
      docId: this.opts.docId,
      baseVersion: this.version,
      siteId: currentSiteId(),
      ...(this.opts.user ? { userId: this.opts.user.id } : {}),
      origin: "command",
      ts: Date.now(),
      ops,
    };
    this.ws.send(JSON.stringify({ type: "submit", change }));
  }

  private onServerChange(change: Change): void {
    // Ack of our own in-flight change: it's confirmed. No re-apply — our
    // optimistic local state already matches (server transformed it the same
    // way we rebased). Advance and send the next buffered change.
    if (this.inflight && change.id === this.inflight.id) {
      this.version = (change.seq ?? this.version) + 1;
      this.inflight = null;
      this.flush();
      return;
    }

    // A remote change. Apply it on top of our un-acked ops, and rebase those ops.
    const pending = [...(this.inflight?.ops ?? []), ...this.buffer];
    const remoteForLocal = transformOps(change.ops, pending, "left"); // remote keeps priority
    this.opts.editor.applyRemoteOps(remoteForLocal);

    if (this.inflight) {
      this.inflight = { id: this.inflight.id, ops: transformOps(this.inflight.ops, change.ops, "right") };
    }
    this.buffer = transformOps(this.buffer, change.ops, "right");
    this.version = (change.seq ?? this.version) + 1;
    this.opts.onRemote?.(change);
  }

  private applyPeer(siteId: string, user: UserInfo | undefined, selection: DocSelection | null): void {
    if (siteId === currentSiteId()) return; // never show our own caret
    const isNew = !this.peers.has(siteId);
    this.peers.set(siteId, user);
    this.opts.editor.setPeerPresence(siteId, user, selection);
    if (isNew) this.opts.onUserEntered?.(siteId, user);
    this.emitPresence();
  }

  private removePeer(siteId: string): void {
    const user = this.peers.get(siteId);
    if (!this.peers.delete(siteId)) return;
    this.opts.editor.removePeer(siteId);
    this.opts.onUserLeave?.(siteId, user);
    this.emitPresence();
  }

  private emitPresence(): void {
    this.opts.onPresence?.([...this.peers.entries()].map(([siteId, user]) => ({ siteId, user })));
  }

  destroy(): void {
    if (this.presenceTimer) clearTimeout(this.presenceTimer);
    this.ws?.close();
    this.ws = null;
  }
}
