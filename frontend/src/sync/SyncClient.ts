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
// The server independently transforms an in-flight change against any changes
// committed since its baseVersion — using the SAME transform + side — so the
// server's accepted ops equal the client's locally-rebased ops (TP1). Hence an
// ack needs no re-apply: optimistic local state already matches the server.

import { currentSiteId, freshId, transformOps, type Change, type Op } from "@cw/shared";

export interface SyncEditor {
  applyRemoteOps(ops: Op[]): void;
}

export interface SyncClientOptions {
  /** ws:// base, e.g. ws://localhost:8787 */
  wsUrl: string;
  docId: string;
  editor: SyncEditor;
  /** Server version (head) the initial document was loaded at. */
  startVersion: number;
  /** Optional: notified on every applied remote change (UI/presence). */
  onRemote?: (change: Change) => void;
}

export class SyncClient {
  private ws: WebSocket | null = null;
  private version: number;
  private inflight: { id: string; ops: Op[] } | null = null;
  private buffer: Op[] = [];
  private readonly opts: SyncClientOptions;

  constructor(opts: SyncClientOptions) {
    this.opts = opts;
    this.version = opts.startVersion;
  }

  connect(): void {
    const ws = new WebSocket(`${this.opts.wsUrl}/ws?doc=${encodeURIComponent(this.opts.docId)}`);
    this.ws = ws;
    ws.onopen = () => this.flush();
    ws.onmessage = (ev) => {
      const msg = JSON.parse(ev.data as string) as { type: string; change?: Change };
      if (msg.type === "change" && msg.change) this.onServerChange(msg.change);
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

  /** Current confirmed server version. */
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

  destroy(): void {
    this.ws?.close();
    this.ws = null;
  }
}
