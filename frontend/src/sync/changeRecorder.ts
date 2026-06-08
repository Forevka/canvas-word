// Captures committed edits as Changes for the document history. The editor calls
// record() from its single commit chokepoint (and from undo/redo, which apply
// forward ops directly). Each Change carries forward ops only; the recorder
// keeps an in-memory log and forwards to an optional sink (the SyncClient, once
// the backend lands in Phase 3/4).

import { currentSiteId, freshId, type Change, type ChangeOrigin, type DocSelection, type Op } from "@cw/shared";

export type ChangeSink = (change: Change) => void;

export class ChangeRecorder {
  private readonly log: Change[] = [];
  // Tentative-local version counter. The server reassigns the canonical seq once
  // a backend exists; until then this IS the order (single writer).
  private version = 0;

  constructor(
    private readonly docId: string,
    private readonly sink?: ChangeSink,
  ) {}

  /** Record an applied op group. `ops` are the forward ops; the inverse is
   *  derivable via applyOp, so it is not stored. Returns the Change. */
  record(ops: Op[], origin: ChangeOrigin, selectionAfter: DocSelection | null, ts: number): Change | null {
    if (ops.length === 0) return null;
    const change: Change = {
      id: freshId(),
      docId: this.docId,
      baseVersion: this.version,
      seq: this.version,
      siteId: currentSiteId(),
      origin,
      ts,
      ops,
      selectionAfter,
    };
    this.version++;
    this.log.push(change);
    this.sink?.(change);
    return change;
  }

  /** The full ordered change log (since the base snapshot). */
  changes(): Change[] {
    return this.log.slice();
  }

  /** The number of changes recorded — the current local version. */
  head(): number {
    return this.version;
  }
}
