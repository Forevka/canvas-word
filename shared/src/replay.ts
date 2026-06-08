// Replay: reconstruct a document version by folding the change log over a base
// snapshot. This is the heart of the "base + ordered changes" model — the server
// hands a client one base snapshot plus the ordered Change log, and the client
// (or the server, for checkpointing) replays it here. Pure reuse of applyOp; no
// new mutation logic exists.

import { changeOps, orderBySeq, type Change } from "./change";
import type { Document } from "./model/document";
import { applyOp, type Op } from "./model/ops";
import { deserializeDocument, type SerializedDocument } from "./persist/serialize";

/** Fold a flat op stream onto a document (in the given order). Each op must be
 *  valid against the running state — replaying the right log onto the right base
 *  never throws; an out-of-order or wrong-base log will. */
export function applyOps(doc: Document, ops: readonly Op[]): Document {
  let d = doc;
  for (const op of ops) d = applyOp(d, op).doc;
  return d;
}

/** Apply a change set (ordered by seq) onto a document. */
export function applyChanges(doc: Document, changes: readonly Change[]): Document {
  return applyOps(doc, changeOps(changes));
}

/** Reconstruct a document version: deserialize the base snapshot, then fold the
 *  change log over it in seq order. `uptoSeq` (inclusive) yields a historical
 *  version for time-travel; omit it for the latest. */
export function reconstruct(
  base: SerializedDocument,
  changes: readonly Change[],
  uptoSeq?: number,
): Document {
  const ordered = orderBySeq(changes).filter(
    (c) => uptoSeq === undefined || (c.seq ?? Number.MAX_SAFE_INTEGER) <= uptoSeq,
  );
  return applyChanges(deserializeDocument(base), ordered);
}
