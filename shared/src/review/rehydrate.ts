// Late-join rehydration: fold the persisted review-op log into a ReviewLayer
// whose anchors are fast-forwarded to the document's HEAD version. Each stored
// op carries the core version its anchor was authored against (dependsOnSeq);
// replaying the change log gives a position-mapper per change, and an op's
// anchor is mapped through the changes applied after its dependsOnSeq. So a user
// opening a shared document sees every suggestion/comment correctly placed
// against the current text. REVIEW.md §6.

import { applyOp } from "../model/ops";
import type { DocPosition } from "../model/position";
import { deserializeDocument, type SerializedDocument } from "../persist/serialize";
import type { Change } from "../change";
import { emptyReview, type ReviewAnchor, type ReviewLayer } from "./model";
import { applyReviewOp, type ReviewOp, type ReviewOpEnvelope } from "./ops";
import { gcReviewLayer } from "./rebase";

type Mapper = (p: DocPosition) => DocPosition;

const mapAnchor = (a: ReviewAnchor, m: Mapper): ReviewAnchor => ({ start: m(a.start), end: m(a.end) });

/** Map the anchors embedded in a single review op. Ops with no positional
 *  payload (removeSuggestion, addComment, setThreadStatus, …) pass through. */
function mapReviewOpAnchors(op: ReviewOp, m: Mapper): ReviewOp {
  switch (op.type) {
    case "addSuggestion":
      return { ...op, s: { ...op.s, anchor: mapAnchor(op.s.anchor, m) } };
    case "addThread":
      return { ...op, t: { ...op.t, anchor: mapAnchor(op.t.anchor, m) } };
    case "growSuggestion":
      return { ...op, end: m(op.end) };
    default:
      return op;
  }
}

export function rehydrateReview(
  docId: string,
  base: SerializedDocument,
  baseVersion: number,
  changes: readonly Change[],
  reviewOps: readonly ReviewOpEnvelope[],
): ReviewLayer {
  // One composed position-mapper per change, in seq order (index i = the change
  // at seq baseVersion + i). Built by replaying the log over the base snapshot —
  // the same fold reconstruct() does, but capturing each op's mapPosition.
  let doc = deserializeDocument(base);
  const mappers: Mapper[] = [];
  for (const ch of changes) {
    const ms: Mapper[] = [];
    for (const op of ch.ops) {
      const res = applyOp(doc, op);
      doc = res.doc;
      ms.push(res.mapPosition);
    }
    mappers.push((p) => ms.reduce((acc, mm) => mm(acc), p));
  }
  const head = baseVersion + changes.length;

  // Map a position from version D forward to head (apply changes D..head-1).
  const forward = (D: number): Mapper => {
    const from = Math.max(0, Math.min(D, head) - baseVersion);
    const slice = mappers.slice(from);
    return (p) => slice.reduce((acc, mm) => mm(acc), p);
  };

  let layer = emptyReview(docId, head);
  for (const env of reviewOps) {
    layer = applyReviewOp(layer, mapReviewOpAnchors(env.op, forward(env.dependsOnSeq))).layer;
  }
  // Drop suggestions whose text was removed by later edits (collapsed anchors).
  return gcReviewLayer(layer);
}
