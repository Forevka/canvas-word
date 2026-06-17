// Accept / reject — turn a reviewer action into a normal core edit + a review-op.
// Pure + isomorphic (reused by the editor runtime AND the backend export bake).
// See REVIEW.md §1 table and §5.5.
//
//   insert  accept→keep (drop record)         reject→deleteRange + drop
//   delete  accept→deleteRange + drop          reject→keep (drop record)
//   format  accept→keep (drop record)          reject→re-apply inverse + drop
//
// acceptAll/rejectAll fold back-to-front in document order so earlier anchors
// stay valid as later ranges collapse — one Transaction, one undo step.

import { applyOp, applyStylePatchToRuns, containerOf, type Op } from "../model/ops";
import type { Document } from "../model/document";
import { blockById, blockIndexOf } from "../model/text";
import type { ReviewLayer, Suggestion } from "./model";
import { applyReviewOp, type ReviewOp } from "./ops";
import { rebaseReview } from "./rebase";

export interface Resolution {
  ops: Op[];
  reviewOps: ReviewOp[];
}

/** A real deletion of a suggestion's anchored text. Single-block anchors only
 *  (cross-block deletion suggestions are V2 — they resolve to no destructive op,
 *  just the record drop, leaving the text). */
function deleteAnchorOp(s: Suggestion): Op[] {
  const { start, end } = s.anchor;
  if (start.blockId !== end.blockId || start.offset >= end.offset) return [];
  return [{ type: "deleteRange", blockId: start.blockId, start: start.offset, end: end.offset }];
}

/** Re-apply a format suggestion's inverse patch over its anchor (reject). */
function formatRejectOp(doc: Document, s: Suggestion): Op[] {
  const { start, end } = s.anchor;
  if (!s.inverse || start.blockId !== end.blockId || start.offset >= end.offset) return [];
  const block = blockById(doc, start.blockId);
  if (!block) return [];
  return [{ type: "setRuns", blockId: block.id, runs: applyStylePatchToRuns(block.runs, start.offset, end.offset, s.inverse) }];
}

const drop = (id: string): ReviewOp => ({ type: "removeSuggestion", id });

export function acceptSuggestion(doc: Document, review: ReviewLayer, id: string): Resolution | null {
  const s = review.suggestions.find((x) => x.id === id);
  if (!s) return null;
  // structural + format + insert are already applied to the live doc → accept is
  // a pure record drop. Only a tracked DELETE still needs its destructive op.
  return s.kind === "delete" ? { ops: deleteAnchorOp(s), reviewOps: [drop(id)] } : { ops: [], reviewOps: [drop(id)] };
}

export function rejectSuggestion(doc: Document, review: ReviewLayer, id: string): Resolution | null {
  const s = review.suggestions.find((x) => x.id === id);
  if (!s) return null;
  if (s.kind === "insert") return { ops: deleteAnchorOp(s), reviewOps: [drop(id)] };
  if (s.kind === "format") return { ops: formatRejectOp(doc, s), reviewOps: [drop(id)] };
  // structural reject re-applies the exact inverse captured at intercept time.
  if (s.kind === "structural") return { ops: s.structural ? [s.structural.inverse] : [], reviewOps: [drop(id)] };
  return { ops: [], reviewOps: [drop(id)] }; // delete → keep text, drop record
}

/** Document-order key for back-to-front folding of TEXT records (later
 *  blocks/offsets first). Structural records are NOT ordered this way — their
 *  anchored block moved/changed shape, so a positional key is meaningless; they
 *  fold by creation order instead (see resolveAll). */
function orderKey(doc: Document, s: Suggestion): number {
  let bi = blockIndexOf(doc, s.anchor.start.blockId);
  if (bi < 0) {
    const top = containerOf(doc, s.anchor.start.blockId);
    bi = top ? top.index : 1e5;
  }
  return bi * 1e7 + s.anchor.start.offset;
}

function resolveAll(
  doc: Document,
  review: ReviewLayer,
  one: (doc: Document, review: ReviewLayer, id: string) => Resolution | null,
): Resolution {
  // STRUCTURAL records first, newest-first: each structural op was applied on the
  // post-state of the previous (a paste = split then insert), so its inverse must
  // unwind in reverse creation order — exactly like undo. Doing them before the
  // text records keeps block indices/identities the text anchors reference valid
  // (a structural reject can re-add or remove whole blocks).
  const structural = review.suggestions.filter((s) => s.kind === "structural").sort((a, b) => b.createdAt - a.createdAt);
  // TEXT records back-to-front in document order so earlier ranges stay valid as
  // later ones collapse.
  const text = review.suggestions.filter((s) => s.kind !== "structural").sort((a, b) => orderKey(doc, b) - orderKey(doc, a));
  const order = [...structural, ...text].map((s) => s.id);

  // Resolve against a RUNNING doc + review: each resolution's core ops are applied
  // to the running doc and the remaining record anchors are rebased through them
  // (mirroring the commit pipeline). This is what makes a structural reject — which
  // re-merges/re-inserts whole blocks — compose with the text rejects whose anchors
  // live on those blocks (e.g. rejecting a multi-block paste). One Transaction is
  // still emitted: the ops accumulate in resolution order.
  let runDoc = doc;
  let runReview = review;
  const ops: Op[] = [];
  const reviewOps: ReviewOp[] = [];
  for (const id of order) {
    const r = one(runDoc, runReview, id);
    if (!r) continue;
    for (const op of r.ops) {
      const res = applyOp(runDoc, op);
      runDoc = res.doc;
      runReview = rebaseReview(runReview, res.mapPosition); // remaining anchors travel
      ops.push(op);
    }
    for (const rop of r.reviewOps) runReview = applyReviewOp(runReview, rop).layer;
    reviewOps.push(...r.reviewOps);
  }
  return { ops, reviewOps };
}

export const acceptAllSuggestions = (doc: Document, review: ReviewLayer): Resolution =>
  resolveAll(doc, review, acceptSuggestion);
export const rejectAllSuggestions = (doc: Document, review: ReviewLayer): Resolution =>
  resolveAll(doc, review, rejectSuggestion);
