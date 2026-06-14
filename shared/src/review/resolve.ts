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

import { applyStylePatchToRuns, type Op } from "../model/ops";
import type { Document } from "../model/document";
import { blockById, blockIndexOf } from "../model/text";
import type { ReviewLayer, Suggestion } from "./model";
import type { ReviewOp } from "./ops";

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
  return s.kind === "delete" ? { ops: deleteAnchorOp(s), reviewOps: [drop(id)] } : { ops: [], reviewOps: [drop(id)] };
}

export function rejectSuggestion(doc: Document, review: ReviewLayer, id: string): Resolution | null {
  const s = review.suggestions.find((x) => x.id === id);
  if (!s) return null;
  if (s.kind === "insert") return { ops: deleteAnchorOp(s), reviewOps: [drop(id)] };
  if (s.kind === "format") return { ops: formatRejectOp(doc, s), reviewOps: [drop(id)] };
  return { ops: [], reviewOps: [drop(id)] }; // delete → keep text, drop record
}

/** Document-order key for back-to-front folding (later blocks/offsets first). */
function orderKey(doc: Document, s: Suggestion): number {
  const bi = blockIndexOf(doc, s.anchor.start.blockId);
  return (bi < 0 ? 1e9 : bi) * 1e7 + s.anchor.start.offset;
}

function resolveAll(
  doc: Document,
  review: ReviewLayer,
  one: (doc: Document, review: ReviewLayer, id: string) => Resolution | null,
): Resolution {
  const ordered = review.suggestions.slice().sort((a, b) => orderKey(doc, b) - orderKey(doc, a));
  const ops: Op[] = [];
  const reviewOps: ReviewOp[] = [];
  for (const s of ordered) {
    const r = one(doc, review, s.id);
    if (!r) continue;
    ops.push(...r.ops); // back-to-front → earlier anchors stay valid
    reviewOps.push(...r.reviewOps);
  }
  return { ops, reviewOps };
}

export const acceptAllSuggestions = (doc: Document, review: ReviewLayer): Resolution =>
  resolveAll(doc, review, acceptSuggestion);
export const rejectAllSuggestions = (doc: Document, review: ReviewLayer): Resolution =>
  resolveAll(doc, review, rejectSuggestion);
