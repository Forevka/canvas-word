// Export bake — fold all suggestions into a clean Document on a COPY, so the
// existing docx/PDF writers (which are review-blind) emit a plain file. Isomorphic
// (used by the editor's exportDocx AND the backend export route). Default export
// bakes "reject" (the original baseline: insertions removed, deletions kept);
// "accept" yields the final text. Comments aren't written by the core writers, so
// they need no baking.

import { applyOp } from "../model/ops";
import type { Document } from "../model/document";
import type { ReviewLayer } from "./model";
import { acceptAllSuggestions, rejectAllSuggestions } from "./resolve";

export type BakeMode = "accept" | "reject";

/** Resolve every suggestion against `doc` and return the resulting Document.
 *  reject (default) = remove insertions, keep deletions → the original baseline.
 *  accept = keep insertions, remove deletions → the final text. */
export function bakeReview(doc: Document, review: ReviewLayer, mode: BakeMode = "reject"): Document {
  if (review.suggestions.length === 0) return doc;
  const res = mode === "accept" ? acceptAllSuggestions(doc, review) : rejectAllSuggestions(doc, review);
  // resolveAll emits ops back-to-front in document order, so applying them in
  // sequence keeps every op's coordinates valid (later edits don't shift earlier).
  let d = doc;
  for (const op of res.ops) d = applyOp(d, op).doc;
  return d;
}
