// Review layer — track changes ("suggestions") + comments, modeled as an
// ATTRIBUTED, ANCHORED OVERLAY on top of the core document. This module is the
// data model only; it is deliberately a SIBLING of shared/src/model — it imports
// a handful of core *types* (DocPosition/CharStyle/UserInfo/Run) but never the
// core ops, and the core never imports this. The docx import/export pipeline is
// entirely unaware of it. See REVIEW.md for the full design rationale.
//
// The invariant the whole feature relies on: the live core document = "all
// insertions kept, all deletions still present". A tracked deletion never
// removes text until accepted — it is merely painted struck-through. Accept and
// reject are then pure bookkeeping that emit ordinary core ops.

import type { DocPosition } from "../model/position";
import type { CharStyle, Run } from "../model/document";
import type { Op } from "../model/ops";
import type { UserInfo } from "../change";

/** Rich-text comment body: the same style-homogeneous run list the document
 *  uses, so comment composing/rendering reuses fragmentToHtml/htmlToFragment. */
export type Fragment = Run[];

/** A range into the LIVE document. start/end are (blockId, offset) positions —
 *  identical shape to BookmarkRange — and ride mapPosition on every edit, so the
 *  range tracks its text through inserts/deletes/splits. start === end is a point
 *  (a caret-anchored comment, or a collapsed/dead range awaiting GC). May span
 *  blocks. */
export interface ReviewAnchor {
  start: DocPosition;
  end: DocPosition;
}

export type SuggestionKind = "insert" | "delete" | "format" | "structural";

/** A tracked structural edit (split/merge/block/table op). The forward `op` was
 *  already applied to the live core doc when the suggestion was recorded; reject
 *  applies `inverse` to undo it. Both are captured at intercept time (the inverse
 *  is harvested from applyOp's exact inverse), so resolve never reconstructs it.
 *  `blockId` is the doc block this record hangs on — when that block no longer
 *  exists the record is dead (GC drops it), mirroring how a collapsed text anchor
 *  signals a dead insert/delete record. */
export interface StructuralChange {
  op: Op;
  inverse: Op;
  /** Block whose disappearance kills the record (created/merged/removed block,
   *  or the host table for table ops). */
  blockId: string;
}

export interface Suggestion {
  /** Client-minted, globally unique (freshId). */
  id: string;
  kind: SuggestionKind;
  /** Covers text PRESENT in the doc (see the module invariant). */
  anchor: ReviewAnchor;
  /** Attribution. The embedder owns identity; the layer just presents it. */
  author: UserInfo;
  /** Wall-clock ms; display/sort only, never ordering. */
  createdAt: number;
  /** format only: the char-style patch this suggestion applied, plus its exact
   *  inverse so reject can undo it (captured via applyStylePatchToRuns). */
  patch?: Partial<CharStyle>;
  inverse?: Partial<CharStyle>;
  /** structural only: the applied core op + its inverse + the block the record
   *  hangs on. The `anchor` of a structural record is a degenerate point on its
   *  block (there is no text range to highlight) — paint draws a coarse
   *  block-level change-bar instead. */
  structural?: StructuralChange;
  /** Optional grouping: one continuous authoring action (typing a word) is one
   *  suggestion that grows; groupId lets the reviewer accept/reject related
   *  records as a unit. */
  groupId?: string;
}

export interface Comment {
  id: string;
  author: UserInfo;
  body: Fragment;
  createdAt: number;
  editedAt?: number;
  /** Users @-mentioned in this comment (structured, for notifications). The body
   *  text carries the "@Display Name" spans; this is the resolved identity list,
   *  drawn from the editor's `knownUsers`. */
  mentions?: UserInfo[];
}

export type ThreadStatus = "open" | "resolved";

export interface CommentThread {
  id: string;
  /** Highlighted span; start === end ⇒ point comment. */
  anchor: ReviewAnchor;
  status: ThreadStatus;
  /** Ordered; comments[0] is the root, the rest are replies. */
  comments: Comment[];
}

export interface ReviewLayer {
  docId: string;
  /** The core document version (recorder/SyncClient seq) these anchors were last
   *  rebased to. A late joiner whose doc is newer fast-forwards anchors by
   *  replaying core ops baseVersion..head through mapPosition. */
  baseVersion: number;
  suggestions: Suggestion[];
  threads: CommentThread[];
}

export function emptyReview(docId: string, baseVersion = 0): ReviewLayer {
  return { docId, baseVersion, suggestions: [], threads: [] };
}

/** Whether an anchor has collapsed to a point (same block + offset). A point
 *  comment is legitimately collapsed; a collapsed *suggestion* means its text
 *  was removed and the record is dead (GC drops it). */
export function isCollapsedAnchor(a: ReviewAnchor): boolean {
  return a.start.blockId === a.end.blockId && a.start.offset === a.end.offset;
}
