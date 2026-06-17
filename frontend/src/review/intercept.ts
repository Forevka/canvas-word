// Suggestion-mode interceptor. A PURE transform sitting at the commit boundary:
// when the editor is in "suggest" mode, it rewrites a Transaction so destructive
// edits become non-destructive overlay records. When off it is never called
// (the runtime passes the transaction straight through).
//
//   intercept(trn, review, doc, author, now) -> { core, reviewOps }
//
// `core` is a normal Transaction the runtime commits exactly as today (undoable,
// recorded, synced). `reviewOps` mutate the review layer and are applied +
// broadcast AFTER the core ops (so their anchors are in post-core coordinates).
//
// The non-obvious lever: the runtime's rebaseReviewLayer already shifts existing
// anchors through each applied core op (the rebaseBookmarks path). So:
//  - typing at the END of (or INSIDE) the author's own pending insertion auto-
//    grows that record via mapPosition — no growSuggestion op needed; we just
//    skip creating a duplicate.
//  - really deleting text that lies inside the author's own insertion auto-
//    shrinks that record (and GC drops it if it collapses) — we just emit the
//    core delete and add nothing.
//
// Tracked (see REVIEW.md §5.2): text insert/delete, style-only setRuns, the
// structural ops (split/merge/block/table row+column), AND multi-op transactions
// — paste and cross-paragraph deletes are decomposed per-op through the same
// logic and bundled under one groupId. Only non-tracked op kinds (e.g.
// setParaStyle, table-structure, props) and ops on missing blocks pass through.

import {
  applyOp,
  applyReviewOp,
  blockById,
  rebaseReview,
  textOfRuns,
  type CharStyle,
  type Document,
  type Op,
  type Run,
  type ReviewOp,
  type ReviewLayer,
  type StructuralChange,
  type Suggestion,
  type UserInfo,
  freshId,
} from "@cw/shared";
import type { Transaction } from "../editor/state";

export interface InterceptResult {
  core: Transaction;
  reviewOps: ReviewOp[];
}

const pos = (blockId: string, offset: number) => ({ blockId, offset });

/** Author's insertion-suggestion intervals in one block (current coordinates). */
function authorInsertsIn(review: ReviewLayer, blockId: string, authorId: string): Array<{ start: number; end: number }> {
  return review.suggestions
    .filter(
      (s) =>
        s.kind === "insert" &&
        s.author.id === authorId &&
        s.anchor.start.blockId === blockId &&
        s.anchor.end.blockId === blockId,
    )
    .map((s) => ({ start: s.anchor.start.offset, end: s.anchor.end.offset }));
}

/** Is [s,e) fully covered by the union of the given intervals? */
function fullyCovered(s: number, e: number, intervals: Array<{ start: number; end: number }>): boolean {
  if (s >= e) return true;
  // Sweep: sort, merge, check a single merged interval contains [s,e).
  const sorted = intervals.slice().sort((a, b) => a.start - b.start);
  let cursor = s;
  for (const iv of sorted) {
    if (iv.start > cursor) break; // gap before cursor → not covered
    if (iv.end > cursor) cursor = iv.end;
    if (cursor >= e) return true;
  }
  return cursor >= e;
}

/** Does the insert offset extend an existing author insertion (so rebase will
 *  grow it and we must NOT add a duplicate record)? True when O is inside
 *  (start, end] — at-end (append, the typing case) or strictly inside. */
function extendsAuthorInsert(review: ReviewLayer, blockId: string, offset: number, authorId: string): boolean {
  return authorInsertsIn(review, blockId, authorId).some((iv) => offset > iv.start && offset <= iv.end);
}

/** Per-character style of run list at UTF-16 index i (the run containing i). */
function styleAtChar(runs: Run[], i: number): CharStyle | undefined {
  let cum = 0;
  for (const r of runs) {
    const end = cum + r.text.length;
    if (i >= cum && i < end) return r.style;
    cum = end;
  }
  return runs[runs.length - 1]?.style;
}

/** The keys whose values differ between two styles, with old (inverse) + new
 *  (patch) values. Approximate for mixed-style ranges (uses the style at the
 *  range start) — exact for the common uniform-format case. */
function styleDelta(oldS: CharStyle, newS: CharStyle): { patch: Partial<CharStyle>; inverse: Partial<CharStyle> } {
  const patch: Partial<CharStyle> = {};
  const inverse: Partial<CharStyle> = {};
  const keys = new Set<keyof CharStyle>([...(Object.keys(oldS) as (keyof CharStyle)[]), ...(Object.keys(newS) as (keyof CharStyle)[])]);
  for (const k of keys) {
    if (oldS[k] !== newS[k]) {
      // @ts-expect-error keyed copy
      patch[k] = newS[k];
      // @ts-expect-error keyed copy
      inverse[k] = oldS[k];
    }
  }
  return { patch, inverse };
}

/** Detect a pure style change (text identical, styles differ) and return the
 *  covered [first,last) range + the style delta, or null if it's a text change. */
function detectFormat(oldRuns: Run[], newRuns: Run[]): { start: number; end: number; patch: Partial<CharStyle>; inverse: Partial<CharStyle> } | null {
  if (textOfRuns(oldRuns) !== textOfRuns(newRuns)) return null; // not a pure restyle
  const len = textOfRuns(oldRuns).length;
  let first = -1;
  let last = -1;
  for (let i = 0; i < len; i++) {
    const a = styleAtChar(oldRuns, i)!;
    const b = styleAtChar(newRuns, i)!;
    if (Object.keys(styleDelta(a, b).patch).length > 0) {
      if (first < 0) first = i;
      last = i + 1;
    }
  }
  if (first < 0) return null; // identical
  const a = styleAtChar(oldRuns, first)!;
  const b = styleAtChar(newRuns, first)!;
  const { patch, inverse } = styleDelta(a, b);
  return { start: first, end: last, patch, inverse };
}

function mkSuggestion(kind: Suggestion["kind"], blockId: string, s: number, e: number, author: UserInfo, now: number, extra?: Partial<Suggestion>): Suggestion {
  return { id: freshId(), kind, anchor: { start: pos(blockId, s), end: pos(blockId, e) }, author, createdAt: now, ...extra };
}

/** The block a structural op hangs on for GC (its disappearance kills the
 *  record). For removeBlock that's the still-present (text kept) block; for the
 *  others it's the created / merged-into / host-table block. */
function structuralBlockId(op: Op): string | null {
  switch (op.type) {
    case "splitParagraph":
      return op.newBlockId;
    case "mergeParagraphs":
      return op.firstBlockId;
    case "insertBlock":
      return op.block.id;
    case "removeBlock":
      return op.blockId;
    case "insertTableRow":
    case "removeTableRow":
    case "insertTableColumn":
    case "removeTableColumn":
      return op.tableId;
    default:
      return null;
  }
}

/** Build a structural suggestion for an already-buildable core op, harvesting
 *  the op's EXACT inverse from applyOp (run against the pre-op doc — the result
 *  doc is discarded; we only keep the inverse). The forward op stays in core so
 *  the doc stays live; reject re-applies the stored inverse. */
function mkStructural(op: Op, doc: Document, author: UserInfo, now: number, groupId?: string): Suggestion | null {
  const blockId = structuralBlockId(op);
  if (blockId === null) return null;
  let inverse: Op;
  try {
    inverse = applyOp(doc, op).inverse;
  } catch {
    return null; // op can't apply to this doc → don't track (caller passes through)
  }
  const structural: StructuralChange = { op, inverse, blockId };
  const extra: Partial<Suggestion> = { structural };
  if (groupId !== undefined) extra.groupId = groupId;
  // Degenerate point anchor on the block; paint uses a block-level change-bar.
  return mkSuggestion("structural", blockId, 0, 0, author, now, extra);
}

/** The pass-through result (no tracking). */
const passThrough = (trn: Transaction): InterceptResult => ({ core: trn, reviewOps: [] });

/** One op's contribution to the rewrite: the core ops to actually run (the op
 *  itself, or [] when it became a non-destructive overlay) and the review ops to
 *  emit. Anchors are in the CURRENT running-doc coordinates; the driver rebases
 *  them forward through later ops (mirroring commitCore's per-op rebase). */
interface OpRewrite {
  coreOps: Op[];
  reviewOps: ReviewOp[];
}

/** Route ONE op through the tracked/structural logic against the running doc +
 *  running review. Pure: returns the rewrite; the driver advances the run state.
 *  `groupId` (multi-op transactions) bundles every record from one user action. */
function interceptOne(op: Op, review: ReviewLayer, doc: Document, author: UserInfo, now: number, groupId?: string): OpRewrite {
  const grouped: Partial<Suggestion> = groupId !== undefined ? { groupId } : {};
  switch (op.type) {
    case "insertText":
    case "insertRuns": {
      const blockId = op.at.blockId;
      const offset = op.at.offset;
      if (!blockById(doc, blockId)) return { coreOps: [op], reviewOps: [] }; // dangling target → don't anchor
      const len = op.type === "insertText" ? op.text.length : textOfRuns(op.runs).length;
      if (len === 0) return { coreOps: [op], reviewOps: [] };
      // Contiguous typing extends an existing record via rebase — no new record.
      if (extendsAuthorInsert(review, blockId, offset, author.id)) return { coreOps: [op], reviewOps: [] };
      // Fresh insertion point: anchor is the inserted span in POST-insert coords.
      const s = mkSuggestion("insert", blockId, offset, offset + len, author, now, grouped);
      return { coreOps: [op], reviewOps: [{ type: "addSuggestion", s }] };
    }

    case "deleteRange": {
      const { blockId, start, end } = op;
      if (start >= end || !blockById(doc, blockId)) return { coreOps: [op], reviewOps: [] };
      const covered = fullyCovered(start, end, authorInsertsIn(review, blockId, author.id));
      if (covered) {
        // Deleting only the author's own pending insertion → REAL delete; rebase
        // shrinks/drops those insertion records automatically.
        return { coreOps: [op], reviewOps: [] };
      }
      // Touches original / others' text → non-destructive: drop the core delete,
      // keep the text, mark it deleted.
      const s = mkSuggestion("delete", blockId, start, end, author, now, grouped);
      return { coreOps: [], reviewOps: [{ type: "addSuggestion", s }] };
    }

    case "setRuns": {
      const block = blockById(doc, op.blockId);
      if (!block) return { coreOps: [op], reviewOps: [] };
      const fmt = detectFormat(block.runs, op.runs);
      if (!fmt) return { coreOps: [op], reviewOps: [] }; // text change (case/sdt) → untracked
      const s = mkSuggestion("format", op.blockId, fmt.start, fmt.end, author, now, { patch: fmt.patch, inverse: fmt.inverse, ...grouped });
      return { coreOps: [op], reviewOps: [{ type: "addSuggestion", s }] };
    }

    case "splitParagraph":
    case "mergeParagraphs":
    case "insertBlock":
    case "removeBlock":
    case "insertTableRow":
    case "removeTableRow":
    case "insertTableColumn":
    case "removeTableColumn": {
      // Structural edits stay in core (the doc stays live) and are tracked as a
      // structural suggestion carrying the op + its exact inverse. Reject
      // re-applies the inverse; accept just drops the record.
      const s = mkStructural(op, doc, author, now, groupId);
      if (!s) return { coreOps: [op], reviewOps: [] };
      return { coreOps: [op], reviewOps: [{ type: "addSuggestion", s }] };
    }

    default:
      return { coreOps: [op], reviewOps: [] }; // setParaStyle, table-structure, props…: untracked
  }
}

/** Rebase an addSuggestion review op's anchor forward through a core op's
 *  position-mapper. Only addSuggestion carries an anchor that drifts; the others
 *  intercept emits (it emits none here) are id-keyed and immune. */
function rebaseReviewOp(rop: ReviewOp, mapPosition: (p: { blockId: string; offset: number }) => { blockId: string; offset: number }): ReviewOp {
  if (rop.type !== "addSuggestion") return rop;
  const a = rop.s.anchor;
  return { ...rop, s: { ...rop.s, anchor: { start: mapPosition(a.start), end: mapPosition(a.end) } } };
}

export function intercept(trn: Transaction, review: ReviewLayer, doc: Document, author: UserInfo, now: number): InterceptResult {
  if (trn.ops.length === 0) return passThrough(trn);

  // Single-op fast path: no coordinate drift, no group needed — keep the original
  // transaction object so callers/tests that identity-compare `r.core` still see it.
  if (trn.ops.length === 1) {
    const rw = interceptOne(trn.ops[0]!, review, doc, author, now);
    if (rw.reviewOps.length === 0 && rw.coreOps.length === 1 && rw.coreOps[0] === trn.ops[0]) {
      return passThrough(trn); // untracked → unchanged transaction
    }
    return { core: { ops: rw.coreOps, selectionAfter: trn.selectionAfter, origin: trn.origin }, reviewOps: rw.reviewOps };
  }

  // Multi-op (paste, cross-paragraph delete): decompose. Route each op through the
  // SAME logic against a RUNNING doc + review so later ops see earlier ones'
  // post-state, and rebase already-emitted record anchors forward through each
  // applied core op (coordinate drift — mirrors commitCore's per-op rebase). One
  // shared groupId bundles the whole transaction into a single accept/reject unit.
  const groupId = freshId();
  let runDoc = doc;
  let runReview = review;
  const coreOps: Op[] = [];
  let reviewOps: ReviewOp[] = [];

  for (const op of trn.ops) {
    const rw = interceptOne(op, runReview, runDoc, author, now, groupId);
    // Advance the running doc by the CORE ops this op contributed (a dropped
    // delete leaves the text in place — runDoc must not move past it). Rebase the
    // PREVIOUSLY-emitted records through each — but NOT this op's own new records:
    // interceptOne already expressed those in post-core coordinates.
    for (const cop of rw.coreOps) {
      const res = applyOp(runDoc, cop);
      runDoc = res.doc;
      runReview = rebaseReview(runReview, res.mapPosition);
      reviewOps = reviewOps.map((r) => rebaseReviewOp(r, res.mapPosition));
      coreOps.push(cop);
    }
    // Now stage this op's new records (post-core coords) into both accumulators.
    for (const rop of rw.reviewOps) runReview = applyReviewOp(runReview, rop).layer;
    reviewOps.push(...rw.reviewOps);
  }

  return { core: { ops: coreOps, selectionAfter: trn.selectionAfter, origin: trn.origin }, reviewOps };
}
