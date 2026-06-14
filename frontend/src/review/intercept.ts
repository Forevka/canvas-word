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
// V1 scope (documented in REVIEW.md §8): single insertText/insertRuns, single
// deleteRange, and single style-only setRuns are tracked. Structural ops
// (split/merge/block/table) and multi-op transactions (cross-paragraph deletes,
// paste) pass through to core UNTRACKED — finer handling is V2.

import {
  blockById,
  textOfRuns,
  type CharStyle,
  type Document,
  type Op,
  type Run,
  type ReviewOp,
  type ReviewLayer,
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

/** The pass-through result (no tracking). */
const passThrough = (trn: Transaction): InterceptResult => ({ core: trn, reviewOps: [] });

export function intercept(trn: Transaction, review: ReviewLayer, doc: Document, author: UserInfo, now: number): InterceptResult {
  // Only single-op transactions of the tracked kinds are rewritten; everything
  // else passes through untracked (V1 boundary, REVIEW.md §8).
  if (trn.ops.length !== 1) return passThrough(trn);
  const op: Op = trn.ops[0]!;

  switch (op.type) {
    case "insertText":
    case "insertRuns": {
      const blockId = op.at.blockId;
      const offset = op.at.offset;
      const len = op.type === "insertText" ? op.text.length : textOfRuns(op.runs).length;
      if (len === 0) return passThrough(trn);
      // Contiguous typing extends an existing record via rebase — no new record.
      if (extendsAuthorInsert(review, blockId, offset, author.id)) return { core: trn, reviewOps: [] };
      // Fresh insertion point: anchor is the inserted span in POST-insert coords.
      const s = mkSuggestion("insert", blockId, offset, offset + len, author, now);
      return { core: trn, reviewOps: [{ type: "addSuggestion", s }] };
    }

    case "deleteRange": {
      const { blockId, start, end } = op;
      if (start >= end) return passThrough(trn);
      const covered = fullyCovered(start, end, authorInsertsIn(review, blockId, author.id));
      if (covered) {
        // Deleting only the author's own pending insertion → REAL delete; rebase
        // shrinks/drops those insertion records automatically.
        return { core: trn, reviewOps: [] };
      }
      // Touches original / others' text → non-destructive: drop the core delete,
      // keep the text, mark it deleted. Caret already moves to `start`
      // (deleteBackward set selectionAfter there).
      const s = mkSuggestion("delete", blockId, start, end, author, now);
      return {
        core: { ops: [], selectionAfter: trn.selectionAfter, origin: trn.origin },
        reviewOps: [{ type: "addSuggestion", s }],
      };
    }

    case "setRuns": {
      const block = blockById(doc, op.blockId);
      if (!block) return passThrough(trn);
      const fmt = detectFormat(block.runs, op.runs);
      if (!fmt) return passThrough(trn); // text change (case/sdt) → untracked V1
      const s = mkSuggestion("format", op.blockId, fmt.start, fmt.end, author, now, { patch: fmt.patch, inverse: fmt.inverse });
      // The restyle is length-preserving and applied as-is (so it shows); the
      // record drives the change-bar + reject (re-apply inverse).
      return { core: trn, reviewOps: [{ type: "addSuggestion", s }] };
    }

    default:
      return passThrough(trn); // structural ops: untracked V1
  }
}
