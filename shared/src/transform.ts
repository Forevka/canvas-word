// Operational transform — the concurrency core.
//
// transformOp(a, b, side) rewrites op `a` so it can be applied AFTER op `b`,
// preserving both authors' intent (a' such that [b, a'] reaches the same state
// [a, b'] would). It is the inclusion transform of classic OT.
//
// Scope: the genuinely-concurrent cases are intra-paragraph text edits, where
// two ops carry integer offsets into the same block. Ops on DIFFERENT blocks (or
// tables) commute — they transform to themselves. Coarse whole-subtree / scalar
// ops on the SAME target (setParaStyle, setRuns, setTableStructure, …) are left
// as last-writer-wins: both apply, and the later seq overwrites — no offset
// surgery is meaningful for them. This mirrors the plan: fine transform for
// position-bearing ops, LWW for the rest.
//
// `side` breaks insert/insert ties at the same offset: the "left" op keeps its
// position (its content ends up before the other's); the "right" op shifts past.
// When rebasing a local pending op against an already-ordered remote op, the
// remote won the order, so the local op is the "right" side.

import type { Op } from "./model/ops";
import type { DocPosition } from "./model/position";
import { textOfRuns } from "./model/text";

export type Side = "left" | "right";

/** The block an op addresses for position transforms, or null for ops whose
 *  positions don't participate (coarse/structural — treated as LWW). */
function textBlock(op: Op): string | null {
  switch (op.type) {
    case "insertText":
    case "insertRuns":
    case "splitParagraph":
      return op.at.blockId;
    case "deleteRange":
      return op.blockId;
    default:
      return null;
  }
}

/** Number of UTF-16 units an op inserts at its point (0 for non-inserts). */
function insertedLen(op: Op): number {
  if (op.type === "insertText") return op.text.length;
  if (op.type === "insertRuns") return textOfRuns(op.runs).length;
  return 0;
}

/** Shift a single (blockId, offset) position to its place AFTER op `b`. */
function shiftPos(blockId: string, offset: number, b: Op, side: Side): DocPosition {
  if (textBlock(b) !== blockId) return { blockId, offset };
  switch (b.type) {
    case "insertText":
    case "insertRuns": {
      const bo = b.at.offset;
      const len = insertedLen(b);
      // Insert strictly before us, or at the same point when we yield (right).
      if (bo < offset || (bo === offset && side === "right")) return { blockId, offset: offset + len };
      return { blockId, offset };
    }
    case "deleteRange": {
      const len = b.end - b.start;
      if (offset <= b.start) return { blockId, offset };
      if (offset >= b.end) return { blockId, offset: offset - len };
      return { blockId, offset: b.start }; // inside the removed range → collapse to its start
    }
    case "splitParagraph": {
      // Everything after the split point moves into the new paragraph.
      const bo = b.at.offset;
      if (offset > bo) return { blockId: b.newBlockId, offset: offset - bo };
      return { blockId, offset };
    }
    default:
      return { blockId, offset };
  }
}

/** Transform `a` to apply after `b`. Returns 0+ ops: usually one, but a delete
 *  that straddles a concurrent insert splits into two ranges around the surviving
 *  inserted text (a single range can't exclude a middle point). */
export function transformOp(a: Op, b: Op, side: Side): Op[] {
  const aBlk = textBlock(a);
  const bBlk = textBlock(b);
  // Coarse op on either side, or ops on different blocks → no offset change.
  if (aBlk === null || bBlk === null || aBlk !== bBlk) return [a];

  switch (a.type) {
    case "insertText":
    case "insertRuns": {
      return [{ ...a, at: shiftPos(a.at.blockId, a.at.offset, b, side) }];
    }
    case "deleteRange": {
      // An insert strictly inside our range must survive — split the delete
      // around it (later range first so applying both doesn't shift the earlier).
      if ((b.type === "insertText" || b.type === "insertRuns") && textBlock(b) === a.blockId) {
        const bo = b.at.offset;
        const len = insertedLen(b);
        if (a.start < bo && bo < a.end) {
          return [
            { ...a, start: bo + len, end: a.end + len },
            { ...a, start: a.start, end: bo },
          ];
        }
      }
      const s = shiftPos(a.blockId, a.start, b, side);
      const e = shiftPos(a.blockId, a.end, b, side);
      if (s.offset >= e.offset) return []; // fully swallowed by a concurrent delete
      return [{ ...a, blockId: s.blockId, start: s.offset, end: e.offset }];
    }
    case "splitParagraph": {
      return [{ ...a, at: shiftPos(a.at.blockId, a.at.offset, b, side) }];
    }
    default:
      return [a];
  }
}

/** Transform a list of ops against one op. */
export function transformOpsAgainst(ops: readonly Op[], b: Op, side: Side): Op[] {
  return ops.flatMap((a) => transformOp(a, b, side));
}

/** Transform `ops` to apply after the sequential run `against` (each subsequent
 *  against-op lives in the space the previous one produced). */
export function transformOps(ops: readonly Op[], against: readonly Op[], side: Side): Op[] {
  let cur: Op[] = ops.slice();
  for (const b of against) cur = transformOpsAgainst(cur, b, side);
  return cur;
}
