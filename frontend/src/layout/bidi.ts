// Bidirectional text support (UAX#9) for the layout engine. pretext computes the
// per-character embedding levels (see /pretext-patch — we reuse its UBA impl and
// the explicit-base-level extension); this module turns those levels into a
// VISUAL fragment order and resolves direction-aware paragraph alignment.
//
// Division of labour:
//   - pretext (computeSegmentLevels) — Unicode bidi classes → embedding levels.
//   - this module — UAX#9 rule L2 (reorder runs by level) + alignment mapping.
//   - the canvas (ctx.fillText) — intra-fragment shaping + RTL glyph ordering,
//     which is why a SINGLE-level fragment never needs its characters reversed
//     by us; only the ORDER of fragments on the line does.

import { computeSegmentLevels } from "@chenglou/pretext/bidi";
import type { ParaStyle } from "@cw/shared";

/** The bidi base embedding level for a paragraph: 1 (RTL) for `direction:"rtl"`,
 *  else 0 (LTR). Word's w:bidi forces the base regardless of content. */
export function baseLevelFor(direction: ParaStyle["direction"]): 0 | 1 {
  return direction === "rtl" ? 1 : 0;
}

// Strong-RTL Unicode blocks (Hebrew, Arabic, Syriac, Thaana, NKo, Arabic
// Supplement/Extended, plus the Arabic/Hebrew presentation forms). A cheap gate
// so the 99% LTR-only paragraph never pays for a full bidi scan.
const RTL_CHAR = /[֐-ࣿיִ-﷿ﹰ-﻿]/;

/** Cheap pre-check: could this text need bidi reordering? */
export function hasRtlChars(text: string): boolean {
  return RTL_CHAR.test(text);
}

/** Per-UTF-16-code-unit embedding levels for `text` under `baseLevel`, or null
 *  when the text is entirely the base direction (so no reordering is needed —
 *  the LTR fast path, and pure-base-direction lines). */
export function levelsFor(text: string, baseLevel: 0 | 1): Int8Array | null {
  const n = text.length;
  if (n === 0) return null;
  // segStarts = every code unit, so the returned array is per-code-unit.
  const segStarts = new Array<number>(n);
  for (let i = 0; i < n; i++) segStarts[i] = i;
  return computeSegmentLevels(text, segStarts, baseLevel);
}

const reverseInPlace = <T>(a: T[], lo: number, hi: number): void => {
  while (lo < hi) {
    const t = a[lo]!;
    a[lo] = a[hi]!;
    a[hi] = t;
    lo++;
    hi--;
  }
};

/** UAX#9 rule L2: given the embedding level of each item in LOGICAL order,
 *  return the indices in VISUAL (left-to-right) order. Reverses each contiguous
 *  run of level ≥ L, from the highest level down to the lowest odd level. */
export function visualOrder(levels: readonly number[]): number[] {
  const n = levels.length;
  const order = Array.from({ length: n }, (_, i) => i);
  if (n <= 1) return order;
  const lev = levels.slice();
  let max = 0;
  let minOdd = Number.POSITIVE_INFINITY;
  for (const l of levels) {
    if (l > max) max = l;
    if (l & 1 && l < minOdd) minOdd = l;
  }
  if (minOdd === Number.POSITIVE_INFINITY) return order; // all even → no reorder
  for (let level = max; level >= minOdd; level--) {
    let i = 0;
    while (i < n) {
      if (lev[i]! >= level) {
        let j = i;
        while (j < n && lev[j]! >= level) j++;
        reverseInPlace(order, i, j - 1);
        reverseInPlace(lev, i, j - 1);
        i = j;
      } else i++;
    }
  }
  return order;
}

/** Map a paragraph's stored alignment onto a PHYSICAL alignment given its base
 *  direction. In an RTL paragraph the "start" edge is the right, so `left`/`right`
 *  are interpreted as start/end and swap; center/justify are unchanged. This makes
 *  an RTL paragraph right-aligned by default (matching Word's w:bidi behaviour). */
export function effectiveAlign(
  align: ParaStyle["align"],
  direction: ParaStyle["direction"],
): ParaStyle["align"] {
  if (direction !== "rtl") return align;
  if (align === "left") return "right";
  if (align === "right") return "left";
  return align;
}
