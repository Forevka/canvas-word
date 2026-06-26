// Bridge from a stored MathEquation (model) to a typeset MathBox (geometry).
// Memoized per (equation object, family, size) so repeated layouts/renders of the
// same equation are free; the AST is immutable, so object identity is a safe key.

import type { MathEquation } from "@cw/shared";
import { layoutMathRow } from "./layoutMath";
import { createMathFont, type MathFont } from "./mathFont";
import type { MathBox } from "./mathBox";

/** Display equations lay out a touch larger than body text for prominence. */
export const EQUATION_DISPLAY_PX = 22;

const fontCache = new Map<string, MathFont>();
const boxCache = new WeakMap<MathEquation, Map<string, MathBox>>();

function fontFor(family: string): MathFont {
  let f = fontCache.get(family);
  if (!f) {
    f = createMathFont(family);
    fontCache.set(family, f);
  }
  return f;
}

/** Typeset an equation to a positioned box. Requires a measurement context to be
 *  installed (the browser canvas is automatic; Node hosts call installMeasureHost
 *  first). Inline vs display only differ in `sizePx` chosen by the caller. */
export function equationBox(eq: MathEquation, family: string, sizePx: number): MathBox {
  const key = `${family}|${sizePx}`;
  let byKey = boxCache.get(eq);
  if (byKey) {
    const hit = byKey.get(key);
    if (hit) return hit;
  } else {
    byKey = new Map();
    boxCache.set(eq, byKey);
  }
  const box = layoutMathRow(eq.root.children, {
    font: fontFor(family),
    style: { sizePx, cramped: false, display: eq.display },
  });
  byKey.set(key, box);
  return box;
}
