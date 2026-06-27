// Paint a typeset MathBox to a 2D canvas. The PDF exporter has its own mirror in
// export/pdf (pdfkit API differs); both consume the same PlacedEquation geometry.
// Glyph coordinates are relative to the equation baseline; the caller passes the
// absolute baseline Y (= block.y + equation.baseline).

import { cloneFamilyFor, firstFamilyToken } from "../fonts/clones";
import type { PlacedEquation } from "../layout/layoutTree";

/** Canvas font string for a math glyph, using the same metric clone the layout
 *  measured with so painted widths match the reserved boxes exactly. */
export function mathGlyphFont(family: string, sizePx: number, italic: boolean, bold: boolean): string {
  const clone = cloneFamilyFor(firstFamilyToken(family)).clone;
  return `${italic ? "italic " : ""}${bold ? "700" : "400"} ${sizePx}px ${clone}`;
}

import type { MathBox } from "../layout/math/mathBox";

/** Draw a raw MathBox at an absolute (originX, baselineY). Glyph/rule offsets are
 *  relative to the math baseline. Used for both block equations (baseline =
 *  block.y + box.ascent) and inline equations (baseline = the text line baseline). */
export function paintMathBoxCanvasRaw(
  ctx: CanvasRenderingContext2D,
  box: MathBox,
  originX: number,
  baselineY: number,
  family: string,
  color: string,
): void {
  ctx.textBaseline = "alphabetic";
  (ctx as CanvasRenderingContext2D & { wordSpacing: string }).wordSpacing = "0px";
  ctx.fillStyle = color;
  for (const r of box.rules) ctx.fillRect(originX + r.x, baselineY + r.y, r.w, r.h);
  for (const g of box.glyphs) {
    ctx.font = mathGlyphFont(family, g.sizePx, g.italic, g.bold);
    ctx.fillText(g.text, originX + g.x, baselineY + g.y);
  }
}

export function paintMathBoxCanvas(
  ctx: CanvasRenderingContext2D,
  eq: PlacedEquation,
  originX: number,
  baselineY: number,
  family: string,
  color: string,
): void {
  paintMathBoxCanvasRaw(ctx, eq.box, originX, baselineY, family, color);
}
