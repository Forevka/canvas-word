// A DOM-free stand-in for the slice of CanvasRenderingContext2D that pretext
// (rich-inline) and metrics.ts measure through. Widths come from fontkit glyph
// advances of the resolved bundled face; vertical metrics from its font tables.
//
// Letter/word spacing are intentionally ignored: pretext folds letterSpacing in
// arithmetically (RichInlineItem.letterSpacing), and word spacing is a paint-time
// concern. This matches the deterministic test shim (test-canvas-setup.ts), which
// also measures pure advances — the engine never relies on the canvas for either.

import { resolveFont } from "./fontRegistry";

interface ParsedFont {
  font: import("fontkit").Font;
  sizePx: number;
}

const parseCache = new Map<string, ParsedFont>();

// charStyleToFont emits e.g. "italic 700 16px Calibri, serif".
function parseFont(spec: string): ParsedFont {
  const hit = parseCache.get(spec);
  if (hit) return hit;

  const italic = /(^|\s)italic(\s|$)/.test(spec);
  const bold = /(^|\s)(700|bold)(\s|$)/.test(spec);
  const sizeMatch = /(\d+(?:\.\d+)?)px/.exec(spec);
  const sizePx = sizeMatch ? parseFloat(sizeMatch[1]!) : 16;
  // Everything after "<size>px " is the CSS family stack; take the first token.
  const afterPx = sizeMatch ? spec.slice((sizeMatch.index ?? 0) + sizeMatch[0].length) : spec;
  const family = (afterPx.split(",")[0] ?? "sans-serif").trim() || "sans-serif";

  const resolved = resolveFont(family, bold, italic);
  const parsed: ParsedFont = { font: resolved.font, sizePx };
  parseCache.set(spec, parsed);
  return parsed;
}

interface TextMetricsLike {
  width: number;
  fontBoundingBoxAscent: number;
  fontBoundingBoxDescent: number;
  actualBoundingBoxAscent: number;
  actualBoundingBoxDescent: number;
  actualBoundingBoxLeft: number;
  actualBoundingBoxRight: number;
}

/** Implements `{ font; measureText }` for both pretext and metrics.ts. */
export class FontkitMeasureContext {
  font = "16px sans-serif";
  // Present so callers that poke these canvas props don't throw; no effect here.
  letterSpacing = "0px";
  wordSpacing = "0px";

  measureText(text: string): TextMetricsLike {
    const { font, sizePx } = parseFont(this.font);
    const upm = font.unitsPerEm;
    const scale = sizePx / upm;
    // fontkit's layout applies the font's default shaping (kerning/ligatures),
    // matching what canvas measureText does for the same string.
    const width = text.length === 0 ? 0 : font.layout(text).advanceWidth * scale;
    const ascent = font.ascent * scale;
    const descent = -font.descent * scale; // fontkit descent is negative; canvas wants positive
    return {
      width,
      fontBoundingBoxAscent: ascent,
      fontBoundingBoxDescent: descent,
      actualBoundingBoxAscent: ascent,
      actualBoundingBoxDescent: descent,
      actualBoundingBoxLeft: 0,
      actualBoundingBoxRight: width,
    };
  }
}
