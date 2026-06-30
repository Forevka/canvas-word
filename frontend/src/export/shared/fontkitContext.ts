// A DOM-free stand-in for the slice of CanvasRenderingContext2D that pretext
// (rich-inline) and metrics.ts measure through. Widths come from fontkit glyph
// advances of the resolved bundled face; vertical metrics from its font tables.
//
// Per-glyph fallback: when the resolved face lacks a codepoint, the advance is
// taken from the bundled fallback face (StixTwoMath) when it has it — NotoSansSC
// is intentionally NOT in the chain (CJK is pre-routed via scriptSplitRuns and is
// opt-out-able). This is the SAME face segmentation the PDF paint path uses, so
// reserved widths always match painted glyph positions.
//
// Letter/word spacing are intentionally ignored: pretext folds letterSpacing in
// arithmetically (RichInlineItem.letterSpacing), and word spacing is a paint-time
// concern. This matches the deterministic test shim (test-canvas-setup.ts), which
// also measures pure advances — the engine never relies on the canvas for either.

import { resolveFont, type ResolvedFont } from "./fontRegistry";
import { segmentByFace } from "./glyphFallback";

interface ParsedFont {
  resolved: ResolvedFont;
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
  const parsed: ParsedFont = { resolved, sizePx };
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
    const { resolved, sizePx } = parseFont(this.font);
    const scale = sizePx / resolved.font.unitsPerEm;

    let width: number;
    if (text.length === 0) {
      width = 0;
    } else {
      // Split by face so fallback-face advances are used for characters the
      // primary face lacks — this must agree with the PDF paint path.
      const segs = segmentByFace(text, resolved);
      if (segs.length === 1) {
        // Common case: all characters in one face — single layout call.
        const seg = segs[0]!;
        width = seg.face.font.layout(text).advanceWidth * (sizePx / seg.face.font.unitsPerEm);
      } else {
        width = 0;
        for (const seg of segs) {
          width += seg.face.font.layout(seg.text).advanceWidth * (sizePx / seg.face.font.unitsPerEm);
        }
      }
    }

    // Vertical metrics always come from the primary face: line height is driven
    // by the run's stated font, not by any fallback glyphs within it.
    const ascent = resolved.font.ascent * scale;
    const descent = -resolved.font.descent * scale; // fontkit descent is negative; canvas wants positive
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
