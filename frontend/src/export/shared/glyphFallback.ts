// Per-glyph font fallback for PDF export — when the resolved face lacks a
// glyph for a codepoint, try StixTwoMath (the bundled math/symbol font) so
// technical symbols like ✓ U+2713 and ☒ U+2612 render as real glyphs rather
// than .notdef/tofu.
//
// Used by BOTH the fontkit measure context (fontkitContext.ts, so reserved
// widths are correct) and the PDF paint path (paintBlock.ts paintLine, so
// painted glyphs land exactly where measured). The two sides must use the
// SAME face for each character or layout and paint will disagree.
//
// NotoSansSC is intentionally NOT in the fallback chain: CJK text is
// pre-routed by scriptSplitRuns and the user can opt out via
// `cjk.fallbackFont: ""`. Including NotoSansSC here would bypass that.

import { MATH_FONT_FAMILY } from "../../fonts/clones";
import { resolveFont, type ResolvedFont } from "./fontRegistry";

// Only StixTwoMath is in the per-glyph fallback chain. NotoSansSC is
// intentionally excluded: CJK text is pre-routed via scriptSplitRuns and the
// user can opt out of that via `cjk.fallbackFont: ""`. Including NotoSansSC
// here would override that opt-out by silently embedding it for CJK glyphs
// that slip through. StixTwoMath covers the symbols we actually need
// (✓ U+2713, ☒ U+2612, ☐/☑ U+2610–2611, and the full dingbat range).
const FALLBACK_FAMILIES = [MATH_FONT_FAMILY] as const;

/** A contiguous run of `text` that should all be rendered in `face`. */
export interface FaceSegment {
  text: string;
  face: ResolvedFont;
}

/**
 * Return the best bundled face for a single Unicode code point.
 * Returns `primary` when it has the glyph; otherwise tries each FALLBACK_FAMILY
 * in order and returns the first that has it; falls back to `primary` if none
 * does (unavoidable .notdef, but nothing further can be done).
 */
function faceForCP(cp: number, primary: ResolvedFont): ResolvedFont {
  if (primary.font.hasGlyphForCodePoint(cp)) return primary;
  for (const fam of FALLBACK_FAMILIES) {
    const fb = resolveFont(fam, false, false);
    // Guard against a fallback that itself fell through to the Arimo default
    // (e.g. when its font file failed to load): such a face would be substituted
    // and still not have the glyph, so skip it.
    if (!fb.substituted && fb.font.hasGlyphForCodePoint(cp)) return fb;
  }
  return primary;
}

/**
 * Split `text` into contiguous segments that all use the same face.
 * Adjacent code points that map to the same face are merged into one segment.
 *
 * Fast path: pure-ASCII strings (code points ≤ 0x7F) return a single
 * primary-face segment without any per-glyph checks — all bundled Latin
 * clones cover the full ASCII range, so there is never a fallback for ASCII.
 */
export function segmentByFace(text: string, primary: ResolvedFont): FaceSegment[] {
  if (text.length === 0) return [];

  // Fast path: ASCII-only text — no per-glyph checks needed.
  let hasNonAscii = false;
  for (let i = 0; i < text.length; i++) {
    if (text.charCodeAt(i) > 0x007f) {
      hasNonAscii = true;
      break;
    }
  }
  if (!hasNonAscii) return [{ text, face: primary }];

  const segs: FaceSegment[] = [];
  let curFace: ResolvedFont | null = null;
  let curText = "";

  // Iterate over Unicode code points so surrogate pairs (e.g. emoji) are
  // treated as a single unit rather than two half-characters.
  for (const char of text) {
    const cp = char.codePointAt(0)!;
    // Variation selectors (VS1–16 U+FE00–FE0F, and the supplementary range
    // U+E0100–E01EF) modify the PRECEDING base character — keep them glued to
    // its segment so e.g. ☑️ (U+2611 U+FE0F) stays one run on the symbol face
    // instead of splitting the selector onto the primary face.
    const isVariationSelector = (cp >= 0xfe00 && cp <= 0xfe0f) || (cp >= 0xe0100 && cp <= 0xe01ef);
    if (isVariationSelector && curFace !== null) {
      curText += char;
      continue;
    }
    const face = faceForCP(cp, primary);
    if (curFace !== null && curFace.file === face.file) {
      curText += char;
    } else {
      if (curFace !== null && curText) segs.push({ text: curText, face: curFace });
      curFace = face;
      curText = char;
    }
  }
  if (curText && curFace !== null) segs.push({ text: curText, face: curFace });
  return segs;
}
