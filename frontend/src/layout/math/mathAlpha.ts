// Map a math identifier character to its styled Unicode code point in the
// Mathematical Alphanumeric Symbols block, so a single upright math font (STIX
// Two Math) renders TRUE italic/bold/blackboard/script glyphs — no faux-slanting.
// MathML <mi> defaults to italic; <mi mathvariant="…"> selects the others.
//
// The block is contiguous EXCEPT for letters that already existed in the
// Letterlike Symbols block (Unicode left "holes" and points there instead); those
// exceptions are tabled below. Greek follows the same scheme.

import type { MathVariant } from "@cw/shared";

const A = 0x41, Z = 0x5a, a = 0x61, z = 0x7a;
const isAsciiUpper = (c: number): boolean => c >= A && c <= Z;
const isAsciiLower = (c: number): boolean => c >= a && c <= z;

// Base code points for 'A' of each variant (lower base = upper base + 26).
const BASE: Partial<Record<MathVariant, number>> = {
  bold: 0x1d400,
  italic: 0x1d434,
  "bold-italic": 0x1d468,
  script: 0x1d49c,
  "bold-script": 0x1d4d0,
  fraktur: 0x1d504,
  "double-struck": 0x1d538,
  "bold-fraktur": 0x1d56c,
  "sans-serif": 0x1d5a0,
  "bold-sans-serif": 0x1d5d4,
  "sans-serif-italic": 0x1d608,
  "sans-serif-bold-italic": 0x1d63c,
  monospace: 0x1d670,
};

// Letters that live in the Letterlike Symbols block instead of the contiguous run.
const HOLES: Partial<Record<MathVariant, Record<string, number>>> = {
  italic: { h: 0x210e },
  script: {
    B: 0x212c, E: 0x2130, F: 0x2131, H: 0x210b, I: 0x2110, L: 0x2112, M: 0x2133, R: 0x211b,
    e: 0x212f, g: 0x210a, o: 0x2134,
  },
  fraktur: { C: 0x212d, H: 0x210c, I: 0x2111, R: 0x211c, Z: 0x2128 },
  "double-struck": { C: 0x2102, H: 0x210d, N: 0x2115, P: 0x2119, Q: 0x211a, R: 0x211d, Z: 0x2124 },
};

// Greek math blocks. The uppercase math block reserves a slot (the theta-symbol
// variant) exactly where Unicode Greek reserves U+03A2 between Ρ and Σ, so the
// straight `cp - GREEK_UPPER` offset stays aligned across the whole range
// (Α 0x391→0x1D6E2 … Ω 0x3A9→0x1D6FA). Italic is the default; bold-italic Greek
// has its own block, so don't fold it into italic.
const GREEK_UPPER = 0x391, GREEK_LOWER = 0x3b1, GREEK_LAST_LOWER = 0x3c9;
const GREEK_BASE: Partial<Record<MathVariant, { upper: number; lower: number }>> = {
  italic: { upper: 0x1d6e2, lower: 0x1d6fc },
  "bold-italic": { upper: 0x1d71c, lower: 0x1d736 },
  bold: { upper: 0x1d6a8, lower: 0x1d6c2 },
};

/** The code point STIX should render for an identifier char under `variant`
 *  (default = italic). Non-letters and unmapped chars return their own code point. */
export function styledMathCodePoint(ch: string, variant: MathVariant | undefined): number {
  const cp = ch.codePointAt(0) ?? 0;
  const v: MathVariant = variant ?? "italic";
  if (v === "normal") return cp;

  // ASCII letters.
  if (isAsciiUpper(cp) || isAsciiLower(cp)) {
    const hole = HOLES[v]?.[ch];
    if (hole) return hole;
    const base = BASE[v];
    if (base === undefined) return cp;
    return isAsciiUpper(cp) ? base + (cp - A) : base + 26 + (cp - a);
  }

  // Greek letters — italic / bold-italic / bold (other variants fall back to the
  // literal char, which still renders upright in STIX).
  const gb = GREEK_BASE[v];
  if (gb) {
    if (cp >= GREEK_UPPER && cp <= GREEK_UPPER + 24) return gb.upper + (cp - GREEK_UPPER);
    if (cp >= GREEK_LOWER && cp <= GREEK_LAST_LOWER) return gb.lower + (cp - GREEK_LOWER);
  }
  return cp;
}

/** Styled string for a SINGLE-character identifier (the common case). Multi-char
 *  identifiers (function names like "sin"/"log") render upright unchanged. */
export function styledMathText(text: string, variant: MathVariant | undefined): string {
  if ([...text].length !== 1) return text;
  return String.fromCodePoint(styledMathCodePoint(text, variant));
}
