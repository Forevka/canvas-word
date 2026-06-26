// The math typesetter's output: a MathBox — a flattened, absolutely-positioned
// set of glyphs and rules (filled rectangles: fraction bars, radical lines) in
// box-local coordinates. Flattening as we compose means the painters (canvas +
// PDF) consume one glyph list and one rule list with no recursion.
//
// Coordinate convention (matches canvas): x grows right, y grows DOWN, the
// box BASELINE is y = 0. `ascent` is the distance above the baseline (a positive
// number), `descent` the distance below. Total height = ascent + descent.

/** A single glyph to paint at its baseline origin. */
export interface PlacedGlyph {
  x: number;
  /** Baseline y (text is drawn from the baseline). */
  y: number;
  text: string;
  sizePx: number;
  italic: boolean;
  bold: boolean;
}

/** A filled rectangle (fraction/radical rule, bar). */
export interface PlacedRule {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface MathBox {
  width: number;
  ascent: number;
  descent: number;
  glyphs: PlacedGlyph[];
  rules: PlacedRule[];
  /** Italic correction of the box's last glyph — used when attaching a
   *  superscript so it clears a slanted base (e.g. f'). 0 for upright content. */
  italicCorrection: number;
}

export const emptyBox = (): MathBox => ({
  width: 0,
  ascent: 0,
  descent: 0,
  glyphs: [],
  rules: [],
  italicCorrection: 0,
});

/** Translate a box by (dx, dy); dy>0 moves content DOWN. Returns a new box. */
export function shiftBox(b: MathBox, dx: number, dy: number): MathBox {
  return {
    width: b.width,
    ascent: b.ascent - dy,
    descent: b.descent + dy,
    italicCorrection: b.italicCorrection,
    glyphs: b.glyphs.map((g) => ({ ...g, x: g.x + dx, y: g.y + dy })),
    rules: b.rules.map((r) => ({ ...r, x: r.x + dx, y: r.y + dy })),
  };
}

/** Place boxes left-to-right on a shared baseline, inserting `gaps[i]` of
 *  horizontal space BEFORE box i (gaps default 0). */
export function hbox(boxes: MathBox[], gaps: number[] = []): MathBox {
  let x = 0;
  let ascent = 0;
  let descent = 0;
  const glyphs: PlacedGlyph[] = [];
  const rules: PlacedRule[] = [];
  for (let i = 0; i < boxes.length; i++) {
    x += gaps[i] ?? 0;
    const b = boxes[i]!;
    for (const g of b.glyphs) glyphs.push({ ...g, x: g.x + x });
    for (const r of b.rules) rules.push({ ...r, x: r.x + x });
    ascent = Math.max(ascent, b.ascent);
    descent = Math.max(descent, b.descent);
    x += b.width;
  }
  const last = boxes[boxes.length - 1];
  return { width: x, ascent, descent, glyphs, rules, italicCorrection: last?.italicCorrection ?? 0 };
}

/** Overlay already-positioned boxes at a shared origin (no horizontal advance):
 *  used to merge the vertically-shifted parts of a fraction / script / radical
 *  into one box. Metrics are the union; width is the max (callers set the final
 *  width explicitly when it differs). `align` is reserved for future use. */
export function overlayCentered(boxes: MathBox[], _align: "left" | "center" = "left"): MathBox {
  const glyphs: PlacedGlyph[] = [];
  const rules: PlacedRule[] = [];
  let ascent = 0;
  let descent = 0;
  let width = 0;
  let italicCorrection = 0;
  for (const b of boxes) {
    glyphs.push(...b.glyphs);
    rules.push(...b.rules);
    ascent = Math.max(ascent, b.ascent);
    descent = Math.max(descent, b.descent);
    width = Math.max(width, b.width);
    italicCorrection = b.italicCorrection;
  }
  return { width, ascent, descent, glyphs, rules, italicCorrection };
}

/** A bare horizontal rule (e.g. fraction bar). `topAboveBaseline` is the distance
 *  from the baseline to the rule's BOTTOM edge (positive = up); the rule occupies
 *  `[-topAboveBaseline - thickness, -topAboveBaseline]`. ascent/descent are kept
 *  non-negative per the box contract (a rule entirely above the baseline has 0
 *  descent), so it never reports a negative extent into `overlayCentered`. */
export function ruleBox(width: number, thickness: number, topAboveBaseline: number): MathBox {
  const ruleTop = -topAboveBaseline - thickness;
  const ruleBottom = -topAboveBaseline;
  return {
    width,
    ascent: Math.max(0, -ruleTop),
    descent: Math.max(0, ruleBottom),
    italicCorrection: 0,
    glyphs: [],
    rules: [{ x: 0, y: ruleTop, w: width, h: thickness }],
  };
}

/** Horizontal padding box (no ink). */
export const spaceBox = (width: number): MathBox => ({
  width,
  ascent: 0,
  descent: 0,
  glyphs: [],
  rules: [],
  italicCorrection: 0,
});

/** Overall pixel bounds of a laid-out equation (for the inline/replaced box). */
export const boxHeight = (b: MathBox): number => b.ascent + b.descent;
