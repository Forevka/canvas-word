// The math typesetter: a MathML AST node -> a positioned MathBox. Pure (depends
// only on the injected MathFont provider), so it runs in the editor, the PDF
// exporter, and unit tests alike. Geometry follows classic TeX rules; the
// constants come from the provider (TeX defaults today, STIX MATH table later).

import type { MathNode } from "@cw/shared";
import { boxHeight, emptyBox, hbox, type MathBox, overlayCentered, ruleBox, shiftBox, spaceBox } from "./mathBox";
import type { MathFont } from "./mathFont";
import { styledMathText } from "./mathAlpha";

const UPRIGHT: Face = { italic: false, bold: false };

export interface MathStyle {
  sizePx: number;
  /** Cramped style (denominators, under-radicals) suppresses superscript raise. */
  cramped: boolean;
  /** Display style (block equations): big operators enlarge and take over/under
   *  limits. Inline (text) style keeps them small with scripts. Default false. */
  display?: boolean;
}

export interface MathLayoutCtx {
  font: MathFont;
  style: MathStyle;
}

/** Lay out an equation root's children as a single horizontal box. */
export function layoutMathRow(children: MathNode[], ctx: MathLayoutCtx): MathBox {
  return rowBox(children, ctx);
}

/** Lay out one node. */
export function layoutMathNode(n: MathNode, ctx: MathLayoutCtx): MathBox {
  switch (n.type) {
    case "row":
      return rowBox(n.children, ctx);
    case "ident":
      // Remap to the Mathematical Alphanumeric block so STIX renders true
      // italic/bold/blackboard glyphs (upright face — the codepoint carries style).
      return atom(styledMathText(n.text, n.variant), UPRIGHT, ctx);
    case "number":
      return atom(n.text, UPRIGHT, ctx);
    case "op":
      return opAtom(n.text, ctx);
    case "text":
      return atom(n.text, UPRIGHT, ctx);
    case "space":
      return spaceBox((n.widthEm ?? 0) * ctx.style.sizePx);
    case "frac":
      return fracBox(n.num, n.den, ctx, n.thickness === "0", n.bevelled === true);
    case "script":
      // Scripts on a limit-style operator in display style sit UNDER/OVER
      // (∑_{i=1}^{n}); integrals keep sub/superscripts (LIMITS_OPS, not BIG_OPS —
      // matching naryBox), everywhere else they're true sub/superscripts.
      if (ctx.style.display && n.base.type === "op" && LIMITS_OPS.has(n.base.text)) {
        return layoutMathNode(
          { type: "limit", base: n.base, ...(n.sub ? { under: n.sub } : {}), ...(n.sup ? { over: n.sup } : {}) },
          ctx,
        );
      }
      return scriptBox(n.base, n.sub, n.sup, ctx);
    case "radical":
      return radicalBox(n.radicand, n.index, ctx);
    case "fenced":
      return fencedBox(n.open, n.close, n.child, ctx);
    case "limit":
      return limitBox(n.base, n.under, n.over, ctx);
    case "nary":
      return naryBox(n, ctx);
    case "matrix":
      return matrixBox(n.rows, ctx);
    case "phantom": {
      const b = layoutMathNode(n.child, ctx);
      return { ...emptyBox(), width: b.width, ascent: b.ascent, descent: b.descent };
    }
    case "unknown":
      return atom("▯", { italic: false, bold: false }, ctx); // ▯ placeholder
  }
}

// ── Atoms ────────────────────────────────────────────────────────────────────

interface Face {
  italic: boolean;
  bold: boolean;
}

function atom(text: string, face: Face, ctx: MathLayoutCtx): MathBox {
  if (text === "") return emptyBox();
  const m = ctx.font.measure(text, ctx.style.sizePx, face.italic, face.bold);
  return {
    width: m.width,
    ascent: m.ascent,
    descent: m.descent,
    italicCorrection: m.italicCorrection,
    glyphs: [{ x: 0, y: 0, text, sizePx: ctx.style.sizePx, italic: face.italic, bold: face.bold }],
    rules: [],
  };
}

// ── Rows + spacing ───────────────────────────────────────────────────────────

type AtomClass = "ord" | "op" | "bin" | "rel" | "punct" | "open" | "close";

const REL = new Set("=≠<>≤≥≈≡∈∉⊂⊆⊃⊇→←↔⇒⇐⇔∼≅∝".split(""));
const BIN = new Set("+−-±∓×÷·∗⋅∪∩∨∧⊕⊗∖".split(""));
const PUNCT = new Set(",;".split(""));
const OPEN = new Set("([{⟨".split(""));
const CLOSE = new Set(")]}⟩".split(""));

function classOf(n: MathNode): AtomClass {
  if (n.type === "op") {
    const t = n.text;
    if (REL.has(t)) return "rel";
    if (BIN.has(t)) return "bin";
    if (PUNCT.has(t)) return "punct";
    if (OPEN.has(t)) return "open";
    if (CLOSE.has(t)) return "close";
  }
  if (n.type === "fenced") return "close"; // already self-delimited
  return "ord";
}

/** Inter-atom space (px). A binary operator next to a non-operand acts unary. */
function gapBetween(prev: MathNode, cur: MathNode, c: ReturnType<MathFont["constants"]>): number {
  let a = classOf(prev);
  const b = classOf(cur);
  // Unary context: a leading/operator-preceded bin/rel collapses to ordinary.
  const operandBefore = a === "ord" || a === "close" || a === "punct";
  const bb: AtomClass = (b === "bin" && !operandBefore) ? "ord" : b;
  if (a === "bin" && !["ord", "close"].includes(b)) a = "ord";
  if (a === "rel" || bb === "rel") return c.thickSpace;
  if (a === "bin" || bb === "bin") return c.medSpace;
  if (a === "op" || bb === "op") return c.thinSpace;
  if (a === "punct") return c.thinSpace;
  return 0;
}

function rowBox(children: MathNode[], ctx: MathLayoutCtx): MathBox {
  if (children.length === 0) return placeholderSlot(ctx);
  const c = ctx.font.constants(ctx.style.sizePx);
  const boxes = children.map((ch) => layoutMathNode(ch, ctx));
  const gaps: number[] = [0];
  for (let i = 1; i < children.length; i++) gaps.push(gapBetween(children[i - 1]!, children[i]!, c));
  return hbox(boxes, gaps);
}

/** An empty editable slot renders as a faint dotted box at x-height. */
function placeholderSlot(ctx: MathLayoutCtx): MathBox {
  const s = ctx.style.sizePx;
  // Reserve an em-ish empty box; the painter draws the dotted slot outline from
  // the (zero-child) row in the model, so no ink is emitted here.
  return { width: 0.5 * s, ascent: 0.7 * s, descent: 0, italicCorrection: 0, glyphs: [], rules: [] };
}

// ── Scripts (sub / sup / subsup) ─────────────────────────────────────────────

function scriptCtx(ctx: MathLayoutCtx, cramped?: boolean): MathLayoutCtx {
  const c = ctx.font.constants(ctx.style.sizePx);
  return {
    font: ctx.font,
    style: {
      sizePx: Math.max(6, ctx.style.sizePx * c.scriptPercent),
      cramped: cramped ?? ctx.style.cramped,
      display: false, // scripts are never display style
    },
  };
}

// Big operators: ∑ ∏ … take over/under limits in display; ∫ ∮ keep sub/sup bounds.
const LIMITS_OPS = new Set("∑∏∐⋀⋁⋂⋃⨀⨁⨂⨄⨆⨅".split(""));
const BIG_OPS = new Set([...LIMITS_OPS, ..."∫∬∭∮∯∰∱∲∳⨋⨍⨎⨏"]);

/** A glyph rendered at `size` and vertically centered on the math axis (for
 *  enlarged big operators and stretched delimiters). */
function axisCenteredGlyph(chr: string, size: number, ctx: MathLayoutCtx, c: ReturnType<MathFont["constants"]>): MathBox {
  const m = ctx.font.measure(chr, size, false, false);
  const dy = (m.ascent - m.descent) / 2 - c.axisHeight;
  return {
    width: m.width,
    ascent: m.ascent - dy,
    descent: m.descent + dy,
    italicCorrection: 0,
    glyphs: [{ x: 0, y: dy, text: chr, sizePx: size, italic: false, bold: false }],
    rules: [],
  };
}

/** Operator atom — enlarges a big operator (centered on the axis) in display style. */
function opAtom(text: string, ctx: MathLayoutCtx): MathBox {
  if (ctx.style.display && BIG_OPS.has(text)) {
    const c = ctx.font.constants(ctx.style.sizePx);
    return axisCenteredGlyph(text, c.displayOperatorMinHeight / 0.62, ctx, c);
  }
  return atom(text, UPRIGHT, ctx);
}

function scriptBox(base: MathNode, sub: MathNode | undefined, sup: MathNode | undefined, ctx: MathLayoutCtx): MathBox {
  const c = ctx.font.constants(ctx.style.sizePx);
  const baseBox = layoutMathNode(base, ctx);
  const sctx = scriptCtx(ctx);
  const supBox = sup ? layoutMathNode(sup, scriptCtx(ctx, true)) : undefined;
  const subBox = sub ? layoutMathNode(sub, sctx) : undefined;

  const parts: MathBox[] = [baseBox];
  let scriptWidth = 0;

  if (supBox) {
    const raise = Math.max(c.supShiftUp, baseBox.ascent - 0.25 * sctx.style.sizePx);
    parts.push(shiftBox(slideX(supBox, baseBox.width + baseBox.italicCorrection), 0, -raise));
    scriptWidth = Math.max(scriptWidth, baseBox.italicCorrection + supBox.width);
  }
  if (subBox) {
    let drop = Math.max(c.subShiftDown, baseBox.descent + 0.2 * sctx.style.sizePx);
    if (supBox) {
      // Keep a minimum gap between the two scripts.
      const supBottom = -(Math.max(c.supShiftUp, baseBox.ascent - 0.25 * sctx.style.sizePx)) + supBox.descent;
      const subTop = drop - subBox.ascent;
      if (subTop - supBottom < c.supSubGapMin) drop = supBottom + c.supSubGapMin + subBox.ascent;
    }
    parts.push(shiftBox(slideX(subBox, baseBox.width), 0, drop));
    scriptWidth = Math.max(scriptWidth, subBox.width);
  }

  const merged = overlayCentered(parts, "left");
  return { ...merged, width: baseBox.width + scriptWidth + c.spaceAfterScript };
}

/** Translate a box right by dx without changing its metrics' baseline. */
function slideX(b: MathBox, dx: number): MathBox {
  return {
    ...b,
    glyphs: b.glyphs.map((g) => ({ ...g, x: g.x + dx })),
    rules: b.rules.map((r) => ({ ...r, x: r.x + dx })),
  };
}

// ── Fractions ────────────────────────────────────────────────────────────────

function fracBox(
  numN: MathNode,
  denN: MathNode,
  ctx: MathLayoutCtx,
  noBar: boolean,
  bevelled: boolean,
): MathBox {
  const c = ctx.font.constants(ctx.style.sizePx);
  const num = layoutMathNode(numN, ctx);
  const den = layoutMathNode(denN, { font: ctx.font, style: { ...ctx.style, cramped: true } });
  if (bevelled) {
    // Diagonal fraction a⁄b: lay out inline with a slash (approximation).
    return hbox([num, atom("/", { italic: false, bold: false }, ctx), den]);
  }
  const t = noBar ? 0 : c.ruleThickness;
  const axis = c.axisHeight;
  const W = Math.max(num.width, den.width) + 2 * c.thinSpace;

  const numBaseline = Math.min(
    -c.fracNumShiftUp,
    -(axis + t / 2) - c.fracNumGapMin - num.descent,
  );
  const denBaseline = Math.max(
    c.fracDenShiftDown,
    -(axis - t / 2) + c.fracDenGapMin + den.ascent,
  );

  const numPlaced = shiftBox(center(num, W), 0, numBaseline);
  const denPlaced = shiftBox(center(den, W), 0, denBaseline);
  const parts = [numPlaced, denPlaced];
  if (t > 0) parts.push(ruleBox(W, t, axis - t / 2));
  const merged = overlayCentered(parts, "left");
  return { ...merged, width: W, italicCorrection: 0 };
}

function center(b: MathBox, W: number): MathBox {
  return slideX(b, (W - b.width) / 2);
}

// ── Radicals ─────────────────────────────────────────────────────────────────

function radicalBox(radicandN: MathNode, indexN: MathNode | undefined, ctx: MathLayoutCtx): MathBox {
  const c = ctx.font.constants(ctx.style.sizePx);
  const rad = layoutMathNode(radicandN, { font: ctx.font, style: { ...ctx.style, cramped: true } });
  const t = c.radicalRuleThickness;
  const gap = c.radicalVerticalGap;
  const extra = c.radicalExtraAscender;

  // Surd glyph sized to the radicand height (approximation until stretchy glyphs).
  const surdHeight = rad.ascent + rad.descent + gap + t;
  const surd = ctx.font.measure("√", Math.max(ctx.style.sizePx, surdHeight * 0.9), false, false);
  const surdBox: MathBox = {
    width: surd.width,
    ascent: rad.ascent + gap + t + extra,
    descent: rad.descent,
    italicCorrection: 0,
    glyphs: [{ x: 0, y: 0, text: "√", sizePx: Math.max(ctx.style.sizePx, surdHeight * 0.9), italic: false, bold: false }],
    rules: [],
  };

  const radShift = slideX(rad, surd.width);
  const vinculum = slideX(ruleBox(rad.width, t, rad.ascent + gap), surd.width);
  const body = overlayCentered([surdBox, radShift, vinculum], "left");
  const bodyBox: MathBox = { ...body, width: surd.width + rad.width };

  if (!indexN) return bodyBox;
  // Degree (nth root): a small index nestled into the surd's crook.
  const deg = layoutMathNode(indexN, scriptCtx(scriptCtx(ctx)));
  const degRaise = c.radicalDegreeRaise * surdHeight;
  const degPlaced = shiftBox(deg, 0, -degRaise);
  const shiftedBody = slideX(bodyBox, deg.width);
  const combined = overlayCentered([degPlaced, shiftedBody], "left");
  return { ...combined, width: deg.width + bodyBox.width };
}

// ── Fences ───────────────────────────────────────────────────────────────────

function fencedBox(open: string, close: string, childN: MathNode, ctx: MathLayoutCtx): MathBox {
  const c = ctx.font.constants(ctx.style.sizePx);
  const child = layoutMathNode(childN, ctx);
  const boxes: MathBox[] = [];
  if (open) boxes.push(delimiterBox(open, child, ctx, c));
  boxes.push(child);
  if (close) boxes.push(delimiterBox(close, child, ctx, c));
  return hbox(boxes);
}

/** A fence delimiter that GROWS to its content: when the content is taller than a
 *  normal glyph, the bracket is drawn at an enlarged size and vertically centered
 *  on the math axis (a size-variant approximation — no MATH table needed). */
function delimiterBox(chr: string, child: MathBox, ctx: MathLayoutCtx, c: ReturnType<MathFont["constants"]>): MathBox {
  const base = ctx.style.sizePx;
  const contentH = child.ascent + child.descent;
  if (contentH <= base * 1.05) return atom(chr, UPRIGHT, ctx); // no stretch needed
  const size = Math.min(base * 3.5, contentH + 2 * c.ruleThickness);
  return axisCenteredGlyph(chr, size, ctx, c);
}

// ── Under / over limits ──────────────────────────────────────────────────────

function limitBox(baseN: MathNode, underN: MathNode | undefined, overN: MathNode | undefined, ctx: MathLayoutCtx): MathBox {
  const c = ctx.font.constants(ctx.style.sizePx);
  const base = layoutMathNode(baseN, ctx);
  const sctx = scriptCtx(ctx);
  const W = Math.max(base.width, underN ? layoutMathNode(underN, sctx).width : 0, overN ? layoutMathNode(overN, sctx).width : 0);
  const parts: MathBox[] = [center(base, W)];
  if (overN) {
    const over = center(layoutMathNode(overN, sctx), W);
    parts.push(shiftBox(over, 0, -(base.ascent + c.thinSpace + over.descent)));
  }
  if (underN) {
    const under = center(layoutMathNode(underN, sctx), W);
    parts.push(shiftBox(under, 0, base.descent + c.thinSpace + under.ascent));
  }
  const merged = overlayCentered(parts, "left");
  return { ...merged, width: W };
}

// ── N-ary operators ──────────────────────────────────────────────────────────

function naryBox(n: Extract<MathNode, { type: "nary" }>, ctx: MathLayoutCtx): MathBox {
  const c = ctx.font.constants(ctx.style.sizePx);
  // The operator (enlarged for big-ops in display by opAtom) takes over/under
  // limits in display style for ∑ ∏ …, else sub/superscripts (∫). The op node
  // routes through opAtom, so the bounds position around the enlarged glyph.
  const opNode: MathNode = { type: "op", text: n.hideOp ? "" : n.op };
  const limits = ctx.style.display === true && LIMITS_OPS.has(n.op);
  const opCluster: MathNode = limits
    ? { type: "limit", base: opNode, ...(n.sub ? { under: n.sub } : {}), ...(n.sup ? { over: n.sup } : {}) }
    : { type: "script", base: opNode, ...(n.sub ? { sub: n.sub } : {}), ...(n.sup ? { sup: n.sup } : {}) };
  return hbox([layoutMathNode(opCluster, ctx), layoutMathNode(n.body, ctx)], [0, n.hideOp ? 0 : c.thinSpace]);
}

// ── Matrices ─────────────────────────────────────────────────────────────────

function matrixBox(rows: MathNode[][], ctx: MathLayoutCtx): MathBox {
  const c = ctx.font.constants(ctx.style.sizePx);
  const colGap = 2 * c.thickSpace;
  const rowGap = c.thinSpace * 3;
  if (rows.length === 0) return emptyBox();
  const cells = rows.map((r) => r.map((cell) => layoutMathNode(cell, ctx)));
  const nCols = Math.max(...cells.map((r) => r.length));
  if (nCols === 0) return emptyBox(); // all-empty rows → no negative width
  const colW: number[] = [];
  for (let j = 0; j < nCols; j++) colW[j] = Math.max(0, ...cells.map((r) => (r[j] ? r[j]!.width : 0)));
  const rowH = cells.map((r) => ({
    ascent: Math.max(0, ...r.map((b) => b.ascent)),
    descent: Math.max(0, ...r.map((b) => b.descent)),
  }));
  const totalH = rowH.reduce((s, h) => s + h.ascent + h.descent, 0) + rowGap * (rows.length - 1);
  const axis = c.axisHeight;
  // Center the grid on the math axis. The axis sits axisHeight ABOVE the baseline,
  // i.e. at y = -axisHeight (y grows downward), so the grid TOP is at
  // -axisHeight - totalH/2. (A +axis here put the grid below the baseline, leaving
  // the axis-centered fence brackets visibly too high.)
  let y = -(totalH / 2) - axis;
  const parts: MathBox[] = [];
  for (let i = 0; i < cells.length; i++) {
    const h = rowH[i]!;
    const baseline = y + h.ascent;
    let x = 0;
    for (let j = 0; j < nCols; j++) {
      const b = cells[i]![j];
      if (b) parts.push(shiftBox(slideX(center(b, colW[j]!), x), 0, baseline));
      x += colW[j]! + colGap;
    }
    y += h.ascent + h.descent + rowGap;
  }
  const width = colW.reduce((s, w) => s + w, 0) + colGap * (nCols - 1);
  const merged = overlayCentered(parts, "left");
  return { ...merged, width };
}

export { boxHeight };
