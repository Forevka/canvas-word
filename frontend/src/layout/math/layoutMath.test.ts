// Typesetter geometry, verified with a DETERMINISTIC stub font (no canvas / no
// real metrics), so the assertions pin layout rules rather than glyph shapes:
// fractions stack a bar between a raised numerator and a lowered denominator,
// scripts sit above/below the baseline, radicals draw a vinculum, etc.

import { describe, expect, it } from "vitest";
import { parseMathml } from "../../mathml/parse";
import { defaultMathConstants, type MathFont } from "./mathFont";
import { layoutMathRow, type MathLayoutCtx } from "./layoutMath";
import { boxHeight } from "./mathBox";
import { styledMathText } from "./mathAlpha";

// Identifiers are remapped to the Mathematical Alphanumeric block (math-italic).
const mi = (ch: string) => styledMathText(ch, undefined);

const SIZE = 20;

/** Deterministic font: every glyph is 0.5em wide, 0.7em up, 0.2em down. Measured
 *  by code point (not UTF-16 units) so a single astral math-italic glyph counts
 *  as one, like a real font. */
const stubFont: MathFont = {
  family: "stub",
  measure: (text, sizePx) => ({
    width: [...text].length * sizePx * 0.5,
    ascent: sizePx * 0.7,
    descent: sizePx * 0.2,
    italicCorrection: 0,
  }),
  constants: defaultMathConstants,
};

const ctx: MathLayoutCtx = { font: stubFont, style: { sizePx: SIZE, cramped: false } };

const layout = (mathml: string) => layoutMathRow(parseMathml(mathml).root.children, ctx);
const glyph = (box: ReturnType<typeof layout>, text: string) => box.glyphs.find((g) => g.text === text);

describe("atoms and rows", () => {
  it("lays a single identifier as one glyph (remapped to math-italic)", () => {
    const b = layout("<math><mi>x</mi></math>");
    expect(b.glyphs).toHaveLength(1);
    // Remapped to U+1D465 MATHEMATICAL ITALIC SMALL X, drawn upright (no faux slant).
    expect(b.glyphs[0]).toMatchObject({ text: mi("x"), italic: false });
    expect(mi("x")).toBe("\u{1D465}");
  });

  it("inserts spacing around a binary operator", () => {
    const tight = layout("<math><mi>x</mi><mi>y</mi></math>").width;
    const spaced = layout("<math><mi>x</mi><mo>+</mo><mi>y</mi></math>").width;
    // x + y is wider than the bare glyph advances (medium space on both sides).
    expect(spaced).toBeGreaterThan(tight + SIZE * 0.5);
  });

  it("treats a leading minus as unary (no leading space)", () => {
    const unary = layout("<math><mo>-</mo><mi>x</mi></math>");
    const minus = glyph(unary, "-")!;
    expect(minus.x).toBeCloseTo(0);
  });
});

describe("fractions", () => {
  it("stacks a raised numerator over a lowered denominator with a bar", () => {
    const b = layout("<math><mfrac><mn>1</mn><mn>2</mn></mfrac></math>");
    const one = glyph(b, "1")!;
    const two = glyph(b, "2")!;
    expect(one.y).toBeLessThan(0); // numerator above baseline
    expect(two.y).toBeGreaterThan(0); // denominator below baseline
    expect(b.rules.length).toBeGreaterThanOrEqual(1); // the fraction bar
    expect(boxHeight(b)).toBeGreaterThan(SIZE); // taller than a single glyph
  });

  it("omits the bar when thickness is 0 (binomial)", () => {
    const b = layout('<math><mfrac linethickness="0"><mn>1</mn><mn>2</mn></mfrac></math>');
    expect(b.rules).toHaveLength(0);
  });
});

describe("scripts", () => {
  it("raises a superscript above the baseline and shrinks it", () => {
    const b = layout("<math><msup><mi>x</mi><mn>2</mn></msup></math>");
    const two = glyph(b, "2")!;
    expect(two.y).toBeLessThan(0);
    expect(two.sizePx).toBeLessThan(SIZE); // script size reduction
  });

  it("lowers a subscript below the baseline", () => {
    const b = layout("<math><msub><mi>a</mi><mi>n</mi></msub></math>");
    expect(glyph(b, mi("n"))!.y).toBeGreaterThan(0);
  });

  it("places both scripts on a sub-sup", () => {
    const b = layout("<math><msubsup><mi>x</mi><mn>0</mn><mn>1</mn></msubsup></math>");
    expect(glyph(b, "0")!.y).toBeGreaterThan(0);
    expect(glyph(b, "1")!.y).toBeLessThan(0);
  });

  it("in display style: integrals keep sub/sup; sums move them under/over", () => {
    const dctx: MathLayoutCtx = { font: stubFont, style: { sizePx: SIZE, cramped: false, display: true } };
    const layoutD = (m: string) => layoutMathRow(parseMathml(m).root.children, dctx);
    const intg = layoutD("<math><msubsup><mo>∫</mo><mn>0</mn><mn>1</mn></msubsup></math>");
    const sum = layoutD("<math><msubsup><mo>∑</mo><mn>0</mn><mn>1</mn></msubsup></math>");
    // The integral's lower bound stays a subscript (to the RIGHT of the operator);
    // the sum's lower bound is a limit, centered under it (further left).
    expect(glyph(intg, "0")!.x).toBeGreaterThan(glyph(sum, "0")!.x);
  });
});

describe("stretchy fences", () => {
  it("grows the delimiter glyphs to match tall content", () => {
    const small = layout("<math><mfenced open='(' close=')'><mi>x</mi></mfenced></math>");
    const tall = layout("<math><mfenced open='(' close=')'><mfrac><mn>1</mn><mn>2</mn></mfrac></mfenced></math>");
    const smallParen = glyph(small, "(")!;
    const tallParen = glyph(tall, "(")!;
    expect(smallParen).toBeTruthy();
    expect(tallParen).toBeTruthy();
    // The fraction is taller than a single glyph, so its parens are drawn larger.
    expect(tallParen.sizePx).toBeGreaterThan(smallParen.sizePx);
  });
});

describe("radicals", () => {
  it("draws a surd glyph and a vinculum over the radicand", () => {
    const b = layout("<math><msqrt><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow></msqrt></math>");
    expect(glyph(b, "√")).toBeTruthy();
    expect(b.rules.length).toBeGreaterThanOrEqual(1); // vinculum
    // radicand glyphs are shifted right of the surd
    expect(glyph(b, mi("x"))!.x).toBeGreaterThan(0);
  });
});

describe("composite / robustness", () => {
  it("lays out the quadratic formula without throwing and with positive size", () => {
    const quad =
      "<math><mi>x</mi><mo>=</mo><mfrac>" +
      "<mrow><mo>-</mo><mi>b</mi><mo>±</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup>" +
      "<mo>-</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt></mrow>" +
      "<mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>";
    const b = layout(quad);
    expect(b.width).toBeGreaterThan(0);
    expect(boxHeight(b)).toBeGreaterThan(SIZE);
    expect(glyph(b, "√")).toBeTruthy();
  });

  it("handles an n-ary sum", () => {
    const b = layout(
      '<math><mrow><munderover><mo>∑</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></munderover><mi>i</mi></mrow></math>',
    );
    expect(glyph(b, "∑")).toBeTruthy();
    expect(b.width).toBeGreaterThan(0);
  });
});
