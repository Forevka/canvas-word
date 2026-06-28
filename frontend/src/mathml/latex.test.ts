// LaTeX → MathML AST parsing. We round-trip through serialize to assert structure
// compactly, and check a few ASTs directly.

import { describe, expect, it } from "vitest";
import { latexToMath } from "./latex";
import { serializeMathml } from "./serialize";
import { mathmlToOmml } from "./toOmml";

const mml = (latex: string): string => serializeMathml({ root: latexToMath(latex), display: false });
const omml = (latex: string): string => mathmlToOmml({ root: latexToMath(latex), display: false });

describe("latexToMath", () => {
  it("parses a fraction", () => {
    expect(mml("\\frac{1}{2}")).toContain("<mfrac><mn>1</mn><mn>2</mn></mfrac>");
  });

  it("parses superscripts and subscripts", () => {
    expect(mml("x^2")).toContain("<msup><mi>x</mi><mn>2</mn></msup>");
    expect(mml("a_n")).toContain("<msub><mi>a</mi><mi>n</mi></msub>");
    expect(mml("x_i^2")).toContain("<msubsup>");
  });

  it("groups braced arguments", () => {
    expect(mml("x^{2n}")).toContain("<msup><mi>x</mi><mrow><mn>2</mn><mi>n</mi></mrow></msup>");
  });

  it("parses roots", () => {
    expect(mml("\\sqrt{x}")).toContain("<msqrt><mi>x</mi></msqrt>");
    expect(mml("\\sqrt[3]{x}")).toContain("<mroot><mi>x</mi><mn>3</mn></mroot>");
  });

  it("maps Greek and operators", () => {
    expect(mml("\\alpha + \\beta")).toContain("<mi>α</mi>");
    expect(mml("a \\leq b")).toContain("<mo>≤</mo>");
    expect(mml("\\pi")).toContain("<mi>π</mi>");
  });

  it("parses \\left … \\right fences", () => {
    const out = mml("\\left(\\frac{1}{2}\\right)");
    expect(out).toContain("<mfenced");
    expect(out).toContain('open="("');
    expect(out).toContain('close=")"');
  });

  it("parses a bmatrix into an mtable", () => {
    const out = mml("\\begin{bmatrix} 1 & 0 \\\\ 0 & 1 \\end{bmatrix}");
    expect(out).toContain("<mtable>");
    expect(out).toContain("<mtr>");
    expect(out).toContain('open="["');
  });

  it("maps function names to upright identifiers", () => {
    const root = latexToMath("\\sin x");
    expect(root.children[0]).toMatchObject({ type: "ident", text: "sin", variant: "normal" });
  });

  it("applies \\mathbb variant", () => {
    const root = latexToMath("\\mathbb{R}");
    expect(root.children[0]).toMatchObject({ type: "ident", text: "R", variant: "double-struck" });
  });

  it("parses accents", () => {
    expect(mml("\\hat{x}")).toContain("<mover");
    expect(mml("\\vec{v}")).toContain("→");
  });

  it("handles the quadratic formula end to end", () => {
    const out = mml("x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}");
    expect(out).toContain("<mfrac>");
    expect(out).toContain("<msqrt>");
    expect(out).toContain("<msup><mi>b</mi><mn>2</mn></msup>");
    expect(out).toContain("±");
  });

  it("is lenient: unknown commands become text, no throw", () => {
    expect(() => latexToMath("\\foobar x")).not.toThrow();
  });

  it("preserves whitespace inside \\text / \\mbox / \\operatorname", () => {
    expect(latexToMath("\\text{hello world}").children[0]).toMatchObject({ type: "text", text: "hello world" });
    expect(latexToMath("\\mbox{a   b}").children[0]).toMatchObject({ type: "text", text: "a   b" });
    expect(latexToMath("\\operatorname{arg max}").children[0]).toMatchObject({ type: "text", text: "arg max" });
    // The text group keeps spaces even when surrounded by math.
    expect(mml("a + \\text{is positive}")).toContain("<mtext>is positive</mtext>");
  });

  it("balances nested braces in a \\text argument", () => {
    expect(latexToMath("\\text{f(x) = {y}}").children[0]).toMatchObject({ type: "text", text: "f(x) = {y}" });
  });

  // ── #19: big operators become n-ary objects (real OMML m:nary) ──────────────
  it("parses \\sum / \\int as n-ary objects, not plain scripted operators", () => {
    const sum = latexToMath("\\sum_{i=1}^{n}").children[0];
    expect(sum).toMatchObject({
      type: "nary",
      op: "∑",
      sub: { children: [{ type: "ident", text: "i" }, { type: "op", text: "=" }, { type: "number", text: "1" }] },
      sup: { type: "ident", text: "n" },
    });
    expect(latexToMath("\\int_0^1").children[0]).toMatchObject({ type: "nary", op: "∫" });
    // A bare big operator with no bounds is still n-ary.
    expect(latexToMath("\\prod").children[0]).toMatchObject({ type: "nary", op: "∏" });
  });

  it("exports LaTeX big operators as OMML m:nary with the operator char", () => {
    const out = omml("\\sum_{i=1}^{n}");
    expect(out).toContain("<m:nary>");
    expect(out).toContain('<m:chr m:val="∑"/>');
    expect(out).toContain("<m:sub>");
    expect(out).toContain("<m:sup>");
    // Not serialized as an ordinary scripted run.
    expect(out).not.toContain("<m:sSubSup>");
  });

  // ── #21(b): styled numeric glyphs survive ───────────────────────────────────
  it("keeps the variant on styled numbers (\\mathbb{1}, \\mathtt{0})", () => {
    expect(latexToMath("\\mathbb{1}").children[0]).toMatchObject({ type: "number", text: "1", variant: "double-struck" });
    expect(latexToMath("\\mathtt{0}").children[0]).toMatchObject({ type: "number", text: "0", variant: "monospace" });
    expect(mml("\\mathbb{1}")).toContain('<mn mathvariant="double-struck">1</mn>');
  });

  // ── #21(a): \phantom keeps its invisible-spacing wrapper ─────────────────────
  it("parses \\phantom as a phantom node", () => {
    expect(latexToMath("\\phantom{xy}").children[0]).toMatchObject({ type: "phantom" });
    expect(omml("\\phantom{x}")).toContain("<m:phant>");
  });
});
