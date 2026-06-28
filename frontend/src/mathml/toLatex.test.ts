// AST -> LaTeX, and the editor's both-way switch: MathML view <-> LaTeX view must
// preserve the equation (the bug being fixed: switching corrupted it).

import { describe, expect, it } from "vitest";
import { mathToLatex } from "./toLatex";
import { latexToMath } from "./latex";
import { parseMathml } from "./parse";
import { serializeMathml } from "./serialize";

const fromLatex = (tex: string) => latexToMath(tex);

describe("mathToLatex", () => {
  it("serializes the common constructs", () => {
    expect(mathToLatex(fromLatex("\\frac{1}{2}"))).toBe("\\frac{1}{2}");
    expect(mathToLatex(fromLatex("x^{2}"))).toContain("x^{2}");
    expect(mathToLatex(fromLatex("\\sqrt{x}"))).toBe("\\sqrt{x}");
    expect(mathToLatex(fromLatex("\\sqrt[3]{x}"))).toBe("\\sqrt[3]{x}");
    expect(mathToLatex(fromLatex("\\alpha"))).toBe("\\alpha");
    expect(mathToLatex(fromLatex("a \\leq b"))).toContain("\\leq");
    expect(mathToLatex(fromLatex("\\mathbb{R}"))).toBe("\\mathbb{R}");
  });

  it("round-trips delimiters and matrices", () => {
    expect(mathToLatex(fromLatex("\\left(x\\right)"))).toContain("\\left(");
    const m = mathToLatex(fromLatex("\\begin{bmatrix} 1 & 0 \\\\ 0 & 1 \\end{bmatrix}"));
    expect(m).toContain("\\begin{bmatrix}");
    expect(m).toContain("&");
    expect(m).toContain("\\\\");
  });

  it("round-trips \\begin{cases} as cases (not Bmatrix)", () => {
    const m = mathToLatex(fromLatex("\\begin{cases} a & x>0 \\\\ b & x<0 \\end{cases}"));
    expect(m).toContain("\\begin{cases}");
    expect(m).not.toContain("Bmatrix");
  });

  it("preserves whitespace inside \\text (no global collapse)", () => {
    expect(mathToLatex(fromLatex("\\text{a  b}"))).toBe("\\text{a  b}");
    expect(mathToLatex(parseMathml("<math><mtext>hello world</mtext></math>").root)).toBe("\\text{hello world}");
  });

  it("escapes LaTeX special chars inside \\text{…}", () => {
    const tex = mathToLatex(parseMathml("<math><mtext>a_b &amp; 50% #1</mtext></math>").root);
    expect(tex).toBe("\\text{a\\_b \\& 50\\% \\#1}");
    // Backslash / tilde / caret use control words (no single-char text escape).
    expect(mathToLatex(parseMathml("<math><mtext>a\\b</mtext></math>").root)).toBe(
      "\\text{a\\textbackslash{}b}",
    );
    expect(mathToLatex(parseMathml("<math><mtext>x^2~y</mtext></math>").root)).toBe(
      "\\text{x\\textasciicircum{}2\\textasciitilde{}y}",
    );
  });

  it("round-trips \\text{…} special chars through LaTeX (escape ∘ unescape == id)", () => {
    for (const raw of ["a_b", "50% off", "C# & F#", "a\\b", "x^2", "~tilde~", "{braces}", "$math$", "100% \\ {x}"]) {
      const node = { type: "text" as const, text: raw };
      const round = latexToMath(mathToLatex({ type: "row", children: [node] }));
      // Compare the WHOLE row — a leaked sibling (bad brace capture/unescape)
      // would slip past a first-child-only check.
      expect(round.children).toEqual([{ type: "text", text: raw }]);
    }
  });

  it("maps every imported mathvariant to a style command (no plain-identifier fallback)", () => {
    // fromOmml can emit these compound variants; the LaTeX view must keep a style
    // command rather than dropping back to an unstyled identifier.
    const variants: [string, string][] = [
      ["double-struck", "mathbb"],
      ["script", "mathcal"],
      ["fraktur", "mathfrak"],
      ["bold-fraktur", "mathfrak"],
      ["bold-script", "mathcal"],
      ["bold-sans-serif", "mathsf"],
      ["sans-serif-italic", "mathsf"],
      ["sans-serif-bold-italic", "mathsf"],
    ];
    for (const [variant, cmd] of variants) {
      const tex = mathToLatex(parseMathml(`<math><mi mathvariant="${variant}">X</mi></math>`).root);
      expect(tex).toBe(`\\${cmd}{X}`);
    }
  });
});

describe("editor view switch preserves the equation", () => {
  // MathML view -> LaTeX view -> back to MathML view must keep the structure.
  const cases = [
    "<math><mfrac><mrow><mi>a</mi><mo>+</mo><mn>1</mn></mrow><mn>2</mn></mfrac></math>",
    "<math><msup><mi>x</mi><mn>2</mn></msup></math>",
    "<math><msqrt><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow></msqrt></math>",
    "<math><munderover><mo>∑</mo><mrow><mi>n</mi><mo>=</mo><mn>1</mn></mrow><mo>∞</mo></munderover></math>",
  ];
  for (const mml of cases) {
    it(`survives mml→latex→mml for ${mml.slice(6, 26)}…`, () => {
      const root1 = parseMathml(mml).root;
      const latex = mathToLatex(root1); // user switches to LaTeX view
      const root2 = latexToMath(latex); // user switches back to MathML view
      const out = serializeMathml({ root: root2, display: false });
      // A failed conversion would leak a command name as literal text — assert none.
      for (const broken of ["overset", "underset", "frac", "sqrt", "sum", "alpha"]) {
        expect(out).not.toContain(`<mtext>${broken}`);
        expect(out).not.toContain(`<mi>${broken[0]}</mi><mi>${broken[1]}</mi>`); // \frac → f,r,a,c
      }
      // It never collapses to an empty equation.
      expect(out).not.toBe(serializeMathml({ root: { type: "row", children: [] }, display: false }));
      // And the operator/structure is still present.
      expect(out.length).toBeGreaterThan(40);
    });
  }
});
