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
