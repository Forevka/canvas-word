// End-to-end: a display EquationBlock goes through the real layout engine (with
// the headless measurement host) and comes out as a placed equation carrying a
// typeset box of glyphs + rules. Guards the model -> layout -> geometry spine.

import { beforeAll, describe, expect, it } from "vitest";
import type { Document, EquationBlock } from "@cw/shared";
import { installMeasureHost } from "../../export/shared/measureHost";
import { createLayoutEngine } from "../engine";
import { parseMathml } from "../../mathml/parse";
import { latexToMath } from "../../mathml/latex";

beforeAll(async () => {
  await installMeasureHost();
});

const QUAD =
  "<math><mi>x</mi><mo>=</mo><mfrac><mrow><mo>-</mo><mi>b</mi><mo>±</mo>" +
  "<msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>-</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt>" +
  "</mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>";

const docWith = (block: EquationBlock): Document => ({
  section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
  blocks: [block],
});

describe("display equation block layout", () => {
  it("places a centered equation with a typeset box of glyphs and rules", () => {
    const block: EquationBlock = {
      kind: "equation",
      id: "e1",
      revision: 0,
      align: "center",
      equation: { ...parseMathml(QUAD), display: true },
    };
    const tree = createLayoutEngine().layout(docWith(block));
    const placed = tree.pages.flatMap((p) => p.blocks).find((b) => b.blockId === "e1");
    expect(placed?.equation).toBeTruthy();
    const eq = placed!.equation!;
    expect(eq.width).toBeGreaterThan(0);
    expect(eq.height).toBeGreaterThan(0);
    expect(eq.box.glyphs.length).toBeGreaterThan(5); // x = -b ± √ ... 2a
    expect(eq.box.rules.length).toBeGreaterThanOrEqual(2); // fraction bar + radical vinculum
    // The √ surd and the variable b (remapped to math-italic) both appear.
    expect(eq.box.glyphs.some((g) => g.text === "√")).toBe(true);
    expect(eq.box.glyphs.some((g) => g.text === "\u{1D44F}")).toBe(true); // 𝑏
  });

  it("typesets a LaTeX-authored quadratic with a fraction bar + radical", () => {
    const root = latexToMath("x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}");
    const block: EquationBlock = { kind: "equation", id: "lx", revision: 0, align: "center", equation: { root, display: true } };
    const tree = createLayoutEngine().layout(docWith(block));
    const eq = tree.pages.flatMap((p) => p.blocks).find((b) => b.blockId === "lx")!.equation!;
    expect(eq.box.rules.length).toBeGreaterThanOrEqual(2); // fraction bar + radical vinculum
    expect(eq.box.glyphs.some((g) => g.text === "√")).toBe(true);
    expect(eq.box.glyphs.some((g) => g.text === "±")).toBe(true);
    expect(eq.height).toBeGreaterThan(40); // stacked (display 22px), not a flat line
  });

  it("vertically centers matrix brackets with the matrix content", () => {
    const root = latexToMath("\\begin{bmatrix} 1 & 0 \\\\ 0 & 1 \\end{bmatrix}");
    const block: EquationBlock = { kind: "equation", id: "mx", revision: 0, align: "center", equation: { root, display: true } };
    const tree = createLayoutEngine().layout(docWith(block));
    const eq = tree.pages.flatMap((p) => p.blocks).find((b) => b.blockId === "mx")!.equation!;
    const bracket = eq.box.glyphs.find((g) => g.text === "[")!;
    const nums = eq.box.glyphs.filter((g) => g.text === "1" || g.text === "0");
    expect(bracket).toBeTruthy();
    expect(nums.length).toBe(4);
    const numCenter = (Math.min(...nums.map((g) => g.y)) + Math.max(...nums.map((g) => g.y))) / 2;
    // Bracket visual center (STIX 0.762 asc / 0.238 desc of its own size).
    const bracketCenter = bracket.y + ((0.238 - 0.762) / 2) * bracket.sizePx;
    // The brackets must hug the content, not float above it (the bug): centers
    // within a fraction of the bracket size, and the content sits inside the bracket.
    expect(Math.abs(bracketCenter - numCenter)).toBeLessThan(bracket.sizePx * 0.3);
    const bTop = bracket.y - 0.762 * bracket.sizePx;
    const bBot = bracket.y + 0.238 * bracket.sizePx;
    for (const g of nums) expect(g.y).toBeGreaterThan(bTop), expect(g.y).toBeLessThan(bBot);
  });

  it("centers within the content width", () => {
    const block: EquationBlock = {
      kind: "equation", id: "e2", revision: 0, align: "center",
      equation: { ...parseMathml("<math><mi>e</mi></math>"), display: true },
    };
    const tree = createLayoutEngine().layout(docWith(block));
    const placed = tree.pages.flatMap((p) => p.blocks).find((b) => b.blockId === "e2")!;
    // content box is 816 - 96*2 = 624; a tiny glyph centers well right of the margin.
    expect(placed.x).toBeGreaterThan(96 + 100);
  });
});
