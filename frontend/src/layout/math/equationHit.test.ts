// hitTestEquation finds the display-equation block under a page point — the
// geometry that powers right-click → "Edit Equation…".
//
// geometry.ts creates a measuring <canvas> at module load, so we stub a minimal
// `document` (this repo's vitest runs in node, no jsdom) and dynamic-import after.

import { beforeAll, describe, expect, it } from "vitest";
import type { Document, EquationBlock } from "@cw/shared";

beforeAll(() => {
  (globalThis as unknown as { document: unknown }).document = {
    createElement: () => ({ getContext: () => ({ measureText: () => ({ width: 0 }), font: "" }) }),
  };
});

const FRAC = "<math><mfrac><mn>1</mn><mn>2</mn></mfrac></math>";

describe("hitTestEquation", () => {
  it("returns the equation blockId at its center and null off it", async () => {
    const { installMeasureHost } = await import("../../export/shared/measureHost");
    const { createLayoutEngine } = await import("../engine");
    const { hitTestEquation, hitTestSelectableObject, objectRect } = await import("../geometry");
    const { parseMathml } = await import("../../mathml/parse");
    await installMeasureHost();

    const block: EquationBlock = {
      kind: "equation", id: "eq", revision: 0, align: "center",
      equation: { ...parseMathml(FRAC), display: true },
    };
    const doc: Document = {
      section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [block],
    };
    const tree = createLayoutEngine().layout(doc);
    const placed = tree.pages.flatMap((p) => p.blocks).find((b) => b.blockId === "eq")!;
    const cx = placed.x + placed.equation!.width / 2;
    const cy = placed.y + placed.equation!.height / 2;

    expect(hitTestEquation(tree, 0, cx, cy)?.blockId).toBe("eq");
    expect(hitTestEquation(tree, 0, placed.x - 50, cy)).toBeNull();

    // Selectable like an image: the generic object hit-test + objectRect resolve it.
    expect(hitTestSelectableObject(tree, 0, cx, cy)?.blockId).toBe("eq");
    const rect = objectRect(tree, "eq");
    expect(rect).toBeTruthy();
    expect(rect!.width).toBeCloseTo(placed.equation!.width);
    expect(rect!.height).toBeCloseTo(placed.equation!.height);
  });

  it("inlineEquationAt returns the inline equation's character range", async () => {
    const { createLayoutEngine } = await import("../engine");
    const { inlineEquationAt } = await import("../geometry");
    const { parseMathml } = await import("../../mathml/parse");
    const CHAR = { fontFamily: "Times New Roman", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
    const PARA = { align: "left" as const, lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 8, indentFirstLinePx: 0, indentLeftPx: 0 };
    const doc = {
      section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [
        { kind: "paragraph" as const, id: "p", revision: 0, style: PARA, runs: [
          { text: "x = ", style: CHAR },
          { text: "￼", style: { ...CHAR, equation: { ...parseMathml(FRAC), display: false } } },
          { text: " ok", style: CHAR },
        ] },
      ],
    };
    const tree = createLayoutEngine().layout(doc as never);
    const block0 = tree.pages[0]!.blocks[0]!;
    const frag = block0.lines.flatMap((l) => l.fragments).find((f) => (f as { equation?: unknown }).equation)!;
    const cy = block0.y + block0.lines[0]!.y + 8;
    const fx = block0.x + frag.x;
    // The equation run is "x = " (4 chars) → offset 4..5. It must resolve across
    // the WHOLE box — both the left and right halves (the reported bug: only the
    // right half worked).
    for (const frac of [0.2, 0.5, 0.8]) {
      const hit = inlineEquationAt(tree, 0, fx + frag.width * frac, cy);
      expect(hit, `frac=${frac}`).toMatchObject({ blockId: "p", start: 4, end: 5 });
    }
    // And the run AT the sentinel index carries the equation (styleOfCharAt), while
    // the char before it (the space) does not.
    const { styleOfCharAt } = await import("@cw/shared");
    const runs = (doc.blocks[0] as { runs: { style: { equation?: unknown } }[] }).runs;
    expect(styleOfCharAt(runs as never, 4)?.equation).toBeTruthy();
    expect(styleOfCharAt(runs as never, 3)?.equation).toBeUndefined();
  });
});
