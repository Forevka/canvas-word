// A display equation must survive .docx: model → OMML (export) → model (import).
// Proves the m: namespace, m:oMathPara emit, and the OMML→MathML import all line up.

import { beforeAll, describe, expect, it } from "vitest";
import type { Document, EquationBlock } from "@cw/shared";
import { runExport } from "../pipeline";
import { runImport } from "../../import/docx/pipeline";
import { installMeasureHost } from "../shared/measureHost";
import { parseMathml } from "../../mathml/parse";

beforeAll(async () => {
  await installMeasureHost();
});

const FRAC = "<math><mfrac><mrow><mi>a</mi><mo>+</mo><mn>1</mn></mrow><mn>2</mn></mfrac></math>";

describe("equation docx round-trip", () => {
  it("preserves a display equation through export then import", async () => {
    const doc: Document = {
      section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [
        { kind: "equation", id: "e1", revision: 0, align: "center", equation: { ...parseMathml(FRAC), display: true } } satisfies EquationBlock,
      ],
    };
    const { bytes } = await runExport(doc, "docx");
    const xml = new TextDecoder().decode(bytes);
    void xml; // (bytes are a zip; the assertion below re-imports instead of grepping)

    const { doc: back } = runImport(bytes);
    const eq = back.blocks.find((b): b is EquationBlock => b.kind === "equation");
    expect(eq).toBeTruthy();
    const root = eq!.equation.root;
    // Top level should be the fraction (collapsed from the single m:oMath child).
    const frac = root.children.find((n) => n.type === "frac");
    expect(frac).toBeTruthy();
    if (frac && frac.type === "frac") {
      // numerator carries a + b structure, denominator is 2
      expect(JSON.stringify(frac.num)).toContain("\"+\"");
      expect(JSON.stringify(frac.den)).toContain("\"2\"");
    }
  });

  it("preserves a right-aligned display equation through export then import", async () => {
    const doc: Document = {
      section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [
        { kind: "equation", id: "e1", revision: 0, align: "right", equation: { ...parseMathml(FRAC), display: true } } satisfies EquationBlock,
      ],
    };
    const { bytes } = await runExport(doc, "docx");
    const { doc: back } = runImport(bytes);
    const eq = back.blocks.find((b): b is EquationBlock => b.kind === "equation");
    expect(eq?.align).toBe("right");
  });

  it("exports an equation to PDF without crashing and emits real bytes", async () => {
    const doc: Document = {
      section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [
        { kind: "equation", id: "e1", revision: 0, align: "center", equation: { ...parseMathml(FRAC), display: true } } satisfies EquationBlock,
      ],
    };
    const { bytes } = await runExport(doc, "pdf");
    expect(bytes.length).toBeGreaterThan(1000);
    expect(new TextDecoder().decode(bytes.slice(0, 5))).toBe("%PDF-");
  });
});
