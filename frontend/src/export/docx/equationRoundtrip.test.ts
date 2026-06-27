// A display equation must survive .docx: model → OMML (export) → model (import).
// Proves the m: namespace, m:oMathPara emit, and the OMML→MathML import all line up.

import { beforeAll, describe, expect, it } from "vitest";
import type { CharStyle, Document, EquationBlock, Paragraph, Run } from "@cw/shared";
import { runExport } from "../pipeline";
import { runImport } from "../../import/docx/pipeline";
import { installMeasureHost } from "../shared/measureHost";
import { parseMathml } from "../../mathml/parse";

beforeAll(async () => {
  await installMeasureHost();
});

const FRAC = "<math><mfrac><mrow><mi>a</mi><mo>+</mo><mn>1</mn></mrow><mn>2</mn></mfrac></math>";

const CHAR: CharStyle = {
  fontFamily: "Georgia, serif",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#000000",
};
const paraOf = (runs: Run[]): Paragraph => ({
  kind: "paragraph",
  id: "p1",
  revision: 0,
  runs,
  style: { align: "left", lineHeight: 1, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 },
});

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

  it("preserves munderover, a double-struck variant, and mspace through the full pipeline", async () => {
    const mathml =
      "<math display=\"block\">" +
      "<munderover><mi>X</mi><mrow><mi>n</mi><mo>=</mo><mn>1</mn></mrow><mi>∞</mi></munderover>" +
      "<mspace width=\"0.5em\"/>" +
      "<mi mathvariant=\"double-struck\">R</mi>" +
      "</math>";
    const doc: Document = {
      section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [
        { kind: "equation", id: "e1", revision: 0, align: "center", equation: parseMathml(mathml) } satisfies EquationBlock,
      ],
    };
    const { bytes } = await runExport(doc, "docx");
    const { doc: back } = runImport(bytes);
    const eq = back.blocks.find((b): b is EquationBlock => b.kind === "equation");
    expect(eq).toBeTruthy();
    const json = JSON.stringify(eq!.equation.root);
    // munderover came back as ONE limit with both under and over (not left nested).
    const limit = eq!.equation.root.children.find((n) => n.type === "limit") as
      | { type: "limit"; under?: unknown; over?: unknown }
      | undefined;
    expect(limit?.under).toBeTruthy();
    expect(limit?.over).toBeTruthy();
    // double-struck variant survived via m:scr.
    expect(json).toContain("\"double-struck\"");
    // mspace recovered as a space node.
    expect(eq!.equation.root.children.some((n) => n.type === "space")).toBe(true);
  });

  it("keeps a hyperlink on an inline equation through the docx round-trip", async () => {
    const url = "https://example.com/math";
    const doc: Document = {
      section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [
        paraOf([
          { text: "see ", style: CHAR },
          { text: "￼", style: { ...CHAR, equation: { ...parseMathml(FRAC), display: false }, link: url } },
        ]),
      ],
    };
    const { bytes } = await runExport(doc, "docx");
    const { doc: back } = runImport(bytes);
    const para = back.blocks.find((b): b is Paragraph => b.kind === "paragraph");
    const eqRun = para?.runs.find((r) => r.style.equation);
    expect(eqRun).toBeTruthy();
    expect(eqRun!.style.link).toBe(url);
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
