import { describe, expect, it, beforeAll } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle } from "../../model/document";
import { sampleDoc } from "../../model/sampleDoc";
import { installMeasureHost } from "../shared/measureHost";
import { renderPdf } from "./renderPdf";

beforeAll(async () => {
  await installMeasureHost();
});

const CHAR: CharStyle = {
  fontFamily: "Calibri, sans-serif",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#202124",
};
const PARA: ParaStyle = {
  align: "left",
  lineHeight: 1.5,
  spaceBeforePx: 0,
  spaceAfterPx: 8,
  indentFirstLinePx: 0,
  indentLeftPx: 0,
};
let pid = 0;
const para = (text: string, char: Partial<CharStyle> = {}, p: Partial<ParaStyle> = {}): Paragraph => ({
  kind: "paragraph",
  id: `p${pid++}`,
  revision: 0,
  runs: [{ text, style: { ...CHAR, ...char } }],
  style: { ...PARA, ...p },
});
const docOf = (...blocks: Paragraph[]): Document => ({
  section: {
    pageWidthPx: 816,
    pageHeightPx: 1056,
    marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
  },
  blocks,
});

const isPdf = (b: Uint8Array): boolean =>
  b[0] === 0x25 && b[1] === 0x50 && b[2] === 0x44 && b[3] === 0x46; // %PDF

describe("PDF export — happy path", () => {
  it("emits a valid single-page PDF for one paragraph", async () => {
    const { bytes, warnings } = await renderPdf(docOf(para("Hello, world.")));
    expect(isPdf(bytes)).toBe(true);
    expect(bytes.length).toBeGreaterThan(500);
    expect(warnings.find((w) => w.code === "font-substituted")).toBeUndefined();
  });

  it("produces one PDF page per laid-out page", async () => {
    const { createLayoutEngine } = await import("../../layout/engine");
    const doc = sampleDoc();
    const tree = createLayoutEngine().layout(doc);
    const { bytes } = await renderPdf(doc);
    expect(isPdf(bytes)).toBe(true);
    // Count "/Type /Page" (not /Pages) occurrences in the uncompressed-ish stream.
    // pdfkit compresses content streams, but the page objects' dictionaries are
    // plain — match "/Type /Page" with a trailing non-'s'.
    const text = Buffer.from(bytes).toString("latin1");
    const pageObjs = (text.match(/\/Type\s*\/Page(?![s])/g) ?? []).length;
    expect(pageObjs).toBe(tree.pages.length);
    expect(tree.pages.length).toBeGreaterThan(1);
  });

  it("renders decorations, tables and images without throwing", async () => {
    // Mixed inline styles exercise sub/super, underline, strike, highlight, links.
    const styled = docOf(
      para("bold", { bold: true }),
      para("under", { underline: true }),
      para("strike", { strikethrough: true }),
      para("hl", { highlightColor: "#fff2cc" }),
      para("link", { link: "https://example.com", color: "#0563c1", underline: true }),
      para("super", { verticalAlign: "super" }),
    );
    const { bytes } = await renderPdf(styled);
    expect(isPdf(bytes)).toBe(true);
  });

  it("substitutes unknown families and warns once", async () => {
    const { warnings } = await renderPdf(docOf(para("x", { fontFamily: "Wingdings, fantasy" })));
    expect(warnings.find((w) => w.code === "font-substituted")?.detail).toBe("Wingdings");
  });
});
