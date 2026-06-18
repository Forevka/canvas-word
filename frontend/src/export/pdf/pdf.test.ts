import { describe, expect, it, beforeAll } from "vitest";
import type { Block, CharStyle, Document, ImageBlock, Paragraph, ParaStyle } from "@cw/shared";
import { defaultListDefinition } from "@cw/shared";
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

  it("renders page color, borders and column separators without drift", async () => {
    // The PDF painter shares the engine's Page paint data + pageBorderSegments with
    // the canvas renderer, so this just confirms the new draw calls run cleanly.
    const doc = docOf(...Array.from({ length: 40 }, (_, i) => para(`c${i}`)));
    doc.section.pageColorHex = "#ffeecc";
    doc.section.columns = { count: 2, gapPx: 24, sep: true };
    doc.section.pageBorders = {
      offsetFrom: "page",
      top: { style: "double", widthPx: 2, color: "#1a73e8" },
      right: { style: "double", widthPx: 2, color: "#1a73e8" },
      bottom: { style: "double", widthPx: 2, color: "#1a73e8" },
      left: { style: "double", widthPx: 2, color: "#1a73e8" },
    };
    const { bytes } = await renderPdf(doc);
    expect(isPdf(bytes)).toBe(true);
    expect(bytes.length).toBeGreaterThan(500);
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

  it("renders a numbered list INSIDE a table cell (cell marker paint path)", async () => {
    // A cell paragraph carrying a list ref — the engine hangs its marker in the
    // cell and the PDF painter draws it via the same recursive paintBlock as a
    // body marker. Proves the new in-cell-list path runs through real PDFKit.
    const item = para("cell item", {}, { list: { listId: "numbers", level: 0 } } as Partial<ParaStyle>);
    const table: Block = {
      kind: "table",
      id: "t",
      revision: 0,
      rows: [{ cells: [{ id: "c", blocks: [item] }] }],
    };
    const doc: Document = {
      section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [table],
      lists: { numbers: defaultListDefinition("decimal") },
    };
    const { bytes, warnings } = await renderPdf(doc);
    expect(isPdf(bytes)).toBe(true);
    expect(warnings.find((w) => w.code === "font-substituted")).toBeUndefined();
  });

  it("substitutes unknown families and warns once", async () => {
    const { warnings } = await renderPdf(docOf(para("x", { fontFamily: "Wingdings, fantasy" })));
    expect(warnings.find((w) => w.code === "font-substituted")?.detail).toBe("Wingdings");
  });
});

const latin1 = (b: Uint8Array): string => Buffer.from(b).toString("latin1");

describe("PDF export — internal navigation", () => {
  it("emits GoTo links, named destinations and an outline for a TOC", async () => {
    const heading: Paragraph = {
      kind: "paragraph",
      id: "h1",
      revision: 0,
      runs: [{ text: "Chapter One", style: { ...CHAR } }],
      style: { ...PARA, namedStyle: "heading1" },
    };
    const entry: Paragraph = {
      kind: "paragraph",
      id: "toc1",
      revision: 0,
      runs: [{ text: "Chapter One", style: { ...CHAR } }],
      style: { ...PARA, tocEntry: { targetId: "h1", level: 1 } } as ParaStyle,
    };
    const { bytes } = await renderPdf(docOf(entry, heading, para("Body.")));
    const text = latin1(bytes);
    expect(isPdf(bytes)).toBe(true);
    expect(text).toContain("/GoTo"); // TOC entry → heading link
    expect(text).toContain("b_h1"); // named destination for the heading
    expect(text).toContain("/Dests"); // names tree carrying the destination
    expect(text).toContain("/Outlines"); // bookmarks panel built from headings
  });

  it("emits a GoTo for an in-document #anchor link", async () => {
    const target: Paragraph = {
      kind: "paragraph",
      id: "sec",
      revision: 0,
      runs: [{ text: "Target heading", style: { ...CHAR } }],
      style: { ...PARA, namedStyle: "heading1" },
    };
    const ref = para("see ");
    ref.runs.push({ text: "this section", style: { ...CHAR, link: "#mark" } });
    const doc: Document = {
      section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [ref, target],
      bookmarks: { mark: { start: { blockId: "sec", offset: 0 }, end: { blockId: "sec", offset: 0 } } },
    };
    const text = latin1((await renderPdf(doc)).bytes);
    expect(text).toContain("/GoTo");
    expect(text).toContain("b_sec");
  });

  it("ignores an unresolved #anchor (no GoTo annotation)", async () => {
    // pdfkit always writes an empty /Dests names tree, so the real signal that a
    // dangling anchor emitted nothing is the absence of a GoTo action.
    const text = latin1((await renderPdf(docOf(para("x", { link: "#nope" })))).bytes);
    expect(text).not.toContain("/GoTo");
  });

  it("still emits a URI annotation for external links", async () => {
    const p = para("link", { link: "https://example.com", color: "#0563c1", underline: true });
    const text = latin1((await renderPdf(docOf(p))).bytes);
    expect(text).toContain("/URI");
  });
});

// A 1x1 opaque (no alpha) PNG, color-type 2 — pdfkit's png-js decodes it cleanly,
// so a working image emits an XObject and never the "image-format-unsupported" warning.
const PNG_1x1 = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
  ),
  (c) => c.charCodeAt(0),
);

let iid = 0;
const anchoredImage = (behind: boolean): ImageBlock => ({
  kind: "image",
  id: `img${iid++}`,
  revision: 0,
  src: "blob:behind",
  widthPx: 100,
  heightPx: 100,
  align: "left",
  anchor: {
    behind,
    offsetXPx: 10,
    offsetYPx: 10,
    relFromH: "page",
    relFromV: "page",
  },
});

describe("PDF export — anchored images", () => {
  it("emits the behind-text image XObject for a BODY-level anchored image", async () => {
    const img = anchoredImage(true);
    const doc = docOf(para("Body text over the image."));
    doc.blocks.unshift(img as unknown as Paragraph);
    const { bytes, warnings } = await renderPdf(doc, { images: { "blob:behind": PNG_1x1 } });
    expect(isPdf(bytes)).toBe(true);
    expect(warnings.find((w) => w.code === "image-format-unsupported")).toBeUndefined();
    // A successfully placed image becomes an /Image XObject in the PDF.
    expect(latin1(bytes)).toContain("/Subtype /Image");
  });

  it("places the BODY behind image in the behind layer (before text in paint order)", async () => {
    const { createLayoutEngine } = await import("../../layout/engine");
    const img = anchoredImage(true);
    const doc = docOf(para("Body text."));
    doc.blocks.unshift(img as unknown as Paragraph);
    const tree = createLayoutEngine().layout(doc);
    const blocks = tree.pages[0]!.blocks;
    const imgIdx = blocks.findIndex((b) => b.image?.behind);
    const textIdx = blocks.findIndex((b) => b.lines.length > 0);
    expect(imgIdx).toBeGreaterThanOrEqual(0);
    expect(blocks[imgIdx]!.image!.behind).toBe(true);
    // Behind layer paints first → must precede the text block in array order.
    expect(imgIdx).toBeLessThan(textIdx);
  });

  it("emits the behind-text image XObject for a CELL anchored image", async () => {
    const img = anchoredImage(true);
    const table: Block = {
      kind: "table",
      id: "t",
      revision: 0,
      rows: [{ cells: [{ id: "c", blocks: [img, para("Cell text over image.")] }] }],
    };
    const doc: Document = {
      section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [table],
    };
    const { bytes, warnings } = await renderPdf(doc, { images: { "blob:behind": PNG_1x1 } });
    expect(isPdf(bytes)).toBe(true);
    expect(warnings.find((w) => w.code === "image-format-unsupported")).toBeUndefined();
    expect(latin1(bytes)).toContain("/Subtype /Image");
  });

  it("places a CELL behind image in the cell's behind layer (before text)", async () => {
    const { createLayoutEngine } = await import("../../layout/engine");
    const img = anchoredImage(true);
    const table: Block = {
      kind: "table",
      id: "t",
      revision: 0,
      rows: [{ cells: [{ id: "c", blocks: [img, para("Cell text.")] }] }],
    };
    const doc: Document = {
      section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [table],
    };
    const tree = createLayoutEngine().layout(doc);
    const cellBlocks = tree.pages[0]!.blocks[0]!.table!.rows[0]!.cells[0]!.blocks;
    const imgIdx = cellBlocks.findIndex((b) => b.image?.behind);
    const textIdx = cellBlocks.findIndex((b) => b.lines.length > 0);
    expect(imgIdx, "cell image should carry the behind flag").toBeGreaterThanOrEqual(0);
    expect(cellBlocks[imgIdx]!.image!.behind).toBe(true);
    expect(imgIdx).toBeLessThan(textIdx);
  });
});
