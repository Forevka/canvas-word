// Page-query tests. Import the canvas stub FIRST (chars × ½ font-size) so
// pagination is deterministic without a real font engine — same setup the layout
// engine tests use.
import "./test-canvas-setup";

import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps } from "@cw/shared";
import { getPages, indexOnPage, pageOfBlock } from "./pages";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" };
const PARA: ParaStyle = { align: "left", lineHeight: 1, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

const para = (id: string, text: string, extra: Partial<ParaStyle> = {}): Paragraph => ({
  kind: "paragraph",
  id,
  revision: 0,
  runs: [{ text, style: CHAR }],
  style: { ...PARA, ...extra },
});

// Force deterministic pagination with explicit page breaks.
const threePageDoc = (section: Partial<SectionProps> = {}): Document => ({
  section: { ...SECTION, ...section },
  blocks: [
    para("a", "page one"),
    para("b", "page two", { pageBreakBefore: true }),
    para("c", "page three", { pageBreakBefore: true }),
  ],
});

describe("getPages", () => {
  it("returns one PageInfo per page with the placed block ids", () => {
    const pages = getPages(threePageDoc());
    expect(pages).toHaveLength(3);
    expect(pages.map((p) => p.blockIds)).toEqual([["a"], ["b"], ["c"]]);
  });

  it("carries page geometry and 1-based numbering", () => {
    const pages = getPages(threePageDoc());
    expect(pages.map((p) => p.number)).toEqual([1, 2, 3]);
    expect(pages.map((p) => p.index)).toEqual([0, 1, 2]);
    expect(pages[0]).toMatchObject({ widthPx: 816, heightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } });
  });

  it("honors section pageNumberStart on the displayed number", () => {
    const pages = getPages(threePageDoc({ pageNumberStart: 5 }));
    expect(pages.map((p) => p.number)).toEqual([5, 6, 7]);
  });
});

describe("pageOfBlock", () => {
  it("returns the displayed page number a block first appears on", () => {
    const doc = threePageDoc();
    expect(pageOfBlock(doc, "a")).toBe(1);
    expect(pageOfBlock(doc, "c")).toBe(3);
  });

  it("returns null for a block that is not placed", () => {
    expect(pageOfBlock(threePageDoc(), "missing")).toBeNull();
  });
});

describe("indexOnPage", () => {
  it("reports a block's page index and order among that page's blocks", () => {
    const pages = getPages(threePageDoc());
    expect(indexOnPage(pages, "a")).toEqual({ index: 0, order: 0 });
    expect(indexOnPage(pages, "c")).toEqual({ index: 2, order: 0 });
  });

  it("finds the order for multiple blocks on the same page", () => {
    // Two blocks share page 1 (no break between a2 and a1).
    const pages = getPages({
      section: SECTION,
      blocks: [para("a1", "one"), para("a2", "two"), para("b", "next", { pageBreakBefore: true })],
    });
    expect(indexOnPage(pages, "a2")).toEqual({ index: 0, order: 1 });
    expect(indexOnPage(pages, "b")).toEqual({ index: 1, order: 0 });
  });

  it("returns null for an unplaced block", () => {
    expect(indexOnPage(getPages(threePageDoc()), "missing")).toBeNull();
  });
});
