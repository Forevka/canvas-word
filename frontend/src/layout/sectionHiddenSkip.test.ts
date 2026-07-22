// Regression: a fully-hidden paragraph (w:vanish anchor — e.g. a bookmark) that
// sits BEFORE a Next Page section break must not shift the break. The measure
// pass drops hidden-layout-skip paragraphs, so the measured[] array is shorter
// than doc.blocks; the placement walk must advance sections by each entry's DOC
// index, not its measured index, or the break fires one block late and strands
// the new section's first block (a TOC title) alone on its own page.
//
// Real-world trigger: RenderedReportsDocx TOC — a hidden anchor at blocks[0]
// pushed the "TABLE OF CONTENTS" heading onto a page by itself, with the entry
// list on the next page. In Word the heading and list share a page.
import "./test-canvas-setup";

import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps } from "@cw/shared";
import { createLayoutEngine } from "./engine";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" };
const PARA: ParaStyle = { align: "left", lineHeight: 1, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let n = 0;
const para = (text: string, style: Partial<ParaStyle> = {}, hidden = false): Paragraph => ({
  kind: "paragraph", id: `b${n++}`, revision: 0,
  runs: [{ text, style: { ...CHAR, hidden } }],
  style: { ...PARA, ...style },
});

const layout = (blocks: Paragraph[]): ReturnType<ReturnType<typeof createLayoutEngine>["layout"]> =>
  createLayoutEngine().layout({ section: SECTION, blocks } as Document);
const pageOf = (tree: ReturnType<typeof layout>, id: string): number => {
  for (const pg of tree.pages) if (pg.blocks.some((b) => b.blockId === id)) return pg.index;
  return -1;
};

describe("hidden-skip paragraph before a section break", () => {
  it("keeps the new section's title on the same page as its body (not stranded alone)", () => {
    const anchor = para("bookmark-anchor", {}, true); // fully hidden → dropped from measured[]
    const cover = para("cover text", { sectionBreak: { type: "nextPage", props: {} } } as Partial<ParaStyle>);
    const title = para("TABLE OF CONTENTS", { pageBreakBefore: true }); // section-1 first block
    const listItem = para("Chapter 1 .... 1");
    const tree = layout([anchor, cover, title, listItem]);

    const titlePage = pageOf(tree, title.id);
    const listPage = pageOf(tree, listItem.id);
    expect(titlePage).toBeGreaterThan(pageOf(tree, cover.id)); // TOC section is on a fresh page
    expect(listPage).toBe(titlePage); // …and its list rides along with the heading
  });

  it("still pages correctly with NO hidden paragraph (control)", () => {
    const cover = para("cover text", { sectionBreak: { type: "nextPage", props: {} } } as Partial<ParaStyle>);
    const title = para("TABLE OF CONTENTS", { pageBreakBefore: true });
    const listItem = para("Chapter 1 .... 1");
    const tree = layout([cover, title, listItem]);
    expect(pageOf(tree, listItem.id)).toBe(pageOf(tree, title.id));
    expect(pageOf(tree, title.id)).toBeGreaterThan(pageOf(tree, cover.id));
  });
});
