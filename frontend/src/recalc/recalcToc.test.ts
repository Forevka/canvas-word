// The headless TOC recalc core: a TOC entry ("label\t<number>") whose run carries
// an in-document link resolves through doc.bookmarks to a target block, and the
// trailing number is rewritten to that block's real layout page. Requires the
// fontkit measure host (layout measures text).

import { beforeAll, describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps } from "@cw/shared";
import { textOfRuns } from "@cw/shared";
import { installMeasureHost } from "../export/shared/measureHost";
import { computeTocEditsWithLayout, recalcToc } from "./recalcToc";

beforeAll(async () => {
  await installMeasureHost();
});

const SECTION: SectionProps = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
};
const CHAR: CharStyle = {
  fontFamily: "Georgia",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#000",
};
const PARA: ParaStyle = {
  align: "left",
  lineHeight: 1.2,
  spaceBeforePx: 0,
  spaceAfterPx: 0,
  indentFirstLinePx: 0,
  indentLeftPx: 0,
};

/** A TOC entry: "<label>" (linked to #anchor) + "\t<staleNum>". */
const tocEntry = (id: string, label: string, anchor: string, staleNum: string): Paragraph => ({
  kind: "paragraph",
  id,
  revision: 0,
  runs: [
    { text: label, style: { ...CHAR, link: `#${anchor}` } },
    { text: `\t${staleNum}`, style: CHAR },
  ],
  style: PARA,
});
const heading = (id: string, text: string): Paragraph => ({
  kind: "paragraph",
  id,
  revision: 0,
  runs: [{ text, style: { ...CHAR, bold: true } }],
  style: { ...PARA, namedStyle: "heading1" },
});
const bookmarkTo = (blockId: string, len: number) => ({
  start: { blockId, offset: 0 },
  end: { blockId, offset: len },
});

describe("computeTocEdits / recalcToc", () => {
  it("rewrites a stale TOC page number to the target's real page", () => {
    const doc: Document = {
      section: SECTION,
      blocks: [tocEntry("toc1", "Chapter One", "_Toc1", "9"), heading("h1", "Chapter One")],
      bookmarks: { _Toc1: bookmarkTo("h1", "Chapter One".length) },
    };
    const edits = computeTocEditsWithLayout(doc);
    expect(edits).toHaveLength(1);
    expect(edits[0]!.blockId).toBe("toc1");
    expect(edits[0]!.text).toBe("1"); // heading lands on page 1

    const { doc: out, changed } = recalcToc(doc);
    expect(changed).toBe(1);
    const entry = out.blocks.find((b) => b.id === "toc1") as Paragraph;
    expect(textOfRuns(entry.runs)).toBe("Chapter One\t1");
  });

  it("is a no-op when the number is already current", () => {
    const doc: Document = {
      section: SECTION,
      blocks: [tocEntry("toc1", "Chapter One", "_Toc1", "1"), heading("h1", "Chapter One")],
      bookmarks: { _Toc1: bookmarkTo("h1", "Chapter One".length) },
    };
    expect(computeTocEditsWithLayout(doc)).toHaveLength(0);
    expect(recalcToc(doc).changed).toBe(0);
  });

  it("skips entries with no in-document anchor", () => {
    const plain: Paragraph = {
      kind: "paragraph",
      id: "toc1",
      revision: 0,
      runs: [{ text: "Chapter One\t9", style: CHAR }], // no link
      style: PARA,
    };
    const doc: Document = {
      section: SECTION,
      blocks: [plain, heading("h1", "Chapter One")],
      bookmarks: { _Toc1: bookmarkTo("h1", "Chapter One".length) },
    };
    expect(computeTocEditsWithLayout(doc)).toHaveLength(0);
  });

  it("skips entries without a trailing tab+number", () => {
    const noNum: Paragraph = {
      kind: "paragraph",
      id: "toc1",
      revision: 0,
      runs: [{ text: "Chapter One", style: { ...CHAR, link: "#_Toc1" } }],
      style: PARA,
    };
    const doc: Document = {
      section: SECTION,
      blocks: [noNum, heading("h1", "Chapter One")],
      bookmarks: { _Toc1: bookmarkTo("h1", "Chapter One".length) },
    };
    expect(computeTocEditsWithLayout(doc)).toHaveLength(0);
  });

  it("skips dangling entries whose bookmark target is gone", () => {
    const doc: Document = {
      section: SECTION,
      blocks: [tocEntry("toc1", "Chapter One", "_Toc1", "9")], // heading deleted
      bookmarks: {}, // anchor unresolvable
    };
    expect(computeTocEditsWithLayout(doc)).toHaveLength(0);
  });
});
