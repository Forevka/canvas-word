// Editor-authored TOCs (tocEntry paragraphs) export as a LIVE Word TOC field with
// per-entry PAGEREF fields and synthesized heading bookmarks — not frozen text.
import { strFromU8, unzipSync } from "fflate";
import { beforeAll, describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { installMeasureHost } from "../shared/measureHost";
import { runExport } from "../pipeline";

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
const para = (id: string, text: string, style: Partial<ParaStyle> = {}): Paragraph => ({
  kind: "paragraph",
  id,
  revision: 0,
  runs: [{ text, style: CHAR }],
  style: { ...PARA, ...style },
});

describe("live TOC field export", () => {
  beforeAll(async () => {
    await installMeasureHost();
  });

  it("wraps tocEntry paragraphs in a TOC field with PAGEREF + synthesized bookmark", async () => {
    const doc: Document = {
      section: SECTION,
      blocks: [
        para("t1", "Intro", { tocEntry: { targetId: "h1", level: 1 } }),
        para("h1", "Intro", { namedStyle: "heading1" }),
      ],
    };
    const { bytes } = await runExport(doc, "docx");
    const xml = strFromU8(unzipSync(bytes)["word/document.xml"]!);

    // A live TOC complex field (not w:fldSimple, not plain text).
    expect(xml).toContain('w:fldCharType="begin"');
    expect(xml).toMatch(/<w:instrText[^>]*>\s*TOC\b/);
    // Each entry's number is a live PAGEREF pointing at a (synthesized) heading bookmark.
    expect(xml).toMatch(/PAGEREF\s+_Toc9\d+/);
    expect(xml).toMatch(/<w:bookmarkStart[^>]*w:name="_Toc9\d+"/);
    // The cached page number (heading lands on page 1) is present.
    expect(xml).toContain("<w:t>1</w:t>");

    // Re-imports cleanly: the entry resolves to its heading bookmark.
    const b = runImport(bytes).doc;
    const entry = b.blocks.find((bl) => bl.kind === "paragraph" && bl.runs.some((r) => r.style.link?.startsWith("#_Toc"))) as Paragraph | undefined;
    expect(entry).toBeDefined();
  });
});
