// Headless TOC field-result generation: an empty TOC field + headings → generated
// entries with correct PAGEREF page numbers, hyperlinks, and a dot leader, spliced
// into the original XML (drift-free).
import { strFromU8, unzipSync } from "fflate";
import { beforeAll, describe, expect, it } from "vitest";
import { simpleDocx } from "../import/docx/fixture";
import { installMeasureHost } from "../export/shared/measureHost";
import { generateTocInDocx } from "./generateTocDocx";

const heading = (name: string, text: string, breakBefore = false): string => {
  const id = name.replace(/\D/g, "") || "1"; // unique per _TocN
  return (
    `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/>${breakBefore ? "<w:pageBreakBefore/>" : ""}</w:pPr>` +
    `<w:bookmarkStart w:id="${id}" w:name="${name}"/><w:r><w:t>${text}</w:t></w:r><w:bookmarkEnd w:id="${id}"/></w:p>`
  );
};

/** Empty TOC field + two bookmarked headings, the 2nd forced onto page 2. */
function emptyTocDocx(instr = ' TOC \\o "1-3" \\h \\z \\u '): Uint8Array {
  const tocField =
    `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve">${instr}</w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t>Press F9</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
  return simpleDocx(tocField + heading("_Toc1", "Chapter One") + heading("_Toc2", "Chapter Two", true));
}

/** The cached page number inside a PAGEREF field (the first numeric w:t after it). */
const cachedNum = (xml: string, anchor: string): string | undefined => {
  const i = xml.indexOf(`PAGEREF ${anchor}`);
  return i < 0 ? undefined : xml.slice(i).match(/<w:t[^>]*>(\d+)<\/w:t>/)?.[1];
};

describe("generateTocInDocx", () => {
  beforeAll(async () => {
    await installMeasureHost();
  });

  it("generates entries with correct page numbers, hyperlinks, and a leader", () => {
    const src = emptyTocDocx();
    const out = generateTocInDocx(src);
    expect(out.generated).toBe(2);
    expect(out.headings).toBe(2);
    expect(out.bookmarksSynthesized).toBe(0); // headings already had _Toc bookmarks

    const xml = strFromU8(unzipSync(out.bytes)["word/document.xml"]!);
    expect(xml).toMatch(/<w:instrText[^>]*>\s*TOC\b/); // TOC field preserved
    expect(xml).toContain("PAGEREF _Toc1");
    expect(xml).toContain("PAGEREF _Toc2");
    expect(xml).toContain('w:anchor="_Toc1"'); // \h hyperlink
    expect(xml).toContain('w:leader="dot"'); // dot leader tab stop

    // Cached PAGEREF numbers are the real pages (Chapter Two forced to page 2).
    expect(cachedNum(xml, "_Toc1")).toBe("1");
    expect(cachedNum(xml, "_Toc2")).toBe("2");
  });

  it("is drift-free / idempotent: re-running yields the same pages", () => {
    const once = generateTocInDocx(emptyTocDocx()).bytes;
    const twice = generateTocInDocx(once);
    expect(twice.generated).toBe(2);
    const xml = strFromU8(unzipSync(twice.bytes)["word/document.xml"]!);
    expect(cachedNum(xml, "_Toc1")).toBe("1");
    expect(cachedNum(xml, "_Toc2")).toBe("2");
  });

  it("honors \\n (omit page numbers) — no PAGEREF for the hidden level", () => {
    const out = generateTocInDocx(emptyTocDocx(' TOC \\o "1-3" \\h \\n "1-3" '));
    const xml = strFromU8(unzipSync(out.bytes)["word/document.xml"]!);
    expect(xml).not.toContain("PAGEREF");
    expect(out.generated).toBe(2);
  });

  it("returns the doc unchanged when there is no TOC field", () => {
    const src = simpleDocx(`<w:p><w:r><w:t>no toc here</w:t></w:r></w:p>`);
    const out = generateTocInDocx(src);
    expect(out.generated).toBe(0);
    expect(out.headings).toBe(0);
  });

  it("synthesizes + injects a bookmark for a heading that lacks one", () => {
    const tocField =
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const noBmHeading = `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>Lonely Heading</w:t></w:r></w:p>`;
    const out = generateTocInDocx(simpleDocx(tocField + noBmHeading));
    expect(out.generated).toBe(1);
    expect(out.bookmarksSynthesized).toBe(1);
    const xml = strFromU8(unzipSync(out.bytes)["word/document.xml"]!);
    expect(xml).toMatch(/<w:bookmarkStart[^>]*w:name="_CwToc0"/);
    expect(xml).toContain("PAGEREF _CwToc0");
  });
});
