// Stateless headless render: raw .docx (empty TOC field + headings) → PDF, with the
// TOC entries built from the headings and footer page numbers computed by layout.
import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { renderPdfFromDocx } from "./serverExport";

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const docxOf = (body: string): Uint8Array =>
  zipSync({
    "[Content_Types].xml": strToU8(CONTENT_TYPES),
    "word/document.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`,
    ),
  });

const EMPTY_TOC =
  `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;

const heading = (name: string, text: string, brk = false): string =>
  `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/>${brk ? "<w:pageBreakBefore/>" : ""}</w:pPr>` +
  `<w:r><w:t>${text}</w:t></w:r></w:p>`;

/** ASCII "%PDF" — the file signature. */
const isPdf = (b: Uint8Array): boolean => b[0] === 0x25 && b[1] === 0x50 && b[2] === 0x44 && b[3] === 0x46;

describe("renderPdfFromDocx", () => {
  it("renders a .docx with an empty TOC field to a valid PDF", async () => {
    const out = await renderPdfFromDocx(docxOf(EMPTY_TOC + heading("_Toc1", "Chapter One") + heading("_Toc2", "Chapter Two", true)));
    expect(isPdf(out)).toBe(true);
    expect(out.length).toBeGreaterThan(1000);
  });

  it("renders a plain .docx (no TOC) to a valid PDF", async () => {
    const out = await renderPdfFromDocx(docxOf(`<w:p><w:r><w:t>Just a paragraph.</w:t></w:r></w:p>`));
    expect(isPdf(out)).toBe(true);
  });

  it("accepts TocOptions without throwing", async () => {
    const out = await renderPdfFromDocx(docxOf(EMPTY_TOC + heading("_Toc1", "Only")), { title: { text: "Contents" } });
    expect(isPdf(out)).toBe(true);
  });

  it("renders Chinese text via the bundled CJK fallback (no config)", async () => {
    // Before the bundled fallback, CJK runs hit a Latin clone and embedded as
    // .notdef/tofu. The backend passes no cjk config, so this also covers that
    // runExport defaults the fallback to the bundled face.
    const out = await renderPdfFromDocx(docxOf(`<w:p><w:r><w:t>中文测试，你好世界。</w:t></w:r></w:p>`));
    expect(isPdf(out)).toBe(true);
    expect(out.length).toBeGreaterThan(1000);
  });
});
