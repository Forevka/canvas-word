// Backend TOC generation wiring: generateTocBytes imports + lays out + splices the
// generated field result into the original .docx in place.
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { generateTocBytes } from "./serverGenerateToc";

/** The cached page number inside a PAGEREF field (the first numeric w:t after it). */
const cachedNum = (xml: string, anchor: string): string | undefined => {
  const i = xml.indexOf(`PAGEREF ${anchor}`);
  return i < 0 ? undefined : xml.slice(i).match(/<w:t[^>]*>(\d+)<\/w:t>/)?.[1];
};

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const heading = (name: string, text: string, brk = false): string =>
  `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/>${brk ? "<w:pageBreakBefore/>" : ""}</w:pPr>` +
  `<w:bookmarkStart w:id="${name.replace(/\D/g, "")}" w:name="${name}"/><w:r><w:t>${text}</w:t></w:r>` +
  `<w:bookmarkEnd w:id="${name.replace(/\D/g, "")}"/></w:p>`;

function emptyTocDocx(): Uint8Array {
  const toc =
    `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
  const body = toc + heading("_Toc1", "Chapter One") + heading("_Toc2", "Chapter Two", true);
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return zipSync({ "[Content_Types].xml": strToU8(CONTENT_TYPES), "word/document.xml": strToU8(document) });
}

describe("generateTocBytes", () => {
  it("generates the field result with correct pages and a valid docx", async () => {
    const r = await generateTocBytes(emptyTocDocx());
    expect(r.generated).toBe(2);
    expect(r.headings).toBe(2);
    expect(r.bytes[0]).toBe(0x50); // PK
    const xml = strFromU8(unzipSync(r.bytes)["word/document.xml"]!);
    expect(cachedNum(xml, "_Toc1")).toBe("1");
    expect(cachedNum(xml, "_Toc2")).toBe("2"); // Chapter Two forced to page 2
  });

  it("applies caller TocOptions (custom title text + leader off)", async () => {
    const r = await generateTocBytes(emptyTocDocx(), { leader: "none", title: { text: "Contents" } });
    const xml = strFromU8(unzipSync(r.bytes)["word/document.xml"]!);
    expect(xml).not.toContain('w:leader="dot"');
  });
});
