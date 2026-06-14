// Backend recalc wiring: recalcDocxBytes imports + lays out + patches the original
// .docx in place. A TOC entry (nested HYPERLINK→PAGEREF) with a stale cached page
// number is rewritten to the heading's real layout page (1 here), and the output is
// a valid .docx with the live PAGEREF field intact.

import { strToU8, zipSync } from "fflate";
import { runImport } from "@forevka/wordcanvas/import";
import { textOfRuns, type Paragraph } from "@cw/shared";
import { describe, expect, it } from "vitest";
import { recalcDocxBytes } from "./serverRecalc";

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

function staleTocDocx(cached: string): Uint8Array {
  const body =
    `<w:p>` +
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> HYPERLINK \\l "_Toc1" </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t>Chapter One</w:t></w:r><w:r><w:tab/></w:r>` +
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t>${cached}</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
    `</w:p>` +
    `<w:p><w:bookmarkStart w:id="1" w:name="_Toc1"/><w:r><w:t>Chapter One</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>`;
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return zipSync({ "[Content_Types].xml": strToU8(CONTENT_TYPES), "word/document.xml": strToU8(document) });
}

describe("recalcDocxBytes (patch-in-place)", () => {
  it("rewrites a stale cached TOC page number and returns a valid docx", async () => {
    const out = await recalcDocxBytes(staleTocDocx("9"));
    expect(out.changed).toBe(1);
    expect(out.skipped).toBe(0);
    expect(out.bytes[0]).toBe(0x50); // PK
    expect(out.bytes[1]).toBe(0x4b);
    const entry = runImport(out.bytes).doc.blocks[0] as Paragraph;
    expect(textOfRuns(entry.runs)).toBe("Chapter One\t1");
  });

  it("is a no-op (changed 0) when the cached number is already correct", async () => {
    const out = await recalcDocxBytes(staleTocDocx("1"));
    expect(out.changed).toBe(0);
  });

  it("skips a non-arabic cached page number", async () => {
    const out = await recalcDocxBytes(staleTocDocx("iv"));
    expect(out.changed).toBe(0);
    expect(out.skipped).toBe(1);
  });
});
