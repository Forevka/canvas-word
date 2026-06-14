// patch-in-place TOC recalculation: rewrite cached PAGEREF result digits in the
// original document.xml without disturbing anything else.
import { strFromU8, unzipSync } from "fflate";
import { beforeAll, describe, expect, it } from "vitest";
import { textOfRuns, type Paragraph } from "@cw/shared";
import { runImport } from "../import/docx/pipeline";
import { simpleDocx } from "../import/docx/fixture";
import { installMeasureHost } from "../export/shared/measureHost";
import { patchTocFromLayout, patchTocPageNumbers } from "./patchTocDocx";

/** A TOC entry (nested HYPERLINK + PAGEREF) whose cached page number is `cached`,
 *  plus a heading carrying the _Toc1 bookmark. */
function tocDocx(cached: string): Uint8Array {
  return simpleDocx(
    `<w:p>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> HYPERLINK \\l "_Toc1" </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r><w:t>Chapter One</w:t></w:r><w:r><w:tab/></w:r>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      cached +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
      `</w:p>` +
      `<w:p><w:bookmarkStart w:id="1" w:name="_Toc1"/><w:r><w:t>Chapter One</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>`,
  );
}

const docXml = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)["word/document.xml"]!);
const entryText = (bytes: Uint8Array): string =>
  textOfRuns((runImport(bytes).doc.blocks[0] as Paragraph).runs);

describe("patchTocPageNumbers", () => {
  it("rewrites a stale cached PAGEREF number to the mapped page", () => {
    const out = patchTocPageNumbers(tocDocx(`<w:r><w:t>9</w:t></w:r>`), { _Toc1: "5" });
    expect(out.changed).toBe(1);
    expect(out.skipped).toBe(0);
    expect(entryText(out.bytes)).toBe("Chapter One\t5");
  });

  it("leaves a non-arabic (roman) cached value untouched and counts it skipped", () => {
    const out = patchTocPageNumbers(tocDocx(`<w:r><w:t>iv</w:t></w:r>`), { _Toc1: "5" });
    expect(out.changed).toBe(0);
    expect(out.skipped).toBe(1);
    expect(docXml(out.bytes)).toContain("<w:t>iv</w:t>");
  });

  it("handles a page number split across runs (sets first, blanks the rest)", () => {
    const out = patchTocPageNumbers(tocDocx(`<w:r><w:t>1</w:t></w:r><w:r><w:t>2</w:t></w:r>`), { _Toc1: "7" });
    expect(out.changed).toBe(1);
    expect(entryText(out.bytes)).toBe("Chapter One\t7");
  });

  it("is a no-op when the number is already current", () => {
    const out = patchTocPageNumbers(tocDocx(`<w:r><w:t>5</w:t></w:r>`), { _Toc1: "5" });
    expect(out.changed).toBe(0);
  });

  it("preserves other parts byte-for-byte", () => {
    const src = tocDocx(`<w:r><w:t>9</w:t></w:r>`);
    const out = patchTocPageNumbers(src, { _Toc1: "5" });
    const before = unzipSync(src);
    const after = unzipSync(out.bytes);
    expect([...after["[Content_Types].xml"]!]).toEqual([...before["[Content_Types].xml"]!]);
  });
});

describe("patchTocFromLayout", () => {
  beforeAll(async () => {
    await installMeasureHost();
  });

  it("computes the page from layout and patches it (heading on page 1)", () => {
    const src = tocDocx(`<w:r><w:t>9</w:t></w:r>`);
    const out = patchTocFromLayout(src, runImport(src).doc);
    expect(out.changed).toBe(1);
    expect(entryText(out.bytes)).toBe("Chapter One\t1");
  });
});
