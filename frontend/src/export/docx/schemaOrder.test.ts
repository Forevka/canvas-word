// Word validates every property element (w:pPr, w:rPr, w:tcPr, w:sectPr, w:settings)
// against the strict OOXML xsd:sequence and rejects the whole file on the FIRST
// out-of-order child ("Word experienced an error trying to open the file"). Lenient
// consumers (Google Docs, LibreOffice) and our own order-independent importer don't,
// which let the reordering bug in issue #180 ship unnoticed. These tests pin the
// emitted child order to the schema sequence so a future field added in the wrong
// spot fails here instead of in Word.
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps, TableBlock } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { writeDocx } from "./writeDocx";
import { sampleDoc } from "../../model/sampleDoc";
import { orderChildren, PPR_ORDER, RPR_ORDER, SECTPR_ORDER, TCPR_ORDER } from "./schemaOrder";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#202124" };
const PARA: ParaStyle = { align: "left", lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let n = 0;
const para = (text: string, style: Partial<ParaStyle> = {}, char: Partial<CharStyle> = {}): Paragraph => ({
  kind: "paragraph", id: `p${n++}`, revision: 0, runs: [{ text, style: { ...CHAR, ...char } }], style: { ...PARA, ...style },
});
const exportedXml = (doc: Document): string => strFromU8(unzipSync(writeDocx(doc).bytes)["word/document.xml"]!);
const settingsXml = (doc: Document): string => strFromU8(unzipSync(writeDocx(doc).bytes)["word/settings.xml"]!);

/** Assert the child element tags found in `container` XML appear in `order`'s
 *  sequence — i.e. the schema position of each child is non-decreasing. */
function expectSchemaOrder(container: string, order: readonly string[]): void {
  const rank = new Map(order.map((t, i) => [t, i]));
  const tags = [...container.matchAll(/<([\w]+:[\w]+)[\s/>]/g)].map((m) => m[1]!).filter((t) => rank.has(t));
  const ranks = tags.map((t) => rank.get(t)!);
  for (let i = 1; i < ranks.length; i++) {
    expect(ranks[i]! >= ranks[i - 1]!, `${tags[i - 1]} (schema #${ranks[i - 1]}) must not precede ${tags[i]} (schema #${ranks[i]})`).toBe(true);
  }
}

/** First w:pPr / w:rPr / w:tcPr / w:sectPr block in a document XML string. */
const firstBlock = (xml: string, tag: string): string => {
  const open = xml.indexOf(`<${tag}>`);
  const openAttr = open >= 0 ? open : xml.indexOf(`<${tag} `);
  const end = xml.indexOf(`</${tag}>`, openAttr);
  return openAttr >= 0 && end >= 0 ? xml.slice(openAttr, end) : "";
};

describe("orderChildren", () => {
  it("stably reorders to the given schema order", () => {
    // tabs(11) < bidi(19) < spacing(22) < jc(27): pushed out of order, sorted in.
    const sorted = orderChildren(["<w:jc/>", "<w:spacing/>", "<w:tabs/>", "<w:bidi/>"], PPR_ORDER);
    expect(sorted).toEqual(["<w:tabs/>", "<w:bidi/>", "<w:spacing/>", "<w:jc/>"]);
  });
  it("keeps equal-rank elements in insertion order (stable)", () => {
    const a = '<w:headerReference w:type="default"/>';
    const b = '<w:headerReference w:type="first"/>';
    expect(orderChildren([b, a], SECTPR_ORDER)).toEqual([b, a]);
  });
  it("sorts unrecognized tags to the end, preserving order", () => {
    expect(orderChildren(["<w:zzz/>", "<w:jc/>", "<w:aaa/>"], PPR_ORDER)).toEqual(["<w:jc/>", "<w:zzz/>", "<w:aaa/>"]);
  });
});

describe("exported OOXML is in strict schema order (issue #180)", () => {
  it("w:pPr — tabs/shading/borders precede spacing/jc", () => {
    const xml = exportedXml({
      section: SECTION,
      blocks: [para("x", {
        tabStops: [{ posPx: 96, align: "left" }],
        shading: "#ffff00",
        borders: { bottom: { style: "single", widthPx: 1, color: "#000000" } },
        keepWithNext: true,
        align: "center",
      })],
    });
    expectSchemaOrder(firstBlock(xml, "w:pPr"), PPR_ORDER);
    // The specific violation Word rejected: w:tabs after w:spacing.
    const pPr = firstBlock(xml, "w:pPr");
    expect(pPr.indexOf("<w:tabs")).toBeLessThan(pPr.indexOf("<w:spacing"));
    expect(pPr.indexOf("<w:shd")).toBeLessThan(pPr.indexOf("<w:tabs"));
  });

  it("w:rPr — w:spacing(letter)/w:sz precede w:u; effects land in sequence", () => {
    const xml = exportedXml({
      section: SECTION,
      blocks: [para("x", {}, { letterSpacingPx: 1, underline: true, doubleStrikethrough: true, kerningMinPx: 2, emphasisMark: "dot" })],
    });
    // First rPr is the run's.
    expectSchemaOrder(firstBlock(xml, "w:rPr"), RPR_ORDER);
  });

  it("w:tcPr — w:tcBorders precedes w:shd", () => {
    const table: TableBlock = {
      kind: "table", id: "t1", revision: 0, widthMode: "fixed",
      rows: [{ cells: [{
        id: "cell1",
        blocks: [para("c")],
        shading: "#ff0000",
        borders: { top: { style: "single", widthPx: 1, color: "#000000" } },
      }] }],
    };
    const xml = exportedXml({ section: SECTION, blocks: [table] });
    const tcPr = firstBlock(xml, "w:tcPr");
    expectSchemaOrder(tcPr, TCPR_ORDER);
    expect(tcPr.indexOf("<w:tcBorders")).toBeLessThan(tcPr.indexOf("<w:shd"));
  });

  it("w:sectPr — w:type precedes w:pgSz/w:lnNumType", () => {
    const xml = exportedXml({
      section: { ...SECTION, lineNumbering: { countBy: 1 }, breakType: "evenPage" },
      blocks: [para("x")],
    });
    const sectPr = firstBlock(xml, "w:sectPr");
    expectSchemaOrder(sectPr, SECTPR_ORDER);
    expect(sectPr.indexOf("<w:type")).toBeLessThan(sectPr.indexOf("<w:pgSz"));
    expect(sectPr.indexOf("<w:pgSz")).toBeLessThan(sectPr.indexOf("<w:lnNumType"));
  });

  it("settings.xml — w:defaultTabStop precedes w:evenAndOddHeadersAndFooters", () => {
    const xml = settingsXml({
      section: { ...SECTION, headerEven: [para("e")], header: [para("h")] },
      blocks: [para("x")],
      defaultTabStopPx: 96,
    });
    expect(xml.indexOf("<w:defaultTabStop")).toBeLessThan(xml.indexOf("<w:evenAndOddHeadersAndFooters"));
  });

  it("every property block in the real showcase document is schema-ordered", () => {
    // The offline example (doc-editor.forevka.dev/examples/offline) exports exactly
    // this document — the file the issue #180 reporter could not open in Word. Scan
    // EVERY w:pPr / w:rPr / w:tcPr / w:sectPr in the output, not just crafted cases.
    const xml = exportedXml(sampleDoc());
    // Collapse nested blocks to a self-closing stub so a container is checked on its
    // OWN direct children only. A pPr embeds the paragraph-mark w:rPr and a section-
    // break w:sectPr, whose descendants (w:spacing, w:shd, …) also exist at pPr level
    // and would otherwise be misread as out-of-order pPr children.
    const collapse = (s: string): string =>
      s.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, "<w:rPr/>").replace(/<w:sectPr[\s\S]*?<\/w:sectPr>/g, "<w:sectPr/>");
    const containers: [RegExp, readonly string[]][] = [
      [/<w:pPr>[\s\S]*?<\/w:pPr>/g, PPR_ORDER],
      [/<w:rPr>[\s\S]*?<\/w:rPr>/g, RPR_ORDER],
      [/<w:tcPr>[\s\S]*?<\/w:tcPr>/g, TCPR_ORDER],
      [/<w:sectPr[\s\S]*?<\/w:sectPr>/g, SECTPR_ORDER],
    ];
    let scanned = 0;
    // pPr first: collapse its nested rPr/sectPr, then check. The other containers are
    // leaf-level (no same-name or overlapping-tag nesting) so are checked as-is.
    for (const m of xml.matchAll(containers[0]![0])) { expectSchemaOrder(collapse(m[0]), PPR_ORDER); scanned++; }
    for (const [re, order] of containers.slice(1)) {
      for (const m of xml.matchAll(re)) { expectSchemaOrder(m[0], order); scanned++; }
    }
    expect(scanned).toBeGreaterThan(20); // sanity: the showcase is feature-rich
  });

  it("the full sweep still re-imports equal (order change is lossless)", () => {
    const doc: Document = {
      section: SECTION,
      blocks: [para("x", { tabStops: [{ posPx: 96, align: "left" }], shading: "#ffff00", align: "center" })],
    };
    const re = runImport(writeDocx(doc).bytes).doc;
    const p = re.blocks[0] as Paragraph;
    expect(p.style.align).toBe("center");
    expect(p.style.tabStops?.[0]?.posPx).toBe(96);
    expect(p.style.shading?.toLowerCase()).toBe("#ffff00");
  });
});
