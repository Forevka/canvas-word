// Headless TOC build: an empty `TOC` field (a C# pipeline emits the field + headings
// but can't compute layout) imports with no entries; import records the field's
// block (doc.tocAnchorBlockId) so generateTocIntoDoc builds the entries in place —
// honoring the field's \o level range and inheriting the doc's own TOC styling.
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { generateTocIntoDoc, textOfRuns, type Paragraph } from "@cw/shared";
import { runImport } from "./pipeline";
import { CONTENT_TYPES_XML, documentXml, drawingXml, makeDocx, PNG_1PX, REL_TYPES, relsXml, simpleDocx } from "./fixture";

const emptyToc = (range = "1-3", runPr = ""): string =>
  `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>` +
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> TOC \\o "${range}" \\h \\z </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
  `<w:r>${runPr}<w:t>Please press F9 to update the Table of Contents.</w:t></w:r>` +
  `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;

const heading = (name: string, text: string, level = 1, brk = false): string =>
  `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/><w:outlineLvl w:val="${level - 1}"/>${brk ? "<w:pageBreakBefore/>" : ""}</w:pPr>` +
  `<w:bookmarkStart w:id="${name.replace(/\D/g, "")}" w:name="${name}"/><w:r><w:t>${text}</w:t></w:r>` +
  `<w:bookmarkEnd w:id="${name.replace(/\D/g, "")}"/></w:p>`;

const tocEntries = (blocks: { kind: string }[]): Paragraph[] =>
  blocks.filter((b): b is Paragraph => b.kind === "paragraph" && !!(b as Paragraph).style.tocEntry);

describe("import → generateTocIntoDoc", () => {
  it("anchors an empty TOC field and builds entries at its location (no auto-title)", () => {
    const doc = runImport(simpleDocx(emptyToc() + heading("_Toc1", "Chapter One") + heading("_Toc2", "Chapter Two", 1, true))).doc;
    expect(doc.tocInstruction).toBeDefined();
    expect(doc.tocAnchorBlockId).toBe(doc.blocks[0]!.id); // the empty TOC paragraph
    expect(tocEntries(doc.blocks).length).toBe(0); // nothing built yet
    const h1 = doc.blocks[1]!.id;
    const h2 = doc.blocks[2]!.id;

    const r = generateTocIntoDoc(doc, {});
    expect(r.generated).toBe(2);
    const entries = tocEntries(r.doc.blocks);
    expect(entries.map((e) => e.style.tocEntry!.targetId)).toEqual([h1, h2]);
    expect(r.doc.blocks.some((b) => b.id === doc.tocAnchorBlockId)).toBe(false); // empty para replaced
    // No generated title — the doc owns its own "TABLE OF CONTENTS" heading.
    expect(r.doc.blocks.some((b) => b.kind === "paragraph" && (b as Paragraph).style.namedStyle === "tocTitle")).toBe(false);
  });

  it("honors the field's \\o range (1-5 lists five levels, not the default three)", () => {
    const body = emptyToc("1-5") + [1, 2, 3, 4, 5].map((l) => heading(`_Toc${l}`, `H${l}`, l)).join("");
    const r = generateTocIntoDoc(runImport(simpleDocx(body)).doc, {});
    expect(r.generated).toBe(5);
    expect(tocEntries(r.doc.blocks).map((e) => e.style.tocEntry!.level)).toEqual([1, 2, 3, 4, 5]);
  });

  it("inherits the TOC field paragraph's font into the built entries", () => {
    // The placeholder run carries Arial 20pt (sz=40 half-points) — entries inherit it
    // instead of the built-in Georgia default.
    const runPr = `<w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="40"/></w:rPr>`;
    const doc = runImport(simpleDocx(emptyToc("1-3", runPr) + heading("_Toc1", "Only"))).doc;
    const entry = tocEntries(generateTocIntoDoc(doc, {}).doc.blocks)[0]!;
    expect(entry.runs[0]!.style.fontFamily).toMatch(/Arial/);
    expect(entry.runs[0]!.style.fontSizePx).toBeCloseTo(26.67, 1); // 20pt
  });

  it("lets explicit TocOptions override the inherited look (title + size)", () => {
    const doc = runImport(simpleDocx(emptyToc() + heading("_Toc1", "Only"))).doc;
    const r = generateTocIntoDoc(doc, { title: { text: "Contents" }, levels: { 1: { char: { fontSizePx: 99 } } } });
    expect((r.doc.blocks[0] as Paragraph).runs[0]!.text).toBe("Contents");
    expect(tocEntries(r.doc.blocks)[0]!.runs[0]!.style.fontSizePx).toBe(99);
  });

  // Regression: a real C#-rendered report (cover-page <w:sdt>s, a table, drawings)
  // where ordinal counting placed the TOC ~11 blocks too early, before the heading.
  it("anchors at the TOC field even with sdt/table/drawing noise (real report)", () => {
    const bytes = readFileSync(
      "RenderedReportsDocx_Report-Version-5-CM099026-2876-6323 Sunset Ave, Panama City Beach, Bay, Florida, 32408 EMPTY TOC.docx",
    );
    const doc = runImport(new Uint8Array(bytes)).doc;
    const idx = (pred: (p: Paragraph) => boolean): number =>
      doc.blocks.findIndex((b) => b.kind === "paragraph" && pred(b as Paragraph));
    const headingIdx = idx((p) => textOfRuns(p.runs).trim() === "TABLE OF CONTENTS");
    const anchorIdx = doc.blocks.findIndex((b) => b.id === doc.tocAnchorBlockId);
    expect(headingIdx).toBeGreaterThanOrEqual(0);
    // The anchor must be the TOC field paragraph, just after its heading (+ a spacer).
    expect(anchorIdx).toBeGreaterThan(headingIdx);
    expect(anchorIdx - headingIdx).toBeLessThanOrEqual(2);

    // After building, the first entry lands right after the heading, not pages earlier.
    const r = generateTocIntoDoc(doc, {});
    const builtHeadingIdx = r.doc.blocks.findIndex((b) => b.kind === "paragraph" && textOfRuns((b as Paragraph).runs).trim() === "TABLE OF CONTENTS");
    const firstEntryIdx = r.doc.blocks.findIndex((b) => b.kind === "paragraph" && (b as Paragraph).style.tocEntry);
    expect(firstEntryIdx).toBeGreaterThan(builtHeadingIdx);
    expect(firstEntryIdx - builtHeadingIdx).toBeLessThanOrEqual(2);
  });

  // Regression: a report whose TOC field ALREADY has entries must be preserved
  // (its own styling + layout-recomputed page numbers), NOT rebuilt with defaults.
  it("preserves an already-populated TOC (real report) instead of rebuilding", () => {
    const bytes = readFileSync(
      "RenderedReportsDocx_Report-Version-5-CM099026-2876-6323 Sunset Ave, Panama City Beach, Bay, Florida, 32408.docx",
    );
    const doc = runImport(new Uint8Array(bytes)).doc;
    const before = tocEntries(doc.blocks);
    expect(before.length).toBeGreaterThan(0); // imported entries were marked as tocEntry
    expect(doc.tocAnchorBlockId).toBeUndefined(); // populated → no empty-field anchor

    const r = generateTocIntoDoc(doc, {});
    expect(r.generated).toBe(0); // preserved, not regenerated
    // Same entry blocks (identical ids ⇒ untouched; a rebuild would mint fresh ids).
    expect(tocEntries(r.doc.blocks).map((e) => e.id)).toEqual(before.map((e) => e.id));
  });

  // A misapplied heading style on an image/table must NOT pull that object into the
  // TOC (cf. Syncfusion, which embeds it). Only non-empty text paragraphs are entries.
  it("never lists an image or table as a TOC entry", () => {
    const imageHeading =
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/></w:pPr>` +
      `<w:r>${drawingXml("rId1", 914400, 914400)}</w:r></w:p>`; // heading style, image only, no text
    const table = `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const body = emptyToc() + heading("_Toc1", "Real Heading") + imageHeading + table;
    const docx = makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(body),
      "word/_rels/document.xml.rels": relsXml([{ id: "rId1", type: REL_TYPES.image, target: "media/image1.png" }]),
      "word/media/image1.png": PNG_1PX,
    });
    const r = generateTocIntoDoc(runImport(docx).doc, {});
    // Only the real text heading — the image-only heading (empty text) and the table
    // are excluded.
    expect(tocEntries(r.doc.blocks).map((e) => textOfRuns(e.runs))).toEqual(["Real Heading"]);
  });
});
