// Paragraph contextual spacing (w:contextualSpacing) must survive a full .docx
// round-trip — both as direct paragraph formatting and as a paragraph-style delta
// in styles.xml — and be parsed from a hand-written pPr. The importer is the
// oracle: a model → writeDocx → re-import that drops contextualSpacing surfaces here.
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import type { CharStyle, Document, Paragraph, SectionProps, Stylesheet } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { simpleDocx } from "../../import/docx/fixture";
import { writeDocx } from "./writeDocx";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 14, bold: false, italic: false, underline: false, strikethrough: false, color: "#202124" };
const PARA = { align: "left" as const, lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 12, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let n = 0;
const para = (text: string, contextualSpacing?: boolean): Paragraph => ({
  kind: "paragraph", id: `p${n++}`, revision: 0, runs: [{ text, style: { ...CHAR } }],
  style: { ...PARA, ...(contextualSpacing ? { contextualSpacing: true } : {}) },
});

const roundTrip = (doc: Document): Document => runImport(writeDocx(doc).bytes).doc;
const exportedXml = (doc: Document): string => strFromU8(unzipSync(writeDocx(doc).bytes)["word/document.xml"]!);
const paraOf = (d: Document, i: number): Paragraph => d.blocks[i] as Paragraph;

describe("paragraph contextual spacing — w:contextualSpacing", () => {
  it("parses w:contextualSpacing from a hand-written pPr", () => {
    const body = `<w:p><w:pPr><w:contextualSpacing/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>y</w:t></w:r></w:p>`;
    const d = runImport(simpleDocx(body)).doc;
    expect(paraOf(d, 0).style.contextualSpacing).toBe(true);
    expect(paraOf(d, 1).style.contextualSpacing).toBeUndefined();
  });

  it("emits w:contextualSpacing only for flagged paragraphs", () => {
    const xml = exportedXml({ section: SECTION, blocks: [para("a", true), para("b")] });
    expect(xml).toContain("<w:contextualSpacing/>");
    // Exactly one paragraph carries it.
    expect(xml.match(/<w:contextualSpacing\/>/g)).toHaveLength(1);
  });

  it("preserves the flag through a full export → re-import round-trip", () => {
    const out = roundTrip({ section: SECTION, blocks: [para("a", true), para("b", true), para("c")] });
    expect(paraOf(out, 0).style.contextualSpacing).toBe(true);
    expect(paraOf(out, 1).style.contextualSpacing).toBe(true);
    expect(paraOf(out, 2).style.contextualSpacing).toBeUndefined();
  });

  it("round-trips contextualSpacing carried on a paragraph STYLE delta", () => {
    const stylesheet: Stylesheet = {
      defaultStyleId: "Normal",
      styles: [
        { id: "Normal", name: "Normal", char: { ...CHAR }, para: { ...PARA } },
        { id: "Verse", name: "Verse", basedOn: "Normal", char: {}, para: { contextualSpacing: true } },
      ],
    };
    const styled: Paragraph = { ...para("v"), style: { ...PARA, namedStyle: "Verse" } };
    const out = roundTrip({ section: SECTION, stylesheet, blocks: [styled] });
    const verse = out.stylesheet!.styles.find((s) => s.id === "Verse")!;
    expect(verse.para.contextualSpacing).toBe(true);
  });
});
