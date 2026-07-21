// Issue #63 — miscellaneous OOXML features must survive a full .docx round-trip:
//   • w:sym         — symbol-font glyphs (font + hex code point)
//   • a:srcRect      — DrawingML image crop insets
//   • in-cell images — anchored (floating) images inside a table cell
//   • settings.xml   — w:defaultTabStop + w:compat/w:compatSetting
// Each is exercised both from a hand-written .docx (the importer is the oracle) and
// through a model → writeDocx → re-import cycle.
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import type { CharStyle, Document, ImageBlock, Paragraph, Run, SectionProps, TableBlock, TableCell } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { CONTENT_TYPES_XML, documentXml, makeDocx, PNG_1PX, REL_TYPES, relsXml, simpleDocx } from "../../import/docx/fixture";
import { writeDocx } from "./writeDocx";

const CHAR: CharStyle = { fontFamily: "Times New Roman, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" };
const PARA = { align: "left" as const, lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let n = 0;
const run = (text: string, patch: Partial<CharStyle> = {}): Run => ({ text, style: { ...CHAR, ...patch } });
const para = (runs: Run[]): Paragraph => ({ kind: "paragraph", id: `p${n++}`, revision: 0, runs, style: { ...PARA } });

const roundTrip = (doc: Document, images: Record<string, Uint8Array> = {}): Document => runImport(writeDocx(doc, images).bytes).doc;
const exportedDocXml = (doc: Document, images: Record<string, Uint8Array> = {}): string => strFromU8(unzipSync(writeDocx(doc, images).bytes)["word/document.xml"]!);
const exportedSettingsXml = (doc: Document): string => strFromU8(unzipSync(writeDocx(doc).bytes)["word/settings.xml"]!);
const firstPara = (d: Document): Paragraph => d.blocks.find((b): b is Paragraph => b.kind === "paragraph")!;
const firstImage = (d: Document): ImageBlock => d.blocks.find((b): b is ImageBlock => b.kind === "image")!;
const firstTable = (d: Document): TableBlock => d.blocks.find((b): b is TableBlock => b.kind === "table")!;

// ── w:sym (symbol-font glyphs) ───────────────────────────────────────────────
describe("symbols — w:sym", () => {
  it("parses w:sym into a run carrying the font + code point (text = the decoded glyph)", () => {
    const r = runImport(simpleDocx(`<w:p><w:r><w:sym w:font="Wingdings" w:char="F0E0"/></w:r></w:p>`)).doc;
    const run0 = firstPara(r).runs.find((x) => x.style.symbol)!;
    expect(run0.style.symbol).toEqual({ font: "Wingdings", char: "F0E0" });
    expect(run0.text).toBe(String.fromCodePoint(0xf0e0));
    // the font is set so the glyph paints in the symbol face, not body text
    expect(run0.style.fontFamily).toContain("Wingdings");
  });

  it("emits w:sym (not w:t) on export", () => {
    const doc: Document = { section: SECTION, blocks: [para([run("x", { fontFamily: "Wingdings, sans-serif", symbol: { font: "Wingdings", char: "F04A" } })])] };
    expect(exportedDocXml(doc)).toContain('<w:sym w:font="Wingdings" w:char="F04A"/>');
  });

  it("falls back to the replacement glyph for a code point above the Unicode max (no throw)", () => {
    // 0x110000 > 0x10FFFF would make String.fromCodePoint throw — import must survive.
    const r = runImport(simpleDocx(`<w:p><w:r><w:sym w:font="Wingdings" w:char="110000"/></w:r></w:p>`)).doc;
    const run0 = firstPara(r).runs.find((x) => x.style.symbol)!;
    expect(run0.text).toBe("�");
    expect(run0.style.symbol).toEqual({ font: "Wingdings", char: "110000" });
  });

  it("survives a model → export → re-import round-trip", () => {
    const glyph = String.fromCodePoint(0xf04a);
    const doc: Document = { section: SECTION, blocks: [para([run("before "), run(glyph, { fontFamily: "Wingdings, sans-serif", symbol: { font: "Wingdings", char: "F04A" } }), run(" after")])] };
    const out = firstPara(roundTrip(doc));
    const sym = out.runs.find((x) => x.style.symbol)!;
    expect(sym.style.symbol).toEqual({ font: "Wingdings", char: "F04A" });
    expect(sym.text).toBe(glyph);
    // the symbol run must NOT merge with the surrounding plain text
    expect(out.runs.filter((x) => x.style.symbol)).toHaveLength(1);
  });
});

// ── a:srcRect (image crop) ───────────────────────────────────────────────────
/** document.xml + a PNG media part, body supplied by the caller. */
const docxWithImage = (body: string): Uint8Array =>
  makeDocx({
    "[Content_Types].xml": CONTENT_TYPES_XML,
    "word/document.xml": documentXml(body),
    "word/_rels/document.xml.rels": relsXml([{ id: "rId10", type: REL_TYPES.image, target: "media/img.png" }]),
    "word/media/img.png": PNG_1PX,
  });

const croppedDrawing = (l: number, t: number, r: number, b: number): string =>
  `<w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>` +
  `<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId10"/>` +
  `<a:srcRect l="${l}" t="${t}" r="${r}" b="${b}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill></pic:pic></a:graphicData></a:graphic>` +
  `</wp:inline></w:drawing>`;

describe("image crop — a:srcRect", () => {
  it("parses a:srcRect (1/1000 percent) into 0..1 crop fractions", () => {
    const r = runImport(docxWithImage(`<w:p><w:r>${croppedDrawing(10000, 5000, 12500, 0)}</w:r></w:p>`)).doc;
    expect(firstImage(r).crop).toEqual({ left: 0.1, top: 0.05, right: 0.125, bottom: 0 });
  });

  it("ignores a degenerate / empty srcRect (no crop)", () => {
    const r = runImport(docxWithImage(`<w:p><w:r>${croppedDrawing(0, 0, 0, 0)}</w:r></w:p>`)).doc;
    expect(firstImage(r).crop).toBeUndefined();
  });

  it("emits a:srcRect on export and round-trips the crop", () => {
    const img: ImageBlock = { kind: "image", id: "img0", revision: 0, src: "img1", widthPx: 120, heightPx: 90, align: "left", crop: { left: 0.1, top: 0.2, right: 0.05, bottom: 0.15 } };
    const doc: Document = { section: SECTION, blocks: [img] };
    expect(exportedDocXml(doc, { img1: PNG_1PX })).toContain('<a:srcRect l="10000" t="20000" r="5000" b="15000"/>');
    expect(firstImage(roundTrip(doc, { img1: PNG_1PX })).crop).toEqual({ left: 0.1, top: 0.2, right: 0.05, bottom: 0.15 });
  });
});

// ── image rotation — pic spPr a:xfrm@rot (issue #236) ─────────────────────────
const rotatedDrawing = (rot: number): string =>
  `<w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>` +
  `<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm rot="${rot}"><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
  `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;

describe("image rotation — a:xfrm@rot", () => {
  it("parses pic spPr a:xfrm@rot (60000ths) into degrees", () => {
    const r = runImport(docxWithImage(`<w:p><w:r>${rotatedDrawing(1200000)}</w:r></w:p>`)).doc;
    expect(firstImage(r).rotation).toBe(20);
  });

  it("ignores rot=0 (no rotation)", () => {
    const r = runImport(docxWithImage(`<w:p><w:r>${rotatedDrawing(0)}</w:r></w:p>`)).doc;
    expect(firstImage(r).rotation).toBeUndefined();
  });

  it("emits a:xfrm@rot on export and round-trips the rotation (degrees ↔ 60000ths)", () => {
    const img: ImageBlock = { kind: "image", id: "img0", revision: 0, src: "img1", widthPx: 120, heightPx: 90, align: "left", rotation: 20 };
    const doc: Document = { section: SECTION, blocks: [img] };
    expect(exportedDocXml(doc, { img1: PNG_1PX })).toContain('<a:xfrm rot="1200000">');
    expect(firstImage(roundTrip(doc, { img1: PNG_1PX })).rotation).toBe(20);
  });

  it("omits the rot attribute when the image is unrotated", () => {
    const img: ImageBlock = { kind: "image", id: "img0", revision: 0, src: "img1", widthPx: 120, heightPx: 90, align: "left" };
    const xml = exportedDocXml({ section: SECTION, blocks: [img] }, { img1: PNG_1PX });
    expect(xml).toContain("<a:xfrm>");
    expect(xml).not.toContain("<a:xfrm rot=");
  });
});

// ── inline image paragraph spacing (#198) ────────────────────────────────────
describe("inline image paragraph spacing", () => {
  const inlineImage = (): ImageBlock => ({ kind: "image", id: "img0", revision: 0, src: "img1", widthPx: 150, heightPx: 110, align: "left" });

  it("pins the image paragraph to zero before/after and single (auto) line spacing", () => {
    // The editor lays an image out at exactly heightPx with no paragraph spacing, so
    // consecutive images render flush. Without an explicit spacing override the image
    // w:p inherits Normal's space-after + line multiplier and Word opens an unclickable
    // gap around the image; pinning it keeps Word flush with the canvas.
    const doc: Document = { section: SECTION, blocks: [inlineImage()] };
    const xml = exportedDocXml(doc, { img1: PNG_1PX });
    expect(xml).toContain('<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>');
    // w:spacing must precede w:jc in the CT_PPr schema sequence.
    expect(xml.indexOf("<w:spacing")).toBeLessThan(xml.indexOf("<w:jc"));
    // it is an inline drawing (no floating anchor), so the spacing applies to flow.
    expect(xml).toContain("wp:inline");
  });

  it("still round-trips the image itself", () => {
    const out = firstImage(roundTrip({ section: SECTION, blocks: [inlineImage()] }, { img1: PNG_1PX }));
    expect(out.widthPx).toBeCloseTo(150, 0);
    expect(out.heightPx).toBeCloseTo(110, 0);
  });
});

// ── anchored (floating) images inside a table cell ───────────────────────────
const anchoredDrawing = `<w:drawing><wp:anchor behindDoc="0" relativeHeight="2">` +
  `<wp:positionH relativeFrom="column"><wp:posOffset>91440</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>45720</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="914400" cy="457200"/><wp:wrapNone/>` +
  `<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>` +
  `</wp:anchor></w:drawing>`;

describe("anchored images inside table cells", () => {
  it("keeps an anchored image inside a table cell (not dropped)", () => {
    const body = `<w:tbl><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:p><w:r><w:t>caption</w:t></w:r></w:p><w:p><w:r>${anchoredDrawing}</w:r></w:p></w:tc></w:tr></w:tbl>`;
    const cell = firstTable(runImport(docxWithImage(body)).doc).rows[0]!.cells[0]!;
    const img = cell.blocks.find((b): b is ImageBlock => b.kind === "image");
    expect(img, "the anchored in-cell image should survive import").toBeTruthy();
    expect(img!.anchor).toBeTruthy();
    expect(img!.anchor!.behind).toBe(false);
  });

  it("round-trips an anchored in-cell image through export", () => {
    const cellImg: ImageBlock = {
      kind: "image", id: "ci0", revision: 0, src: "img1", widthPx: 96, heightPx: 48, align: "left",
      anchor: { behind: false, offsetXPx: 10, offsetYPx: 5, relFromH: "column", relFromV: "paragraph", z: 2 },
    };
    const tcell: TableCell = { id: "tc0", blocks: [para([run("caption")]), cellImg] };
    const tbl: TableBlock = { kind: "table", id: "T0", revision: 0, colFractions: [1], rows: [{ cells: [tcell] }] };
    const out = firstTable(roundTrip({ section: SECTION, blocks: [tbl] }, { img1: PNG_1PX }));
    const img = out.rows[0]!.cells[0]!.blocks.find((b): b is ImageBlock => b.kind === "image");
    expect(img, "the anchored in-cell image should survive the round-trip").toBeTruthy();
    expect(img!.anchor!.relFromH).toBe("column");
  });
});

// ── settings.xml: w:defaultTabStop + w:compat ────────────────────────────────
const docxWithSettings = (settings: string): Uint8Array =>
  makeDocx({
    "[Content_Types].xml": CONTENT_TYPES_XML,
    "word/document.xml": documentXml(`<w:p><w:r><w:t>hi</w:t></w:r></w:p>`),
    "word/settings.xml": `<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${settings}</w:settings>`,
  });

describe("settings.xml — w:defaultTabStop + w:compat", () => {
  it("parses a non-default w:defaultTabStop into px and round-trips compat settings", () => {
    const compat = `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>`;
    const r = runImport(docxWithSettings(`<w:defaultTabStop w:val="1440"/>${compat}`)).doc;
    expect(r.defaultTabStopPx).toBeCloseTo(96, 1); // 1440 twips = 1in = 96px
    expect(r.compatSettings).toEqual([{ name: "compatibilityMode", uri: "http://schemas.microsoft.com/office/word", val: "15" }]);
  });

  it("treats Word's 720-twip default as 'no override' (keeps the model field absent)", () => {
    const r = runImport(docxWithSettings(`<w:defaultTabStop w:val="720"/>`)).doc;
    expect(r.defaultTabStopPx).toBeUndefined();
  });

  it("does NOT emit settings a re-import would discard (empty-name compat / non-positive tab stop)", () => {
    const doc: Document = {
      section: SECTION,
      blocks: [para([run("hi")])],
      defaultTabStopPx: 0, // non-positive → must be skipped (matches parseSettings)
      compatSettings: [{ name: "", uri: "u", val: "1" }], // empty name → must be filtered
    };
    const xml = exportedSettingsXml(doc);
    expect(xml).not.toContain("w:defaultTabStop");
    expect(xml).not.toContain("w:compatSetting");
  });

  it("emits w:defaultTabStop + w:compatSetting on export and round-trips them", () => {
    const doc: Document = {
      section: SECTION,
      blocks: [para([run("hi")])],
      defaultTabStopPx: 72,
      compatSettings: [{ name: "compatibilityMode", uri: "http://schemas.microsoft.com/office/word", val: "15" }],
    };
    const xml = exportedSettingsXml(doc);
    expect(xml).toContain('<w:defaultTabStop w:val="1080"/>'); // 72px = 0.75in = 1080 twips
    expect(xml).toContain('<w:compatSetting w:name="compatibilityMode"');
    const out = roundTrip(doc);
    expect(out.defaultTabStopPx).toBeCloseTo(72, 1);
    expect(out.compatSettings).toEqual([{ name: "compatibilityMode", uri: "http://schemas.microsoft.com/office/word", val: "15" }]);
  });

  it("preserves w15:docId (accepting the w14 variant too)", () => {
    const guid = "{2E7A3C41-9B0D-4F1E-8A2B-1C3D4E5F6071}";
    const r15 = runImport(docxWithSettings(`<w15:docId w15:val="${guid}"/>`)).doc;
    expect(r15.docId).toBe(guid);
    // Older producers emit it under w14 — still accepted.
    const r14 = runImport(docxWithSettings(`<w14:docId w14:val="${guid}"/>`)).doc;
    expect(r14.docId).toBe(guid);
    // Export re-emits it (declaring the w15 namespace) and it round-trips.
    const xml = exportedSettingsXml({ section: SECTION, blocks: [para([run("hi")])], docId: guid });
    expect(xml).toContain('xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"');
    expect(xml).toContain(`<w15:docId w15:val="${guid}"/>`);
    expect(roundTrip({ section: SECTION, blocks: [para([run("hi")])], docId: guid }).docId).toBe(guid);
  });
});

// ── w14:paraId / w14:textId — persistent paragraph identity ───────────────────
describe("paragraph ids — w14:paraId / w14:textId", () => {
  it("parses w14:paraId/textId onto the paragraph and re-emits them", () => {
    const r = runImport(simpleDocx(`<w:p w14:paraId="12AB34CD" w14:textId="77777777"><w:r><w:t>hi</w:t></w:r></w:p>`)).doc;
    const p = firstPara(r);
    expect(p.paraId).toBe("12AB34CD");
    expect(p.textId).toBe("77777777");
    const xml = exportedDocXml(r);
    expect(xml).toContain('w14:paraId="12AB34CD"');
    expect(xml).toContain('w14:textId="77777777"');
    expect(firstPara(roundTrip(r)).paraId).toBe("12AB34CD");
  });

  it("a paragraph without an id emits no w14:paraId", () => {
    const doc: Document = { section: SECTION, blocks: [para([run("plain")])] };
    expect(exportedDocXml(doc)).not.toContain("w14:paraId");
  });

  it("de-dups a paraId shared by two blocks (copy/paste clone) — emits it once", () => {
    const a: Paragraph = { ...para([run("first")]), paraId: "AAAA0001" };
    const b: Paragraph = { ...para([run("clone")]), paraId: "AAAA0001" };
    const xml = exportedDocXml({ section: SECTION, blocks: [a, b] });
    expect(xml.match(/w14:paraId="AAAA0001"/g)?.length).toBe(1);
  });
});

// ── w14:checkedState / w14:uncheckedState — checkbox glyph fidelity ───────────
describe("checkbox glyphs — w14:checkedState / w14:uncheckedState", () => {
  const cbDoc = (): Document => ({
    section: SECTION,
    sdts: { cb: { type: "checkbox", checked: true, checkedSymbol: { font: "MS Gothic", val: "2612" }, uncheckedSymbol: { font: "MS Gothic", val: "2610" } } },
    blocks: [para([run("☒", { sdtPath: ["cb"] })])],
  });

  it("emits both states inside w14:checkbox and round-trips them", () => {
    const xml = exportedDocXml(cbDoc());
    expect(xml).toContain('<w14:checkedState w14:font="MS Gothic" w14:val="2612"/>');
    expect(xml).toContain('<w14:uncheckedState w14:font="MS Gothic" w14:val="2610"/>');
    const back = roundTrip(cbDoc());
    const props = Object.values(back.sdts!).find((s) => s.type === "checkbox")!;
    expect(props.checkedSymbol).toEqual({ font: "MS Gothic", val: "2612" });
    expect(props.uncheckedSymbol).toEqual({ font: "MS Gothic", val: "2610" });
  });

  it("upper-cases a lowercase hex code point on import", () => {
    const cb = `<w14:checkbox><w14:checked w14:val="1"/><w14:checkedState w14:font="Wingdings" w14:val="f0fe"/></w14:checkbox>`;
    const body = `<w:p><w:sdt><w:sdtPr><w:tag w:val="c"/>${cb}</w:sdtPr><w:sdtContent><w:r><w:t>☒</w:t></w:r></w:sdtContent></w:sdt></w:p>`;
    const doc = runImport(simpleDocx(body)).doc;
    const props = Object.values(doc.sdts!).find((s) => s.type === "checkbox")!;
    expect(props.checkedSymbol).toEqual({ font: "Wingdings", val: "F0FE" });
  });
});

// ── linked ("Link to File") images — a:blip r:link + External rel ─────────────
const exportedRelsXml = (doc: Document): string =>
  strFromU8(unzipSync(writeDocx(doc).bytes)["word/_rels/document.xml.rels"]!);

describe("linked (external) images — r:link round-trip", () => {
  const URL = "https://example.com/linked.png";
  const linkedImg: ImageBlock = {
    kind: "image", id: "lnk0", revision: 0, src: URL, externalSrc: URL, widthPx: 120, heightPx: 90, align: "left",
  };

  it("emits a:blip r:link + a TargetMode=External rel (no packed bytes)", () => {
    const doc: Document = { section: SECTION, blocks: [linkedImg] };
    // No images map entry needed — a linked image carries no bytes.
    const xml = exportedDocXml(doc);
    expect(xml).toContain("r:link=");
    expect(xml).not.toContain("r:embed=");
    const rels = exportedRelsXml(doc);
    expect(rels).toContain(`Target="${URL}"`);
    expect(rels).toContain('TargetMode="External"');
  });

  it("round-trips the external URL through export → re-import", () => {
    const doc: Document = { section: SECTION, blocks: [linkedImg] };
    const back = firstImage(roundTrip(doc));
    expect(back.src).toBe(URL);
    expect(back.externalSrc).toBe(URL);
  });
});

// ── wp14 anchored-drawing fidelity (issue #188) ──────────────────────────────
// (1) wp14 percentage positioning wraps the absolute offset in mc:AlternateContent,
//     so the importer must recover the mc:Fallback wp:posOffset (was read as 0).
// (2) wp14:anchorId/editId (drawing identity) round-trip, with export de-dup.
const wp14PositionedDrawing =
  `<w:drawing><wp:anchor behindDoc="0" relativeHeight="5">` +
  `<wp:positionH relativeFrom="page"><mc:AlternateContent>` +
  `<mc:Choice Requires="wp14"><wp14:pctPosHOffset>25000</wp14:pctPosHOffset></mc:Choice>` +
  `<mc:Fallback><wp:posOffset>1905000</wp:posOffset></mc:Fallback></mc:AlternateContent></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><mc:AlternateContent>` +
  `<mc:Choice Requires="wp14"><wp14:pctPosVOffset>10000</wp14:pctPosVOffset></mc:Choice>` +
  `<mc:Fallback><wp:posOffset>914400</wp:posOffset></mc:Fallback></mc:AlternateContent></wp:positionV>` +
  `<wp:extent cx="914400" cy="457200"/><wp:wrapNone/>` +
  `<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>` +
  `</wp:anchor></w:drawing>`;

describe("wp14 anchored-drawing fidelity", () => {
  it("recovers the mc:Fallback wp:posOffset for a wp14 percent-positioned float (was 0)", () => {
    const img = firstImage(runImport(docxWithImage(`<w:p><w:r>${wp14PositionedDrawing}</w:r></w:p>`)).doc);
    expect(img.anchor).toBeTruthy();
    // 1905000 EMU = 200px, 914400 EMU = 96px — both read as 0 before the fix.
    expect(img.anchor!.offsetXPx).toBeCloseTo(200, 0);
    expect(img.anchor!.offsetYPx).toBeCloseTo(96, 0);
  });

  it("parses wp14:anchorId/editId off an inline image", () => {
    const drawing = `<w:drawing><wp:inline wp14:anchorId="12AB34CD" wp14:editId="55667788">` +
      `<wp:extent cx="914400" cy="914400"/>` +
      `<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>` +
      `</wp:inline></w:drawing>`;
    const r = runImport(docxWithImage(`<w:p><w:r>${drawing}</w:r></w:p>`)).doc;
    expect(firstImage(r).drawingId).toEqual({ anchorId: "12AB34CD", editId: "55667788" });
  });

  const idImg: ImageBlock = {
    kind: "image", id: "di0", revision: 0, src: "img1", widthPx: 96, heightPx: 96, align: "left",
    drawingId: { anchorId: "12AB34CD", editId: "55667788" },
  };

  it("emits wp14:anchorId/editId (+ namespace) and round-trips them", () => {
    const doc: Document = { section: SECTION, blocks: [idImg] };
    const xml = exportedDocXml(doc, { img1: PNG_1PX });
    expect(xml).toContain('xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"');
    expect(xml).toContain('wp14:anchorId="12AB34CD"');
    expect(xml).toContain('wp14:editId="55667788"');
    expect(firstImage(roundTrip(doc, { img1: PNG_1PX })).drawingId).toEqual({ anchorId: "12AB34CD", editId: "55667788" });
  });

  it("de-dups a drawing anchorId shared by two images (copy/paste clone) — emits it once", () => {
    const xml = exportedDocXml({ section: SECTION, blocks: [{ ...idImg, id: "a" }, { ...idImg, id: "b" }] }, { img1: PNG_1PX });
    expect(xml.match(/wp14:anchorId="12AB34CD"/g)?.length).toBe(1);
  });
});
