// Flagship showcase document — the editor's initial state when no docId is opened.
// It exercises nearly the full Word feature surface (TOC, fields, content controls,
// complex/merged + cross-page tables, images, lists, footnotes, bookmarks, hidden
// text, headers/footers) so a first-time visitor sees how much canvas-word supports.
// Built as plain model data (the same Document the editor/exporter/collab consume).

import type {
  BookmarkRange, CharStyle, Document, FieldDef, FieldSpec, ImageBlock, ParaStyle, Paragraph, Run, SdtProps, TableBlock, TableCell,
} from "@cw/shared";
import { buildInstruction, buildTocParagraphs, DEFAULT_CHAR_STYLE, DEFAULT_PARA_STYLE, defaultStylesheet, evaluateIf, formatFieldDate } from "@cw/shared";
import { defaultListDefinition, DEFAULT_BULLET_LIST_ID, DEFAULT_NUMBER_LIST_ID } from "@cw/shared";

const BODY: CharStyle = DEFAULT_CHAR_STYLE;
const PARA: ParaStyle = DEFAULT_PARA_STYLE;

let nextId = 0;
const id = (): string => `b${nextId++}`;
const run = (text: string, patch: Partial<CharStyle> = {}): Run => ({ text, style: { ...BODY, ...patch } });
const para = (runs: Run[], patch: Partial<ParaStyle> = {}): Paragraph => ({ kind: "paragraph", id: id(), revision: 0, runs, style: { ...PARA, ...patch } });

// --- registries the document references ---------------------------------------
const fields: Record<string, FieldDef> = {};
const sdts: Record<string, SdtProps> = {};
const footnotes: Record<string, Paragraph[]> = {};
const bookmarks: Record<string, BookmarkRange> = {};
const tocItems: { id: string; text: string; level: number }[] = [];

let fldN = 0;
/** An inline field: a fieldId-tagged result run + a registered builtin FieldDef. */
const fieldRun = (spec: FieldSpec, text: string, patch: Partial<CharStyle> = {}): Run => {
  const fid = `fld${fldN++}`;
  fields[fid] = { id: fid, instruction: buildInstruction(spec), name: spec.type, kind: "builtin", spec };
  return { text, style: { ...BODY, ...patch, fieldId: fid } };
};
const pageField = (): Run => fieldRun({ type: "PAGE" }, "{page}");
const dateField = (fmt: string): Run => fieldRun({ type: "DATE", format: fmt }, formatFieldDate(new Date(), fmt));
const ifField = (a: string, op: "=" | "<>" | "<" | ">" | "<=" | ">=", b: string, t: string, f: string): Run => {
  const spec: FieldSpec = { type: "IF", operandA: a, op, operandB: b, trueRuns: [run(t)], falseRuns: [run(f)] };
  return fieldRun(spec, evaluateIf(spec).map((r) => r.text).join(""));
};

let sdtN = 0;
/** Register a content control's props and return its id (for nesting / block-level use). */
const sdtId = (props: SdtProps): string => {
  const sid = `sdt${sdtN++}`;
  sdts[sid] = props;
  return sid;
};
/** An inline content control: an sdtPath-tagged run + registered SdtProps. */
const sdtRun = (props: SdtProps, text: string, patch: Partial<CharStyle> = {}): Run =>
  ({ text, style: { ...BODY, ...patch, sdtPath: [sdtId(props)] } });

const heading = (text: string, level: 1 | 2 | 3): Paragraph => {
  const h = para([run(text, { bold: true, fontSizePx: level === 1 ? 24 : level === 2 ? 19 : 16, color: "#1a1a2e" })], {
    namedStyle: `Heading${level}`, outlineLevel: level - 1, spaceBeforePx: 18, spaceAfterPx: 8,
  });
  tocItems.push({ id: h.id, text, level });
  return h;
};

// --- images --------------------------------------------------------------------
const SVG = (label: string, w: number, h: number): string =>
  "data:image/svg+xml," + encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}">` +
    `<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#1a73e8"/><stop offset="1" stop-color="#9c27b0"/></linearGradient></defs>` +
    `<rect width="${w}" height="${h}" rx="8" fill="url(#g)"/><text x="${w / 2}" y="${h / 2 + 6}" font-family="Arial" font-size="18" fill="#fff" text-anchor="middle">${label}</text></svg>`,
  );
const image = (label: string, w: number, h: number, align: ImageBlock["align"], wrap?: "block" | "square"): ImageBlock => ({
  kind: "image", id: id(), revision: 0, src: SVG(label, w, h), widthPx: w, heightPx: h, align, ...(wrap ? { wrap } : {}) });

// --- tables --------------------------------------------------------------------
const cellPara = (text: string, patch: Partial<CharStyle> = {}, runs?: Run[]): Paragraph => ({
  kind: "paragraph", id: id(), revision: 0, runs: runs ?? [run(text, { fontSizePx: 14, ...patch })], style: { ...PARA, lineHeight: 1.35, spaceAfterPx: 0 } });
const cell = (text: string, patch: Partial<CharStyle> = {}, opts: Partial<TableCell> = {}): TableCell => ({ id: id(), blocks: [cellPara(text, patch)], ...opts });

/** Merged cells: a 3-col header spanning all columns, then a row-spanning cell. */
const mergedTable = (): TableBlock => ({
  kind: "table", id: id(), revision: 0,
  rows: [
    { cells: [cell("Merged header spanning all three columns", { bold: true, color: "#fff" }, { colSpan: 3, shading: "#1a73e8" })] },
    { cells: [cell("Spans\n2 rows", { bold: true }, { rowSpan: 2, shading: "#e8f0fe" }), cell("B1"), cell("C1")] },
    { cells: [cell("B2"), cell("C2", { color: "#188038" })] },
  ],
});

/** A table tall enough to cross a page boundary (row-level pagination). */
const tallTable = (): TableBlock => ({
  kind: "table", id: id(), revision: 0,
  rows: [
    { cells: [cell("#", { bold: true }), cell("Feature", { bold: true }), cell("Status", { bold: true })] },
    ...Array.from({ length: 22 }, (_v, i) => ({
      cells: [cell(String(i + 1)), cell(`Demonstrated capability number ${i + 1} that keeps the table running past the bottom of the page`), cell("✓ supported", { color: "#188038" })],
    })),
  ],
});

/** AutoFit to Contents: columns are solved from cell content, so the table
 *  shrinks below the page width to fit (narrow ID, wider Notes). */
const autofitTable = (): TableBlock => ({
  kind: "table", id: id(), revision: 0, widthMode: "autofitContents",
  rows: [
    { cells: [cell("ID", { bold: true }), cell("Name", { bold: true }), cell("Notes", { bold: true })] },
    { cells: [cell("1"), cell("Ada Lovelace"), cell("first programmer")] },
    { cells: [cell("2"), cell("Grace Hopper"), cell("compiler pioneer")] },
    { cells: [cell("3"), cell("Linus Torvalds"), cell("kernel maintainer")] },
  ],
});

const LOREM = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ";

export function sampleDoc(): Document {
  // Footnote body referenced by a marker run below.
  footnotes["fn1"] = [para([run("Footnotes lay out at the bottom of their page, with a separator rule — just like Word.", { fontSizePx: 12 })], { spaceAfterPx: 0 })];

  // --- body content (headings register themselves for the TOC) -----------------
  const fieldsHeading = heading("Fields", 1);
  const fieldsPara = para([
    run("Fields are first-class objects you can insert and edit (right-click → Insert Field). Today is "),
    dateField("MMMM d, yyyy"),
    run(", this is page "),
    pageField(),
    run(". A conditional (IF) field can branch: "),
    ifField("2", ">", "1", "the condition held", "it did not"),
    run(". Each is outlined like a content control and recomputes on Update Field."),
  ]);

  const ccHeading = heading("Content controls", 1);
  const ccPara = para([
    run("Every Word content-control kind round-trips: "),
    sdtRun({ type: "richText", alias: "Rich text" }, "rich text"),
    run(", "),
    sdtRun({ type: "plainText", alias: "Plain text" }, "plain text"),
    run(", a dropdown "),
    sdtRun({ type: "dropDown", alias: "Choice", listItems: [{ display: "One", value: "1" }, { display: "Two", value: "2" }] }, "One"),
    run(", a combo box "),
    sdtRun({ type: "comboBox", alias: "Combo", listItems: [{ display: "Alpha", value: "a" }, { display: "Beta", value: "b" }, { display: "Gamma", value: "g" }] }, "Alpha"),
    run(", a date picker "),
    sdtRun({ type: "date", alias: "Pick a date", dateFormat: "M/d/yyyy" }, "6/16/2026"),
    run(", and a checkbox "),
    sdtRun({ type: "checkbox", checked: true }, "☒"),
    run("."),
  ]);

  // Nested controls: an outer control wrapping an inner one (runs share a path
  // prefix). And a block-level "section" control wrapping a whole paragraph + a
  // table whose value cell holds its own inner control (the report-style pattern).
  const ccOuter = sdtId({ type: "richText", alias: "Outer control" });
  const ccInner = sdtId({ type: "richText", alias: "Inner control" });
  const ccNestedPara = para([
    run("Controls can also nest: here an "),
    run("outer control wraps ", { sdtPath: [ccOuter] }),
    run("an inner control", { sdtPath: [ccOuter, ccInner] }),
    run(" and trailing text", { sdtPath: [ccOuter] }),
    run("."),
  ]);

  const ccSection = sdtId({ type: "richText", alias: "Section (block-level control)" });
  const ccFee = sdtId({ type: "richText", alias: "Appraisal Fee" });
  const ccSectionIntro: Paragraph = {
    ...para([run("A block-level control can wrap whole paragraphs and tables — this paragraph and the table below are one control, with an inner control around just the value:")], { spaceAfterPx: 6 }),
    sdtPath: [ccSection],
  };
  const ccSectionTable: TableBlock = {
    kind: "table", id: id(), revision: 0, sdtPath: [ccSection],
    rows: [
      { cells: [cell("Field", { bold: true }), cell("Value", { bold: true })] },
      {
        cells: [
          cell("Appraisal Fee", {}),
          { id: id(), blocks: [cellPara("", {}, [run("$200.00", { fontSizePx: 14, sdtPath: [ccFee] })])] },
        ],
      },
    ],
  };

  const tablesHeading = heading("Tables", 1);
  const fieldInCellTable: TableBlock = {
    kind: "table", id: id(), revision: 0,
    rows: [
      { cells: [cell("Metric", { bold: true }), cell("Value", { bold: true })] },
      { cells: [cell("Rendered on", {}), { id: id(), blocks: [cellPara("", {}, [dateField("yyyy-MM-dd")])] }] },
      { cells: [cell("Page", {}), { id: id(), blocks: [cellPara("", {}, [run("p. ", { fontSizePx: 14 }), pageField()])] }] },
    ],
  };

  const richHeading = heading("Rich text, lists & images", 1);

  const bodyBlocks: (Paragraph | TableBlock | ImageBlock)[] = [
    fieldsHeading,
    fieldsPara,
    para([
      run("Footnotes", { bold: true }), run(" are supported too"),
      run("1", { footnoteRef: "fn1", verticalAlign: "super", fontSizePx: 11 }),
      run(". So is "), run("hidden metadata", { hidden: true }), run("bookmarked text", {}),
      run(" and inline formatting: "),
      run("bold", { bold: true }), run(", "), run("italic", { italic: true }), run(", "),
      run("underline", { underline: true }), run(", "), run("strike", { strikethrough: true }), run(", "),
      run("highlight", { highlightColor: "#fff3a3" }), run(", "), run("x", {}), run("2", { verticalAlign: "super", fontSizePx: 11 }),
      run(", and a "), run("hyperlink", { link: "https://forevka.dev", color: "#0b57d0", underline: true }), run("."),
    ]),

    ccHeading,
    ccPara,
    ccNestedPara,
    ccSectionIntro,
    ccSectionTable,

    tablesHeading,
    para([run("Merged cells (column- and row-spanning), shading and borders:")], { spaceAfterPx: 6 }),
    mergedTable(),
    para([run("Fields work inside table cells too:")], { spaceBeforePx: 10, spaceAfterPx: 6 }),
    fieldInCellTable,
    para([run("And a table tall enough to paginate across pages — rows break cleanly:")], { spaceBeforePx: 10, spaceAfterPx: 6 }),
    tallTable(),
    para([run("AutoFit to Contents — columns are solved from cell content so the table shrinks to fit (Table → AutoFit, or drag a border to pin it back to fixed widths):")], { spaceBeforePx: 10, spaceAfterPx: 6 }),
    autofitTable(),

    richHeading,
    para([run("An inline block image:")], { spaceAfterPx: 6 }),
    image("block image", 360, 110, "center", "block"),
    para([
      run("A square-wrapped image floats and text flows around it. " + LOREM.repeat(3)),
    ]),
    image("square wrap", 150, 110, "left", "square"),
    para([run("Multilevel numbered list:")], { spaceBeforePx: 8, spaceAfterPx: 4 }),
    para([run("Model — invertible ops")], { list: { listId: DEFAULT_NUMBER_LIST_ID, level: 0 }, spaceAfterPx: 2 }),
    para([run("Layout — pretext pagination")], { list: { listId: DEFAULT_NUMBER_LIST_ID, level: 0 }, spaceAfterPx: 2 }),
    para([run("line caches keyed by (revision, width)")], { list: { listId: DEFAULT_NUMBER_LIST_ID, level: 1 }, spaceAfterPx: 2 }),
    para([run("Paint — one fillText per fragment")], { list: { listId: DEFAULT_NUMBER_LIST_ID, level: 0 }, spaceAfterPx: 2 }),
    para([run("Bulleted list:")], { spaceBeforePx: 8, spaceAfterPx: 4 }),
    para([run("markers are paint-only")], { list: { listId: DEFAULT_BULLET_LIST_ID, level: 0 }, spaceAfterPx: 2 }),
    para([run("so caches survive renumbering")], { list: { listId: DEFAULT_BULLET_LIST_ID, level: 1 }, spaceAfterPx: 2 }),
    para([run("A fully justified, multi-page paragraph exercises line-level pagination. " + LOREM.repeat(12))], { align: "justify" }),
    para([run("複雑なスクリプトの行分割も Intl.Segmenter が処理します。日本語のテキストは単語間にスペースがありませんが、pretext は文節の境界を正しく検出して行を折り返します。")]),
    para([run("— a tour of canvas-word —", { italic: true, color: "#5f6368" })], { align: "center", spaceBeforePx: 20 }),
  ];

  // Bookmark the literal "bookmarked text" run inside the footnotes/formatting
  // paragraph (bodyBlocks[2]); it sits at offset 51..66 of that paragraph's text
  // ("Footnotes" + " are supported too" + "1" + ". So is " + "hidden metadata").
  bookmarks["sample"] = { start: { blockId: bodyBlocks[2]!.id, offset: 51 }, end: { blockId: bodyBlocks[2]!.id, offset: 66 } };

  // --- Table of Contents (generated from the registered headings) --------------
  const tocHostDoc: Document = { section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } }, blocks: bodyBlocks };
  const tocEntries = buildTocParagraphs(tocHostDoc, { title: null, maxLevel: 3, leader: "dot" });

  const blocks: Document["blocks"] = [
    para([run("canvas-word", { fontFamily: "Arial, sans-serif", fontSizePx: 32, bold: true, color: "#1a1a2e" })], { align: "center", spaceAfterPx: 4, namedStyle: "Title" }),
    para([run("a canvas-rendered, page-accurate Word editor — feature showcase", { italic: true, color: "#5f6368" })], { align: "center", spaceAfterPx: 24, namedStyle: "Subtitle" }),
    para([run("Table of Contents", { bold: true, fontSizePx: 20, color: "#1a1a2e" })], { spaceAfterPx: 8 }),
    ...tocEntries,
    ...bodyBlocks,
  ];

  const doc: Document = {
    stylesheet: defaultStylesheet(),
    lists: { [DEFAULT_BULLET_LIST_ID]: defaultListDefinition("bullet"), [DEFAULT_NUMBER_LIST_ID]: defaultListDefinition("decimal") },
    section: {
      pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
      header: [para([run("canvas-word", { fontFamily: "Arial, sans-serif", fontSizePx: 11, bold: true, color: "#5f6368" }), run("  ·  feature showcase", { fontFamily: "Arial, sans-serif", fontSizePx: 11, color: "#9aa0a6" })], { spaceAfterPx: 0 })],
      footer: [para([run("Page {page} of {pages}", { fontFamily: "Arial, sans-serif", fontSizePx: 11, color: "#9aa0a6" })], { align: "center", spaceAfterPx: 0 })],
    },
    blocks,
    fields,
    sdts,
    footnotes,
    bookmarks,
    tocInstruction: ' TOC \\o "1-3" \\h \\z ',
  };
  return doc;
}
