// Flagship showcase document — the editor's initial state when no docId is opened.
// It exercises nearly the full Word feature surface (TOC, fields, content controls,
// complex/merged + cross-page tables, images, lists, footnotes, bookmarks, hidden
// text, headers/footers) so a first-time visitor sees how much canvas-word supports.
// Built as plain model data (the same Document the editor/exporter/collab consume).

import type {
  BookmarkRange, CellBorder, CellMargin, CharStyle, Document, EquationBlock, FieldDef, FieldSpec, ImageBlock, ParaStyle, Paragraph, RowProps, Run, SdtProps, TableBlock, TableCell,
} from "@cw/shared";
import { buildInstruction, buildTocParagraphs, DEFAULT_CHAR_STYLE, DEFAULT_PARA_STYLE, defaultStylesheet, evaluateIf, formatFieldDate } from "@cw/shared";
import { defaultListDefinition, DEFAULT_BULLET_LIST_ID, DEFAULT_NUMBER_LIST_ID } from "@cw/shared";
import { parseMathml } from "../mathml/parse";

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

// --- equations (MathML) --------------------------------------------------------
/** A display equation block from a MathML string. Stored as the MathML AST (the
 *  canonical form); typeset by the layout engine and round-tripped to .docx as
 *  OMML. `display: true` centers it on its own line like Word's block math. */
const eq = (mathml: string, align: EquationBlock["align"] = "center"): EquationBlock => ({
  kind: "equation", id: id(), revision: 0, equation: { ...parseMathml(mathml), display: true }, align,
});
/** Caption + monospace MathML source, so the demo SHOWS the MathML behind a render. */
const mathmlSource = (xml: string): Run => run(xml, { fontFamily: "Consolas, monospace", fontSizePx: 11, color: "#5f6368" });
/** An INLINE equation run: a single U+FFFC carrying the MathML, sized to the text. */
const inlineEq = (xml: string): Run => ({ text: "￼", style: { ...BODY, equation: { ...parseMathml(xml), display: false } } });

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

/** Cell vertical alignment (w:vAlign): a tall first cell forces a tall row, and the
 *  short label cells beside it sit top / centered / bottom within that height. */
const vAlignTable = (): TableBlock => ({
  kind: "table", id: id(), revision: 0,
  colFractions: [0.4, 0.2, 0.2, 0.2],
  rows: [
    { cells: [
      cell("vAlign", { bold: true, color: "#fff" }, { shading: "#1a73e8" }),
      cell("top", { bold: true, color: "#fff" }, { shading: "#1a73e8" }),
      cell("center", { bold: true, color: "#fff" }, { shading: "#1a73e8" }),
      cell("bottom", { bold: true, color: "#fff" }, { shading: "#1a73e8" }),
    ] },
    { cells: [
      cell("This tall cell holds several lines of text so the row grows well past the height of a single line — making the vertical position of its short neighbours visible.\nLine two.\nLine three.\nLine four."),
      cell("top", { color: "#188038" }, { vAlign: "top" }),
      cell("center", { color: "#188038" }, { vAlign: "center" }),
      cell("bottom", { color: "#188038" }, { vAlign: "bottom" }),
    ] },
  ],
});

/** A cell paragraph carrying an explicit base direction (RTL) — for the bidi
 *  table demo, so the cell's text right-aligns and reorders inside its column. */
const dirCellPara = (runs: Run[], direction?: "rtl"): Paragraph => ({
  kind: "paragraph", id: id(), revision: 0, runs,
  style: { ...PARA, lineHeight: 1.35, spaceAfterPx: 0, ...(direction ? { direction } : {}) },
});
const dirCell = (runs: Run[], direction?: "rtl", opts: Partial<TableCell> = {}): TableCell =>
  ({ id: id(), blocks: [dirCellPara(runs, direction)], ...opts });

/** Bidi + CJK inside a table: an RTL Arabic column, a Japanese column, and a
 *  value cell that combines a content control with a live PAGE field — proving
 *  direction composes with tables, SDTs, and fields. */
const bidiCjkTable = (): TableBlock => ({
  kind: "table", id: id(), revision: 0,
  rows: [
    { cells: [
      dirCell([run("العربية (RTL)", { bold: true, fontSizePx: 14, color: "#fff" })], "rtl", { shading: "#1a73e8" }),
      dirCell([run("日本語 (CJK)", { bold: true, fontSizePx: 14, color: "#fff" })], undefined, { shading: "#1a73e8" }),
      dirCell([run("Control + field", { bold: true, fontSizePx: 14, color: "#fff" })], undefined, { shading: "#1a73e8" }),
    ] },
    { cells: [
      dirCell([run("نص عربي داخل خلية، يُحاذى إلى اليمين تلقائيًا.", { fontSizePx: 14 })], "rtl"),
      dirCell([run("日本語のセル。スペースなしで折り返し、禁則処理も働きます。", { fontSizePx: 14 })]),
      dirCell([
        sdtRun({ type: "plainText", alias: "RTL value" }, "قيمة قابلة للتحرير", { fontSizePx: 14 }),
        run(" · ص ", { fontSizePx: 14 }),
        pageField(),
      ], "rtl"),
    ] },
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

/** Row properties (w:trPr): the first row is marked `repeatHeader` so it re-draws
 *  at the top of every page the table crosses; one row is pinned to an EXACT
 *  height; the data rows are `cantSplit` (kept whole). Tall enough to paginate so
 *  the repeating header is visible on the continuation page. */
const rowPropsTable = (): TableBlock => {
  const head = (text: string): TableCell => cell(text, { bold: true, color: "#fff" }, { shading: "#1a73e8" });
  const headerProps: RowProps = { repeatHeader: true };
  return {
    kind: "table", id: id(), revision: 0,
    colFractions: [0.1, 0.6, 0.3],
    rows: [
      { cells: [head("#"), head("Row property"), head("Effect")], props: headerProps },
      { cells: [cell("0"), cell("trHeight — exact 44px"), cell("forced to exactly 44px tall", { color: "#188038" })],
        props: { height: { value: 44, rule: "exact" } } },
      ...Array.from({ length: 20 }, (_v, i) => ({
        cells: [cell(String(i + 1)), cell(`cantSplit data row ${i + 1} — kept whole across a page break`), cell("✓", { color: "#188038" })],
        props: { cantSplit: true } as RowProps,
      })),
    ],
  };
};

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

/** Table-level defaults (issue #48): w:tblBorders / w:shd / w:tblCellMar are stored
 *  at tblPr level and round-trip as table-WIDE defaults — a Word→edit→Word cycle no
 *  longer drops them. The blue grid, light fill and roomy padding here all come from
 *  the table-level defaults (cascaded onto the cells for the canvas renderer, which
 *  reads concrete per-cell props), not from per-cell formatting. */
const tableDefaultsTable = (): TableBlock => {
  const rule: CellBorder = { color: "#1a73e8", widthPx: 1 };
  const fill = "#eef5ff";
  const pad: CellMargin = { top: 6, right: 10, bottom: 6, left: 10 };
  // Uniform border box (same rule on every edge) — the cascade of the table-level
  // default onto each cell, so the canvas renderer draws the blue grid.
  const c = (text: string, patch: Partial<CharStyle> = {}): TableCell =>
    cell(text, patch, { shading: fill, margin: { ...pad }, borders: { top: rule, right: rule, bottom: rule, left: rule } });
  return {
    kind: "table", id: id(), revision: 0,
    defaultBorders: { top: rule, right: rule, bottom: rule, left: rule, insideH: rule, insideV: rule },
    defaultShading: fill,
    defaultCellMargin: { ...pad },
    rows: [
      { cells: [c("Table-level default", { bold: true }), c("OOXML carrier (tblPr)", { bold: true })] },
      { cells: [c("Borders (outer + interior)"), c("w:tblBorders")] },
      { cells: [c("Shading fill"), c("w:shd")] },
      { cells: [c("Cell margins / padding"), c("w:tblCellMar")] },
    ],
  };
};

// MathML sources for the equations demo (Presentation MathML, the W3C standard).
const MATH_QUADRATIC =
  "<math><mi>x</mi><mo>=</mo><mfrac><mrow><mo>-</mo><mi>b</mi><mo>±</mo>" +
  "<msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>-</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt>" +
  "</mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>";
const MATH_SUM =
  "<math><munderover><mo>∑</mo><mrow><mi>n</mi><mo>=</mo><mn>1</mn></mrow><mo>∞</mo></munderover>" +
  "<mfrac><mn>1</mn><msup><mi>n</mi><mn>2</mn></msup></mfrac><mo>=</mo>" +
  "<mfrac><msup><mi>π</mi><mn>2</mn></msup><mn>6</mn></mfrac></math>";
const MATH_INTEGRAL =
  "<math><munderover><mo>∫</mo><mrow><mo>-</mo><mi>∞</mi></mrow><mo>∞</mo></munderover>" +
  "<msup><mi>e</mi><mrow><mo>-</mo><msup><mi>x</mi><mn>2</mn></msup></mrow></msup>" +
  "<mspace width=\"0.2em\"/><mi>d</mi><mi>x</mi><mo>=</mo><msqrt><mi>π</mi></msqrt></math>";
const MATH_EULER =
  "<math><msup><mi>e</mi><mrow><mi>i</mi><mi>π</mi></mrow></msup><mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn></math>";
const MATH_MATRIX =
  "<math><mfenced open=\"[\" close=\"]\"><mtable>" +
  "<mtr><mtd><mn>1</mn></mtd><mtd><mn>0</mn></mtd></mtr>" +
  "<mtr><mtd><mn>0</mn></mtd><mtd><mn>1</mn></mtd></mtr></mtable></mfenced></math>";

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

  const bodyBlocks: (Paragraph | TableBlock | ImageBlock | EquationBlock)[] = [
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
    // Underline styles + colors (w:u val + color) — double/dotted/dashed/wave/thick.
    para([
      run("Underlines carry a "), run("style", { italic: true }), run(" and an optional "), run("color", { italic: true }), run(": "),
      run("double", { underline: true, underlineStyle: "double" }), run(", "),
      run("dotted", { underline: true, underlineStyle: "dotted" }), run(", "),
      run("dashed", { underline: true, underlineStyle: "dash" }), run(", "),
      run("dot-dash", { underline: true, underlineStyle: "dotDash" }), run(", "),
      run("thick", { underline: true, underlineStyle: "thick" }), run(", "),
      run("a red wavy", { underline: true, underlineStyle: "wave", underlineColor: "#d93025" }), run(", and "),
      run("a blue double", { underline: true, underlineStyle: "double", underlineColor: "#1a73e8" }),
      run(" — each round-trips through Word's w:u (style + color)."),
    ]),

    ccHeading,
    ccPara,
    ccNestedPara,
    ccSectionIntro,
    ccSectionTable,

    tablesHeading,
    para([run("Merged cells (column- and row-spanning), shading and borders:")], { spaceAfterPx: 6 }),
    mergedTable(),
    para([run("Cell vertical alignment (w:vAlign) — short labels sit top, centered and bottom within a tall row:")], { spaceBeforePx: 10, spaceAfterPx: 6 }),
    vAlignTable(),
    para([run("Fields work inside table cells too:")], { spaceBeforePx: 10, spaceAfterPx: 6 }),
    fieldInCellTable,
    para([run("And a table tall enough to paginate across pages — rows break cleanly:")], { spaceBeforePx: 10, spaceAfterPx: 6 }),
    tallTable(),
    para([run("Row properties (w:trPr) — a repeating header row (re-drawn atop each page), an exact-height row, and cant-split data rows kept whole:")], { spaceBeforePx: 10, spaceAfterPx: 6 }),
    rowPropsTable(),
    para([run("AutoFit to Contents — columns are solved from cell content so the table shrinks to fit (Table → AutoFit, or drag a border to pin it back to fixed widths):")], { spaceBeforePx: 10, spaceAfterPx: 6 }),
    autofitTable(),
    para([run("Table-level defaults — borders, a shading fill and cell padding set once on the table (w:tblBorders / w:shd / w:tblCellMar) and round-tripped at that level instead of being baked onto every cell:")], { spaceBeforePx: 10, spaceAfterPx: 6 }),
    tableDefaultsTable(),

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

    // --- Paragraph borders & shading (w:pBdr / paragraph w:shd) -----------------
    heading("Paragraph borders & shading", 1),
    para([run("A whole paragraph can carry a border box and a background fill — Word's w:pBdr and paragraph-level w:shd. The box hugs the paragraph between its indents and round-trips to .docx and PDF.")], {
      spaceBeforePx: 6,
      borders: {
        top: { color: "#1a73e8", widthPx: 1 },
        bottom: { color: "#1a73e8", widthPx: 1 },
        left: { color: "#1a73e8", widthPx: 1 },
        right: { color: "#1a73e8", widthPx: 1 },
      },
      shading: "#eef4ff",
    }),
    para([run("Borders and shading are independent: this indented paragraph combines a warm shading fill with a single thick double-ruled accent on its left edge only — each edge of the box is configured on its own.")], {
      spaceBeforePx: 6,
      indentLeftPx: 24,
      indentRightPx: 24,
      shading: "#fff3e0",
      borders: { left: { color: "#e8710a", widthPx: 3, style: "double" } },
    }),
    // --- Contextual spacing: same-style runs sit tight (w:contextualSpacing) ---
    para([run("Contextual spacing — each verse line below carries 12px after-spacing, yet w:contextualSpacing collapses the gaps between adjacent same-style paragraphs (Word's list-style default); only the run's outer edges keep their space:")], { spaceBeforePx: 10, spaceAfterPx: 4 }),
    para([run("Roses are red,")], { contextualSpacing: true, spaceAfterPx: 12 }),
    para([run("violets are blue,")], { contextualSpacing: true, spaceAfterPx: 12 }),
    para([run("contextual spacing keeps these lines tight,")], { contextualSpacing: true, spaceAfterPx: 12 }),
    para([run("the way Word's list paragraphs do.")], { contextualSpacing: true, spaceAfterPx: 12 }),

    // --- International text: CJK + bidirectional (RTL) -------------------------
    heading("International text — CJK & bidirectional", 1),
    para([run("East-Asian and right-to-left scripts lay out the way Word does — measured on canvas, with Unicode line-breaking and the bidirectional algorithm (UAX #9), not the browser's contenteditable.")], { spaceAfterPx: 8 }),

    para([run("日本語 — CJK line-breaking & kinsoku", { bold: true, color: "#1a1a2e" })], { spaceBeforePx: 6, spaceAfterPx: 2 }),
    para([run("日本語の文章は単語の間にスペースを入れません。それでもエンジンは文字単位で行を折り返し、句読点が行頭に来ないように禁則処理（kinsoku）を行います。「角括弧」のような約物も正しく扱われ、長い段落でもページをまたいで自然に流れます。")]),

    para([run("العربية — right-to-left", { bold: true, color: "#1a1a2e" })], { spaceBeforePx: 8, spaceAfterPx: 2 }),
    para([run("اللغة العربية تُكتب من اليمين إلى اليسار. يعيد المحرّر ترتيب النص بصريًا وفق خوارزمية يونيكود ثنائية الاتجاه، ويحاذي الفقرة إلى اليمين تلقائيًا، ويضع المؤشر في المكان الصحيح عند الكتابة والتحديد.")], { direction: "rtl" }),

    para([run("עברית — right-to-left", { bold: true, color: "#1a1a2e" })], { spaceBeforePx: 8, spaceAfterPx: 2 }),
    para([run("עברית נכתבת מימין לשמאל. העורך מסדר מחדש את הרצף החזותי, מיישר את הפסקה לימין כברירת מחדל, וממשיך לתמוך בעימוד מרובה עמודים.")], { direction: "rtl" }),

    para([run("Nested bidi — numbers & Latin inside RTL", { bold: true, color: "#1a1a2e" })], { spaceBeforePx: 8, spaceAfterPx: 2 }),
    para([run("المنتج «canvas-word» متوفر بسعر 1,299 درهمًا منذ عام 2026 — تبقى الأرقام والكلمات اللاتينية بترتيبها الصحيح من اليسار إلى اليمين داخل النص العربي.")], { direction: "rtl" }),
    para([
      run("And the reverse, inside this left-to-right line: an embedded Hebrew phrase "),
      run("שלום עולם"),
      run(" and an Arabic one "),
      run("مرحبا بالعالم"),
      run(" each reorder on their own while the English keeps reading left-to-right — caret, selection, and arrow keys follow the visual order."),
    ]),

    para([run("Tables, content controls & fields — in CJK / RTL", { bold: true, color: "#1a1a2e" })], { spaceBeforePx: 10, spaceAfterPx: 4 }),
    para([run("Direction composes with every other feature. This table mixes a right-to-left Arabic column with a Japanese one, and its last cell holds a content control plus a live page field:")], { spaceAfterPx: 6 }),
    bidiCjkTable(),

    para([run("A content control with a right-to-left value: "), sdtRun({ type: "richText", alias: "RTL rich text" }, "نصٌّ غنيٌّ قابل للتحرير"), run(" — and one with Japanese: "), sdtRun({ type: "richText", alias: "日本語" }, "編集可能なテキスト"), run(".")], { spaceBeforePx: 10 }),

    para([
      run("وحقول ديناميكية داخل فقرة عربية: هذه هي الصفحة رقم "),
      pageField(),
      run("، أُنشئت بتاريخ "),
      dateField("yyyy-MM-dd"),
      run(" — تُعاد حسابتها تلقائيًا عند التحديث."),
    ], { direction: "rtl" }),

    para([run("RTL lists & justification", { bold: true, color: "#1a1a2e" })], { spaceBeforePx: 10, spaceAfterPx: 4 }),
    para([run("قائمة نقطية بالعربية، والعلامة تتدلّى على اليمين")], { direction: "rtl", list: { listId: DEFAULT_BULLET_LIST_ID, level: 0 }, spaceAfterPx: 2 }),
    para([run("العنصر الثاني في القائمة")], { direction: "rtl", list: { listId: DEFAULT_BULLET_LIST_ID, level: 0 }, spaceAfterPx: 2 }),
    para([run("مستوى متداخل داخل القائمة")], { direction: "rtl", list: { listId: DEFAULT_BULLET_LIST_ID, level: 1 }, spaceAfterPx: 2 }),
    para([run("وفقرة عربية مضبوطة (justify) تمتد على عدة أسطر: تُوزَّع المسافات بين الكلمات حتى يمتلئ كل سطر من الحافة إلى الحافة، مع إعادة الترتيب البصري للكلمات والأرقام — تمامًا كما يفعل Word مع النص ثنائي الاتجاه. ".repeat(3))], { direction: "rtl", align: "justify", spaceBeforePx: 6 }),

    // --- Mathematics: MathML equations ----------------------------------------
    heading("Mathematics — MathML equations", 1),
    para([
      run("Equations are first-class objects. They are stored as "),
      run("MathML", { bold: true }),
      run(" (the W3C standard), typeset by the very same layout engine that paginates these pages — fractions, radicals, scripts, summation/integral limits and matrices are all measured on the canvas — and they round-trip through "),
      run(".docx", { fontFamily: "Consolas, monospace", fontSizePx: 14 }),
      run(" as "),
      run("OMML", { bold: true }),
      run(", the math format Word itself uses. They typeset with the STIX Two Math font (real math glyphs, growing delimiters and big operators), and you can author them by typing "),
      run("LaTeX", { bold: true }),
      run(" — Insert → Equation, e.g. "),
      run("\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}", { fontFamily: "Consolas, monospace", fontSizePx: 13 }),
      run(". A few display equations, rendered live below:"),
    ], { spaceAfterPx: 10 }),

    para([
      run("Equations also flow "),
      run("inline", { italic: true }),
      run(" inside a sentence and sit on the text baseline — for example the Pythagorean identity "),
      inlineEq("<math><msup><mi>a</mi><mn>2</mn></msup><mo>+</mo><msup><mi>b</mi><mn>2</mn></msup><mo>=</mo><msup><mi>c</mi><mn>2</mn></msup></math>"),
      run(", or a quick fraction "),
      inlineEq("<math><mfrac><mn>1</mn><mn>2</mn></mfrac></math>"),
      run(" — right-click either one to edit it (in LaTeX or MathML)."),
    ], { spaceBeforePx: 6, spaceAfterPx: 8 }),

    para([run("The quadratic formula — nested radicals, a fraction and a ± operator. Its MathML source is shown beneath the rendered result:", { color: "#3c4043" })], { spaceBeforePx: 6, spaceAfterPx: 6 }),
    eq(MATH_QUADRATIC),
    para([mathmlSource(MATH_QUADRATIC)], { spaceBeforePx: 2, spaceAfterPx: 10 }),

    para([run("A convergent series, with limits set above and below the summation sign (Basel problem):", { color: "#3c4043" })], { spaceBeforePx: 6, spaceAfterPx: 6 }),
    eq(MATH_SUM),

    para([run("The Gaussian integral — super/subscript bounds on the integral and a nested exponential:", { color: "#3c4043" })], { spaceBeforePx: 10, spaceAfterPx: 6 }),
    eq(MATH_INTEGRAL),

    para([run("Euler's identity — the most beautiful equation in mathematics:", { color: "#3c4043" })], { spaceBeforePx: 10, spaceAfterPx: 6 }),
    eq(MATH_EULER),

    para([run("A delimited matrix, laid out as a grid (the 2×2 identity):", { color: "#3c4043" })], { spaceBeforePx: 10, spaceAfterPx: 6 }),
    eq(MATH_MATRIX),

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
