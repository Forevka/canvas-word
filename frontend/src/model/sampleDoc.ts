// Flagship showcase document — the editor's initial state when no docId is opened.
// It exercises nearly the full Word feature surface (TOC, fields, content controls,
// complex/merged + cross-page tables, images, lists, footnotes, bookmarks, hidden
// text, headers/footers) so a first-time visitor sees how much canvas-word supports.
// Built as plain model data (the same Document the editor/exporter/collab consume).

import type {
  BookmarkRange, CellBorder, CellMargin, CharStyle, Document, EquationBlock, FieldDef, FieldSpec, ImageBlock, ParaStyle, Paragraph, RowProps, Run, SdtProps, ShapeBlock, ShapeFill, ShapePath, ShapePreset, ShapeStroke, TableBlock, TableCell,
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
const endnotes: Record<string, Paragraph[]> = {};
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
// A tiny raster PNG (a #1a73e8 → #9c27b0 diagonal gradient tile), embedded as a
// data URI. Deliberately a PNG, not an SVG: pdfkit (the PDF exporter) only decodes
// PNG/JPEG, so an SVG data URI throws `image-format-unsupported` and renders as a
// gray placeholder box in the export. Stretched to each image box below. General
// SVG-in-export support is tracked in issue #116.
const TILE_PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAPAAAACgCAIAAAC9uXYyAAAEU0lEQVR42u3cQW7jRgBFwdY/Rq6ZS+cWzEKWLY+MWUyAAHoorwSqabLlgtEQmu/219//nHNu55zrnHPdzsfP7Xq8ONfLkXPOda7zNfhx+uPd8/Tu9XLkPv5p8Ncv/3Hwdb+FPzv9nOv7nd/HX9+OXI9Jfc3l6f5/OP23k71+e62n018ud718zudc17eJf1zuT0//ZfDLZH/8s/5HFf8nqtFMc0bzOWc005zRfDtnNNOc0XzOGc00ZzQ/lhw005zQfLvOaKY5o/m+5KCZ5ojmc67RTHNG8+18Ljlopvn9NT+WHDTTnNB8rjOaac5oflpy0Ezz+2t+/ZaDZprfWPM5ZzTTnNF8O9dopjmj+VxnNNOc0fy5l4Nmmgua799y0ExzRPO574emmeaG5u/bR2mm+c01P+/loJnmt9f860OyNNP81pq/PSRLM83vrvnrIVmaaQ5o/nhIlmaaG5o/90PTTPNpJIdGM80ZzR//oWmmORPQGs00l3Jwo5nmUg5uNNNcysGNZppLObjRTHMpBzeaaS7l4EYzzaUc3GimuZSDG800l3Jwo5nmUg5uNNNcysGNZppLObjRTHMpBzeaaS7l4EYzzaUc3GimuZSDG800l3Jwo5nmUg5uNNNcysGNZppLObjRTHMpBzeaaS7l4EYzzaUc3GimuZSDG800l3Jwo5nmUg5uNNNcysGNZppLObjRTHMpBzeaaS7l4EYzzaUc3GimuZSDG800l3Jwo5nmUg5uNNNcysGNZppLObjRTHMpBzeaaS7l4EYzzaUc3GimuZSDG800l3Jwo5nmUg5uNNNcysGNZppLObjRTHMpBzeaaS7l4EYzzaUc3GimuZSDG800l3Jwo5nmUg5uNNNcysGNZppLObjRTHMpBzeaaS7l4EYzzaUc3GimuZSDG800l3Jwo5nmUg5uNNNcysGNZppLObjRTHMpBzeaaS7l4EYzzaUc3GimuZSDG800l3Jwo5nmUg5uNNNcysGNZppLObjRTHMpBzeaaS7l4EYzzaUc3GimuZSDG800l3Jwo5nmUg5uNNNcysGNZppLObjRTHMpBzeaaS7l4EYzzaUc3GimuZSDG800l3Jwo5nmUg5uNNNcysGNZppLObjRTHMpBzeaaS7l4EYzzaUc3GimuZSDG800l3Jwo5nmUg5uNNNcysGNZppLObjRTHMpBzeaaS7l4EYzzaUc3GimuZSDG800l3Jwo5nmUg5uNNNcysGNZppLObjRTHMpBzeaaS7l4EYzzaUc3GimuZSDG800l3Jwo5nmUg5uNNNcysGNZppLObjRTHMpBzeaaS7l4EYzzaUc3GimuZSDG800l3Jwo5nmUg5uNNNcysGNZppLObjRTHMpBzeaaS7l4EYzzaUc3GimuZSDG800l3Jwo5nmUg5uNNNcysGNZppLObjRTHMpBzeaaS7l4EYzzaUc3GimuZSDG800l3Jwo5nmUg5uNNNcysH9C3YjxNIK1Tm9AAAAAElFTkSuQmCC";
const image = (w: number, h: number, align: ImageBlock["align"], wrap?: "block" | "square", rotation?: number): ImageBlock => ({
  kind: "image", id: id(), revision: 0, src: TILE_PNG, widthPx: w, heightPx: h, align, ...(wrap ? { wrap } : {}), ...(rotation ? { rotation } : {}) });
/** A cropped image (OOXML a:srcRect): the same tile, with crop insets so only the
 *  inner window shows — demonstrating the #63 image-crop round-trip. */
const croppedImage = (w: number, h: number, crop: NonNullable<ImageBlock["crop"]>): ImageBlock => ({
  kind: "image", id: id(), revision: 0, src: TILE_PNG, widthPx: w, heightPx: h, align: "left", crop });

// --- drawing shapes (DrawingML wps:wsp) ----------------------------------------
/** A preset drawing shape (OOXML wps:wsp / a:prstGeom) with an optional fill,
 *  outline (with dash), rotation, and a read-only text box body (wps:txbx) — the
 *  #206 shape round-trip. PR 1 rect/ellipse/line; PR 2 breadth+dash+rotation;
 *  PR 3 the text body. */
const shape = (
  preset: ShapePreset, w: number, h: number, align: ShapeBlock["align"],
  fill?: ShapeFill, stroke?: ShapeStroke, rotation?: number, adjust?: Record<string, number>,
  text?: string[],
  extra?: { wrap?: ShapeBlock["wrap"]; anchor?: ShapeBlock["anchor"]; custom?: ShapePath },
): ShapeBlock => ({
  kind: "shape", id: id(), revision: 0,
  geometry: extra?.custom ? { preset, custom: extra.custom } : adjust ? { preset, adjust } : { preset },
  widthPx: w, heightPx: h, align,
  ...(fill ? { fill } : {}), ...(stroke ? { stroke } : {}), ...(rotation ? { rotation } : {}),
  ...(text ? { text: { blocks: text.map((line) => para([run(line)])) } } : {}),
  ...(extra?.wrap ? { wrap: extra.wrap } : {}), ...(extra?.anchor ? { anchor: extra.anchor } : {}),
});

/** A five-pointed star traced as a freeform path (OOXML a:custGeom) — outer/inner
 *  vertices in normalized box coords (0–1), demonstrating the #220 custom-geometry
 *  round-trip. */
const STAR_PATH: ShapePath = {
  segments: [
    { type: "moveTo", x: 0.5, y: 0.0 },
    { type: "lineTo", x: 0.6176, y: 0.3382 },
    { type: "lineTo", x: 0.9755, y: 0.3455 },
    { type: "lineTo", x: 0.6902, y: 0.5618 },
    { type: "lineTo", x: 0.7939, y: 0.9045 },
    { type: "lineTo", x: 0.5, y: 0.7 },
    { type: "lineTo", x: 0.2061, y: 0.9045 },
    { type: "lineTo", x: 0.3098, y: 0.5618 },
    { type: "lineTo", x: 0.0245, y: 0.3455 },
    { type: "lineTo", x: 0.3824, y: 0.3382 },
    { type: "close" },
  ],
};

/** A grouped-shapes container (OOXML wpg:wgp) — member shapes positioned in the
 *  group box's own pixel space (childOffset 0, childExtent = w×h), so they move and
 *  scale as one selectable object. #221 grouped-shapes round-trip. */
type GroupChild = { preset: ShapePreset; xPx: number; yPx: number; w: number; h: number; fill?: ShapeFill; stroke?: ShapeStroke; rotation?: number; text?: string[] };
const shapeGroup = (
  w: number, h: number, align: ShapeBlock["align"], kids: GroupChild[],
): ShapeBlock => ({
  kind: "shape", id: id(), revision: 0, geometry: { preset: "rect" },
  widthPx: w, heightPx: h, align,
  group: {
    childOffsetXPx: 0, childOffsetYPx: 0, childExtentXPx: w, childExtentYPx: h,
    children: kids.map((c) => ({
      xPx: c.xPx, yPx: c.yPx,
      shape: {
        kind: "shape" as const, id: id(), revision: 0,
        geometry: { preset: c.preset }, widthPx: c.w, heightPx: c.h, align: "left" as const,
        ...(c.fill ? { fill: c.fill } : {}), ...(c.stroke ? { stroke: c.stroke } : {}),
        ...(c.rotation ? { rotation: c.rotation } : {}),
        ...(c.text ? { text: { blocks: c.text.map((line) => para([run(line)])) } } : {}),
      },
    })),
  },
});

// --- symbols (w:sym) -----------------------------------------------------------
/** A symbol-font glyph run (OOXML w:sym): font + hex code point, decoded to the
 *  glyph and painted in that font (e.g. Wingdings). Round-trips as a real w:sym. */
const symRun = (font: string, charHex: string): Run => {
  const cp = parseInt(charHex, 16);
  return { text: Number.isFinite(cp) && cp > 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : "�", style: { ...BODY, fontFamily: `${font}, sans-serif`, fontSizePx: 18, symbol: { font, char: charHex.toUpperCase() } } };
};

// --- equations (MathML) --------------------------------------------------------
/** A display equation block from a MathML string. Stored as the MathML AST (the
 *  canonical form); typeset by the layout engine and round-tripped to .docx as
 *  OMML. `display: true` centers it on its own line like Word's block math. */
const eq = (mathml: string, align: EquationBlock["align"] = "center", scale?: number): EquationBlock => ({
  kind: "equation", id: id(), revision: 0, equation: { ...parseMathml(mathml), display: true }, align,
  ...(scale !== undefined ? { scale } : {}),
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

/** Minor & advanced table props (issue #61): the table is indented from the left
 *  margin (w:tblInd) and lays its columns out in RIGHT-TO-LEFT visual order
 *  (w:bidiVisual) — so the "Rank 1" column paints at the right edge — while carrying
 *  accessibility caption/description metadata. The first cell rotates its text
 *  (w:textDirection) and the rank cells keep their content on one line (w:noWrap). */
const advancedPropsTable = (): TableBlock => ({
  kind: "table", id: id(), revision: 0,
  colFractions: [0.4, 0.2, 0.2, 0.2],
  indentPx: 36,
  bidiVisual: true,
  overlap: "never",
  caption: "Leaderboard",
  description: "A right-to-left, indented table demonstrating issue #61 table properties.",
  rows: [
    { cells: [
      cell("Standings", { bold: true, color: "#fff" }, { shading: "#9c27b0", textDirection: "tbRl" }),
      cell("Rank 1", { bold: true, color: "#fff" }, { shading: "#9c27b0", noWrap: true }),
      cell("Rank 2", { bold: true, color: "#fff" }, { shading: "#9c27b0", noWrap: true }),
      cell("Rank 3", { bold: true, color: "#fff" }, { shading: "#9c27b0", noWrap: true }),
    ] },
    { cells: [
      cell("Points", { bold: true }),
      cell("980", { color: "#188038" }, { hideMark: true }),
      cell("845", { color: "#188038" }, { hideMark: true }),
      cell("712", { color: "#188038" }, { hideMark: true }),
    ] },
  ],
});

/** Vertical cell text (issue #100, w:textDirection): the first column's headers are
 *  rotated 90° — `tbRl` (top→bottom, 90° clockwise) and `btLr` (bottom→top, 90°
 *  counter-clockwise). In AutoFit-to-Contents the rotated column auto-sizes NARROW
 *  (a stack of line heights) while each row grows TALL (the text length), exactly
 *  like a Word vertical table header. */
const verticalTextTable = (): TableBlock => ({
  kind: "table", id: id(), revision: 0, widthMode: "autofitContents",
  rows: [
    { cells: [
      cell("Top→bottom (tbRl)", { bold: true, color: "#fff" }, { shading: "#00897b", textDirection: "tbRl" }),
      cell("North"), cell("South"), cell("East"),
    ] },
    { cells: [
      cell("Bottom→top (btLr)", { bold: true, color: "#fff" }, { shading: "#00897b", textDirection: "btLr" }),
      cell("120"), cell("96"), cell("141"),
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
      // A cell that explicitly CLEARS the table default shading (issue #150) —
      // Word's "No Color". It stays unshaded (white) among its shaded siblings and
      // re-emits <w:shd w:val="clear" w:fill="auto"/> so the clear round-trips.
      {
        cells: [
          c('Explicit clear ("No Color")'),
          cell("w:shd val=clear fill=auto", {}, { shadingCleared: true, margin: { ...pad }, borders: { top: rule, right: rule, bottom: rule, left: rule } }),
        ],
      },
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
  // Endnote body referenced by a marker run below — collected at the document end.
  endnotes["en1"] = [para([run("Endnotes collect at the very end of the document, under their own separator rule — Word's “end of document” placement.", { fontSizePx: 12 })], { spaceAfterPx: 0 })];

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

  const bodyBlocks: (Paragraph | TableBlock | ImageBlock | EquationBlock | ShapeBlock)[] = [
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
    // Minor run typography & effects (w:rPr extras): dstrike, position, w:w scaling,
    // kern, emphasis, outline/shadow/emboss/imprint, run border (w:bdr), fitText.
    para([
      run("Minor run effects: "),
      run("double strike", { doubleStrikethrough: true }), run(", "),
      run("raised", { positionPx: 4 }), run(" and "), run("lowered", { positionPx: -4 }), run(" text, "),
      run("w i d e", { widthScalePct: 180 }), run(" and "), run("narrow", { widthScalePct: 66 }), run(" scaling, "),
      run("kerned", { kerningMinPx: 12 }), run(", "),
      run("emphasis", { emphasisMark: "dot" }), run(", "),
      run("outline", { outline: true }), run(", "), run("shadow", { shadow: true }), run(", "),
      run("emboss", { emboss: true }), run(", "), run("imprint", { imprint: true }), run(", a "),
      run("bordered", { runBorder: { color: "#1a73e8", widthPx: 1 } }), run(" run, a "),
      run("grid-snap-off", { snapToGrid: false }), run(" run (w:snapToGrid), a "),
      run("French-tagged", { lang: { val: "fr-FR" } }), run(" run (w:lang), and "),
      run("fitText", { fitTextPx: 60 }),
      run(" — all round-trip through Word's w:rPr."),
    ]),
    // Case transforms (w:caps / w:smallCaps): the model text stays as typed; only the
    // rendered glyphs are uppercased (small caps additionally shrinks the lowercase ones).
    para([
      run("Case transforms: "),
      run("all caps", { caps: true }), run(" (w:caps) and "),
      run("Small Caps", { smallCaps: true }), run(" (w:smallCaps) — both render UPPERCASED, "),
      run("while the underlying text stays exactly as authored."),
    ], { spaceBeforePx: 4 }),

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
    para([run("Minor & advanced table props (w:tblInd indent, w:bidiVisual right-to-left columns, w:textDirection / w:noWrap cells, plus caption/description alt text):")], { spaceBeforePx: 10, spaceAfterPx: 6 }),
    advancedPropsTable(),
    para([run("Vertical cell text (w:textDirection) — the first column's headers are rotated 90°: tbRl (top→bottom, clockwise) and btLr (bottom→top, counter-clockwise). AutoFit keeps the rotated column narrow while its row grows tall:")], { spaceBeforePx: 10, spaceAfterPx: 6 }),
    verticalTextTable(),
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
    image(360, 110, "center", "block"),
    para([
      run("A square-wrapped image floats and text flows around it. " + LOREM.repeat(3)),
    ]),
    image(150, 110, "left", "square"),
    para([run("A rotated image (18°) — the rotate handle spins it and a:xfrm@rot round-trips:")], { spaceBeforePx: 6, spaceAfterPx: 4 }),
    image(150, 110, "center", "block", 18),

    // --- Drawing shapes (DrawingML wps:wsp) ------------------------------------
    heading("Drawing shapes", 2),
    para([run("Vector preset shapes (OOXML wps:wsp / a:prstGeom) draw in the flow like an image — selectable, drag-resizable, and round-tripping losslessly to .docx. Fill and outline (colour, width, dash) are editable from the shape's floating toolbar and the ribbon Shape group.")], { spaceAfterPx: 6 }),
    para([run("The preset gallery — rectangle, rounded rectangle, ellipse, triangle, diamond, right/left arrow, each with its own fill + outline:")], { spaceAfterPx: 4 }),
    shape("rect", 150, 80, "left", { color: "#bcd6ef" }, { color: "#41719c", widthPt: 1.5 }),
    shape("roundRect", 150, 80, "left", { color: "#d9ead3" }, { color: "#38761d", widthPt: 1.5 }, 0, { adj: 18000 }),
    shape("ellipse", 150, 90, "left", { color: "#f4cccc" }, { color: "#cc4125", widthPt: 1 }),
    shape("triangle", 130, 90, "center", { color: "#fce5cd" }, { color: "#e69138", widthPt: 1.5 }),
    shape("diamond", 130, 90, "center", { color: "#d9d2e9" }, { color: "#674ea7", widthPt: 1.5 }),
    shape("rightArrow", 180, 70, "right", { color: "#c9daf8" }, { color: "#3d6ea5", widthPt: 1.5 }),
    shape("leftArrow", 180, 70, "right", { color: "#ead1dc" }, { color: "#a64d79", widthPt: 1.5 }),
    para([run("Outline styles — solid, dash, dot, dash-dot and a thick long-dash:")], { spaceBeforePx: 8, spaceAfterPx: 4 }),
    shape("rect", 150, 56, "left", { color: "#eeeeee" }, { color: "#333333", widthPt: 1.5 }),
    shape("rect", 150, 56, "left", { color: "#eeeeee" }, { color: "#333333", widthPt: 1.5, dash: "dash" }),
    shape("rect", 150, 56, "left", { color: "#eeeeee" }, { color: "#333333", widthPt: 1.5, dash: "dot" }),
    shape("rect", 150, 56, "left", { color: "#eeeeee" }, { color: "#333333", widthPt: 1.5, dash: "dashDot" }),
    shape("rect", 150, 56, "left", { color: "#eeeeee" }, { color: "#333333", widthPt: 3, dash: "lgDash" }),
    para([run("Fill variants — a stroke-only diagonal line (no fill) and an outline-only rectangle (no fill):")], { spaceBeforePx: 8, spaceAfterPx: 4 }),
    shape("line", 200, 60, "left", { none: true }, { color: "#38761d", widthPt: 2 }),
    shape("rect", 160, 70, "left", { none: true }, { color: "#674ea7", widthPt: 1 }),
    para([run("A rotated rectangle (20°) and a rotated arrow — a:xfrm@rot round-trips:")], { spaceBeforePx: 8, spaceAfterPx: 4 }),
    shape("rect", 150, 90, "center", { color: "#fff2cc" }, { color: "#bf9000", widthPt: 1.5 }, 20),
    shape("rightArrow", 170, 80, "center", { color: "#d0e0e3" }, { color: "#134f5c", widthPt: 1.5 }, -15),
    para([run("A text box — a shape carrying a body of text (OOXML wps:txbx). Double-click (or select it and press Enter) to type inside; Escape pops back out. Edits round-trip losslessly:")], { spaceBeforePx: 8, spaceAfterPx: 4 }),
    shape("rect", 300, 96, "center", { color: "#fff2cc" }, { color: "#bf9000", widthPt: 1 }, undefined, undefined, [
      "Editable text box",
      "Double-click to edit this text, then click away to commit.",
    ]),
    para([run("A freeform custom geometry (OOXML a:custGeom) — a five-pointed star traced as a path of line segments rather than a preset. The path is stored in normalized box coordinates and round-trips losslessly to .docx:")], { spaceBeforePx: 8, spaceAfterPx: 4 }),
    shape("rect", 120, 120, "center", { color: "#ffe599" }, { color: "#bf9000", widthPt: 1.5 }, undefined, undefined, undefined, { custom: STAR_PATH }),

    // --- Shape positioning: wrap / float anchor / z-order (issue #217) ----------
    heading("Shape positioning — wrap, float & z-order", 3),
    para([run("A square-wrapped shape floats at the margin and the paragraph text flows around it, exactly like a square-wrapped image:")], { spaceBeforePx: 6, spaceAfterPx: 4 }),
    shape("ellipse", 130, 100, "left", { color: "#d9ead3" }, { color: "#38761d", widthPt: 1.25 }, undefined, undefined, undefined, { wrap: "square" }),
    para([run("Square wrap lifts the shape out of the block flow and registers a float so following text re-breaks beside it. " + LOREM.repeat(4))]),
    para([run("Below, two absolutely-anchored shapes overlap: the second has a higher z-order so it paints on top. A third shape sits BEHIND this text (a soft background tint) while text stays selectable, and a fourth sits IN FRONT of it.")], { spaceBeforePx: 8, spaceAfterPx: 4 }),
    // Overlapping anchored shapes — same layer, different z (higher paints on top).
    shape("rect", 150, 90, "left", { color: "#c9daf8" }, { color: "#3d85c6", widthPt: 1 }, undefined, undefined, undefined, {
      anchor: { behind: false, offsetXPx: 20, offsetYPx: 0, relFromH: "margin", relFromV: "paragraph", z: 1 },
    }),
    shape("ellipse", 150, 90, "left", { color: "#fce5cd" }, { color: "#e69138", widthPt: 1 }, undefined, undefined, undefined, {
      anchor: { behind: false, offsetXPx: 90, offsetYPx: 30, relFromH: "margin", relFromV: "paragraph", z: 2 },
    }),
    // Behind-text tint (selectable text stays on top) + an in-front-of-text accent.
    shape("rect", 260, 70, "left", { color: "#fff2cc" }, { none: true }, undefined, undefined, undefined, {
      anchor: { behind: true, offsetXPx: 0, offsetYPx: 120, relFromH: "margin", relFromV: "paragraph", z: -1 },
    }),
    shape("ellipse", 80, 80, "left", { color: "#ead1dc" }, { color: "#a64d79", widthPt: 1 }, undefined, undefined, undefined, {
      anchor: { behind: false, offsetXPx: 230, offsetYPx: 110, relFromH: "margin", relFromV: "paragraph", z: 3 },
    }),
    para([run("Anchored shapes do not occupy vertical flow space, so this line follows the previous paragraph directly while the shapes float over/under it. " + LOREM.repeat(2))], { spaceBeforePx: 4 }),

    para([run("A grouped drawing (OOXML wpg:wgp) bundles several shapes into one object — selecting, moving or resizing the group transforms every member together. This group is a small flow diagram: two labelled boxes joined by an arrow.")], { spaceBeforePx: 8, spaceAfterPx: 4 }),
    shapeGroup(360, 96, "center", [
      { preset: "roundRect", xPx: 0, yPx: 20, w: 130, h: 56, fill: { color: "#d9ead3" }, stroke: { color: "#38761d", widthPt: 1.5 }, text: ["Input"] },
      { preset: "rightArrow", xPx: 140, yPx: 36, w: 80, h: 24, fill: { color: "#fff2cc" }, stroke: { color: "#bf9000", widthPt: 1 } },
      { preset: "roundRect", xPx: 230, yPx: 20, w: 130, h: 56, fill: { color: "#c9daf8" }, stroke: { color: "#3d85c6", widthPt: 1.5 }, text: ["Output"] },
    ]),

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
    // Explicit "No Color" over a shaded named style (issue #147). The "Callout" style
    // carries a green fill; the first paragraph inherits it, the second CLEARS it with
    // Word's "Shading → No Color" (w:shd w:val="clear" w:fill="auto") so the fill can't
    // silently return on the next save.
    para([run("This paragraph uses the Callout style and inherits its green shading fill.")], {
      spaceBeforePx: 6, namedStyle: "Callout", shading: "#d9ead3",
    }),
    para([run("This paragraph also uses Callout but explicitly clears the shading (No Color) — the style's fill is overridden and stays cleared through a .docx round-trip.")], {
      spaceBeforePx: 6, namedStyle: "Callout", shadingCleared: true,
    }),
    // --- Contextual spacing: same-style runs sit tight (w:contextualSpacing) ---
    para([run("Contextual spacing — each verse line below carries 12px after-spacing, yet w:contextualSpacing collapses the gaps between adjacent same-style paragraphs (Word's list-style default); only the run's outer edges keep their space:")], { spaceBeforePx: 10, spaceAfterPx: 4 }),
    para([run("Roses are red,")], { contextualSpacing: true, spaceAfterPx: 12 }),
    para([run("violets are blue,")], { contextualSpacing: true, spaceAfterPx: 12 }),
    para([run("contextual spacing keeps these lines tight,")], { contextualSpacing: true, spaceAfterPx: 12 }),
    para([run("the way Word's list paragraphs do.")], { contextualSpacing: true, spaceAfterPx: 12 }),
    // Automatic paragraph spacing (issue #160): Word's w:beforeAutospacing/afterAutospacing
    // (the HTML-<p> behavior). The spacing is baked to an approximate auto value and the
    // attributes round-trip, so a web-sourced doc's paragraphs don't collapse together.
    para([run("This paragraph uses Word's automatic before/after spacing (w:beforeAutospacing / w:afterAutospacing) rather than an explicit value — common in documents converted from HTML.")], {
      spaceBeforeAuto: true, spaceAfterAuto: true, spaceBeforePx: 14, spaceAfterPx: 14,
    }),

    // --- Minor paragraph properties (issue #62) -------------------------------
    heading("Minor paragraph properties", 1),
    para([run("Lower-frequency w:pPr settings round-trip too. This paragraph turns OFF widow/orphan control (w:widowControl), is excluded from line numbering (w:suppressLineNumbers), and carries mirrored indents (w:mirrorIndents) plus right-indent adjustment (w:adjustRightInd) — each preserved through a .docx save and reopen. It also carries the unmodeled East-Asian / hyphenation toggles (w:snapToGrid, w:suppressAutoHyphens, w:kinsoku, w:overflowPunct, w:wordWrap, w:topLinePunct, w:autoSpaceDE, w:autoSpaceDN), which survive the round-trip without any layout behavior.")], {
      spaceBeforePx: 6,
      widowControl: false,
      suppressLineNumbers: true,
      mirrorIndents: true,
      adjustRightInd: true,
      snapToGrid: false,
      suppressAutoHyphens: true,
      kinsoku: false,
      overflowPunct: true,
      wordWrap: false,
      topLinePunct: true,
      autoSpaceDE: false,
      autoSpaceDN: false,
    }),
    para([run("And with extra line spacing, "), run("bottom", { bold: true }), run(" vertical line alignment (w:textAlignment) drops the text onto the lower edge of each tall line box — set it to top, center or baseline to move where the glyphs ride.")], {
      spaceBeforePx: 6,
      lineHeight: 2,
      textAlignment: "bottom",
    }),

    // --- Miscellaneous OOXML round-trip (#63): symbols, image crop, tab stop ----
    heading("Miscellaneous OOXML — symbols, image crop & tab stops", 1),
    para([
      run("Symbol-font glyphs (Word's "),
      run("w:sym", { fontFamily: "Consolas, monospace", fontSizePx: 13 }),
      run(") carry their font and code point, so they survive a round-trip instead of becoming stray characters: "),
      symRun("Wingdings", "F04A"), run("  "), symRun("Wingdings", "F0FC"), run("  "), symRun("Wingdings", "F0E0"),
      run("  "), symRun("Webdings", "F069"),
      run(" — each is a glyph from a symbol font, not text."),
    ], { spaceAfterPx: 8 }),
    para([run("Image cropping (a:srcRect) trims the source to a window — here the same tile is shown whole, then center-cropped:")], { spaceAfterPx: 6 }),
    image(150, 110, "left", "block"),
    croppedImage(150, 110, { left: 0.25, top: 0.2, right: 0.25, bottom: 0.2 }),
    para([
      run("Tab stops honor the document's default interval (w:defaultTabStop): columns\tline up\tat\teach default tab."),
    ], { spaceBeforePx: 8 }),
    // --- Run-level typography: tracking, theme tints & script fonts ------------
    heading("Run-level typography — tracking, theme tints & script fonts", 1),
    para([run("These run-level details survive a full .docx round-trip (import → edit → export), matching what Word stores on every run.")], { spaceAfterPx: 8 }),

    para([run("Character tracking", { bold: true, color: "#1a1a2e" })], { spaceBeforePx: 6, spaceAfterPx: 2 }),
    para([
      run("Tracking (OOXML w:spacing) widens or tightens inter-letter spacing: "),
      run("w i d e   t r a c k i n g", { letterSpacingPx: 3 }),
      run(", normal tracking, and "),
      run("tightened", { letterSpacingPx: -0.5 }),
      run(" — each value re-imports exactly."),
    ]),

    para([run("Theme tint & shade", { bold: true, color: "#1a1a2e" })], { spaceBeforePx: 8, spaceAfterPx: 2 }),
    para([
      run("On import, a theme color with a tint or shade (w:themeTint / w:themeShade) resolves to its actual lighter/darker shade rather than flattening to the flat base hue. These are the resolved shades of the Office accent blue "),
      run("#4472C4", { color: "#4472c4", bold: true }),
      run(" — a lighter "),
      run("60% tint", { color: "#8faadc", bold: true }),
      run(" and a darker "),
      run("50% shade", { color: "#223962", bold: true }),
      run(". (The model stores concrete colors, so the import resolution itself is covered by the round-trip tests.)"),
    ]),

    para([run("Complex-script & East-Asian font slots", { bold: true, color: "#1a1a2e" })], { spaceBeforePx: 8, spaceAfterPx: 2 }),
    para([
      run("Runs preserve their complex-script (w:cs) and East-Asian (w:eastAsia) typefaces: the Arabic phrase "),
      run("مرحبا", { fontFamilyComplexScript: "Scheherazade New, serif" }),
      run(" carries a complex-script face and the Japanese phrase "),
      run("日本語", { fontFamilyEastAsia: "Yu Mincho, serif" }),
      run(" an East-Asian one — both round-trip independently of the Latin font."),
    ]),

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

    para([run("Euler's identity — the most beautiful equation in mathematics, shown enlarged (select it and drag a handle to resize like an image):", { color: "#3c4043" })], { spaceBeforePx: 10, spaceAfterPx: 6 }),
    eq(MATH_EULER, "center", 1.5),

    para([run("A delimited matrix, laid out as a grid (the 2×2 identity):", { color: "#3c4043" })], { spaceBeforePx: 10, spaceAfterPx: 6 }),
    eq(MATH_MATRIX),

    // Endnotes (w:endnoteReference) — the marker run points into Document.endnotes;
    // the body lays out at the very end of the document under a separator rule.
    para([
      run("Endnotes", { bold: true }), run(" round-trip too"),
      run("1", { endnoteRef: "en1", verticalAlign: "super", fontSizePx: 11 }),
      run(" — like footnotes, but parked at the document end."),
    ], { spaceBeforePx: 6 }),
    // --- Fixed line spacing (w:lineRule exact/atLeast) -------------------------
    heading("Fixed line spacing", 1),
    para([
      run("Line spacing is usually a "),
      run("multiplier", { bold: true }),
      run(" of the font size (single, 1.5×, double). Word also supports "),
      run("fixed", { bold: true }),
      run(" point spacing via "),
      run("w:lineRule", { fontFamily: "Consolas, monospace", fontSizePx: 14 }),
      run(": "),
      run("exact", { fontFamily: "Consolas, monospace", fontSizePx: 14 }),
      run(" pins each line to a height in points (taller glyphs are clipped), while "),
      run("atLeast", { fontFamily: "Consolas, monospace", fontSizePx: 14 }),
      run(" sets a floor that grows for a taller line. Both round-trip through .docx and drive pagination."),
    ], { spaceAfterPx: 8 }),
    para([run("This paragraph uses EXACT 28px line spacing — every line box is exactly 28px tall regardless of the text. Word stores it as w:spacing w:line=\"420\" w:lineRule=\"exact\". Notice the lines sit tight at a constant pitch even though this run is long enough to wrap across several lines on the page.".repeat(1))],
      { lineRule: "exact", lineHeightPx: 28, spaceAfterPx: 8 }),
    para([run("This paragraph uses AT-LEAST 24px line spacing — lines are at least 24px tall but a line with a larger glyph, ", { }), run("like this 30px word", { fontSizePx: 30 }), run(", grows to fit it. Word stores it as w:spacing w:line=\"360\" w:lineRule=\"atLeast\".")],
      { lineRule: "atLeast", lineHeightPx: 24, spaceAfterPx: 10 }),
    // --- Section breaks & line numbering (w:sectPr/w:type, w:lnNumType) ---------
    heading("Section breaks & line numbering", 1),
    // This paragraph ENDS the main flow as its own section; the next (line-numbered)
    // section follows. A plain Next Page break — geometry inherited from the body.
    para([run("Word sections can force their first page onto an odd or even page number, and print a number beside every line in the margin. The section that follows demonstrates both: it starts on an odd page (a blank filler page is inserted when the running page count is even) and numbers every line via w:lnNumType.")], {
      spaceAfterPx: 6,
      sectionBreak: { type: "nextPage", props: {} },
    }),

    heading("A line-numbered section, starting on an odd page", 2),
    para([run("Every line below carries a small margin line number. Counting restarts on each new page (Word's default), and the numbers round-trip to .docx and PDF. This justified, multi-line paragraph makes the per-line numbering easy to see: " + LOREM.repeat(6))], { align: "justify", spaceAfterPx: 6 }),
    // This paragraph ENDS the line-numbered section: its w:type is oddPage and its
    // sectPr carries the w:lnNumType that numbers every line of this section.
    para([run("Line numbering is a per-section property — the closing line beneath this one belongs to the final (unnumbered) section, so its lines are not numbered.")], {
      sectionBreak: { type: "oddPage", props: { lineNumbering: { countBy: 1, restart: "newPage" } } },
    }),

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

  // Extend the starter gallery with a shaded "Callout" paragraph style so the
  // paragraph-shading section can demonstrate an explicit clear against it (#147).
  const stylesheet = defaultStylesheet();
  stylesheet.styles.push({
    id: "Callout", name: "Callout", type: "paragraph", basedOn: "Normal",
    char: {}, para: { shading: "#d9ead3", spaceBeforePx: 8, spaceAfterPx: 8 },
  });

  const doc: Document = {
    stylesheet,
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
    endnotes,
    bookmarks,
    tocInstruction: ' TOC \\o "1-3" \\h \\z ',
    // Document-level default tab interval (w:defaultTabStop) — 0.75in instead of the
    // engine's 0.5in fallback, so the tab-stop demo above visibly honors it (#63).
    defaultTabStopPx: 72,
  };
  return doc;
}
