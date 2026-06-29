// Layer 1: Document model — single source of truth. Pure data, no DOM/canvas imports.

import type { DocPosition } from "./position";
import type { MathEquation } from "./math";

/** Underline line styles the model can paint + round-trip (OOXML w:u/@w:val).
 *  Word defines more (and "heavy" weights); the importer folds those onto the
 *  nearest of these, so paint/export stay closed over a small set. Absent
 *  `underlineStyle` ⇒ "single" (a plain solid line — the historical behavior). */
export type UnderlineStyle = "single" | "double" | "thick" | "dotted" | "dash" | "dotDash" | "dotDotDash" | "wave";

export interface CharStyle {
  fontFamily: string;
  fontSizePx: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  /** Underline line style (OOXML w:u/@w:val) — only meaningful when `underline`
   *  is true; absent ⇒ "single". A run with `underline:false` ignores it. */
  underlineStyle?: UnderlineStyle | undefined;
  /** Explicit underline color (CSS hex, e.g. "#ff0000"), incl. theme colors
   *  resolved at import. Absent ⇒ the underline paints in the run's text color
   *  (Word's "auto"). Only meaningful when `underline` is true. */
  underlineColor?: string | undefined;
  strikethrough: boolean;
  color: string;
  /** Hidden text (OOXML w:vanish). Preserved through round-trips but NEVER laid
   *  out, painted, or reachable by the caret/selection — and protected from
   *  deletion (it survives Select-All → Delete). Inert metadata, e.g. the
   *  bookmark-anchor paragraphs generated reports hide. */
  hidden?: boolean;
  letterSpacingPx?: number;
  /** Background highlight (Word's text highlight). `| undefined` so patches can remove. */
  highlightColor?: string | undefined;
  /** Sub/superscript: measured at 0.65× size, baseline-shifted at paint time. */
  verticalAlign?: "sub" | "super" | undefined;
  /** Hyperlink target; linked runs paint blue+underlined, Ctrl+click opens. */
  link?: string | undefined;
  /** Footnote reference: this run IS the marker (its text is the note number,
   *  kept in sync by the insert command's renumber pass; typically also
   *  verticalAlign 'super'). Points into Document.footnotes. */
  footnoteRef?: string | undefined;
  /** Content-control membership (OOXML w:sdt) as an ORDERED ANCESTRY PATH,
   *  outermost→innermost. Contiguous runs sharing the same path form one inline
   *  control; runs sharing a path PREFIX are nested inside the same outer
   *  control(s). Properties for every id on the path live in `Document.sdts`.
   *  This is the INLINE chain only — block-level controls put their ids on
   *  `Block.sdtPath`; a run's full enclosing chain is `block.sdtPath ++ sdtPath`.
   *  Mirrors `fieldId`/`Block.fieldId`; an invisible flat marker — no layout
   *  effect. Treat the array as IMMUTABLE (always replace, never mutate in
   *  place — shallow `{...style}` copies alias it). `| undefined` (never `[]`)
   *  so removing the last control strips the marker. */
  sdtPath?: string[] | undefined;
  /** Inline-field membership (OOXML complex field): contiguous runs sharing a
   *  fieldId are one inline field's RESULT (e.g. PAGE, DATE, IF); the definition
   *  lives in `Document.fields`. The marker is what distinguishes a real field
   *  result from literal text (e.g. a `{page}` token a user merely typed).
   *  `| undefined` so removing a field can strip the marker. Mirrors `sdtPath`;
   *  block-level fields use `Block.fieldId` instead. */
  fieldId?: string | undefined;
  /** Character-style reference (OOXML w:rStyle → a type==="character" NamedStyle).
   *  REFERENCE ONLY: the concrete formatting still lives baked on the run (the
   *  model is concrete, not a live cascade), so layout/paint ignore this — it
   *  exists for docx round-trip and for "what style is this run?" UI. `| undefined`
   *  so removing a character style can strip the marker. */
  charStyleId?: string | undefined;
  /** Inline equation payload (OOXML inline `m:oMath`). When present, this run is
   *  a single OBJECT REPLACEMENT CHARACTER (U+FFFC) in `text` that renders as the
   *  laid-out MathML — an inline replaced element. The AST travels on the run
   *  (not a registry) so the layout prepare-cache stays self-contained; `styleEq`
   *  compares it by reference so equation runs never merge with their neighbours.
   *  Block (display) equations use `EquationBlock` instead. `| undefined` so a
   *  patch can strip it. */
  equation?: MathEquation | undefined;
  /** Explicit right-to-left run (OOXML w:rPr/w:rtl). Forces this run's text to a
   *  bidi-RTL embedding regardless of its characters, mirroring Word's per-run
   *  "rtl" toggle. Absent = resolve direction from the characters' Unicode bidi
   *  classes under the paragraph's base direction (the common case — Arabic/Hebrew
   *  text reorders correctly without this flag). `| undefined` so a patch can
   *  remove it. */
  rtl?: boolean | undefined;
}

/** Structured document tag (Word content control) properties — a direct
 *  mirror of w:sdtPr so the importer can map losslessly. */
export type SdtType = "richText" | "plainText" | "checkbox" | "dropDown" | "comboBox" | "date";

export interface SdtProps {
  type: SdtType;
  /** Title shown on the control's tab (w:alias). */
  alias?: string;
  /** Machine-readable tag (w:tag). */
  tag?: string;
  /** Content is currently the gray placeholder — typing replaces it whole. */
  placeholder?: boolean;
  /** Dropdown / combo box choices (w:listItem). */
  listItems?: { display: string; value: string }[];
  /** Date display format (w:date/@w:fullDate format, e.g. "M/d/yyyy"). */
  dateFormat?: string;
  /** Checkbox state (w14:checkbox). */
  checked?: boolean;
  /** w:lock="sdtContentLocked" — contents cannot be edited. */
  lockContent?: boolean;
  /** w:lock="sdtLocked" — the control cannot be deleted. */
  lockControl?: boolean;
}

export interface ParaStyle {
  align: "left" | "center" | "right" | "justify";
  /** Base writing direction (OOXML w:pPr/w:bidi). "rtl" lays the paragraph out
   *  right-to-left: the bidi base level is RTL (so a leading Latin word still sits
   *  on the right), `align: "left"|"right"` are interpreted as START/END (mirrored),
   *  and left/right indents swap to start/end. Absent = "ltr" (the default — every
   *  existing document is unaffected). Resolved through the Stylesheet cascade like
   *  `align`. */
  direction?: "ltr" | "rtl";
  lineHeight: number; // multiplier
  spaceBeforePx: number;
  spaceAfterPx: number;
  indentFirstLinePx: number;
  indentLeftPx: number;
  /** Right-edge indent (docx w:ind/@w:right|@w:end): narrows every line from the
   *  right margin. Absent = 0. Drives the ruler's right-indent marker. */
  indentRightPx?: number;
  /** Never leave this block as the last on a page (headings). */
  keepWithNext?: boolean;
  /** Never split this paragraph across pages/columns (docx w:keepLines) — it
   *  moves whole instead; only a paragraph taller than a page still splits. */
  keepLinesTogether?: boolean;
  /** This paragraph starts a new page (Ctrl+Enter; docx w:pageBreakBefore). */
  pageBreakBefore?: boolean;
  /** List membership (docx w:numPr): definition ref + level 0..8.
   *  Explicitly `| undefined` so a setParaStyle patch can REMOVE it. */
  list?: { listId: string; level: number } | undefined;
  namedStyle?: string; // resolved through Stylesheet cascade, e.g. "heading1"
  /** Effective outline level (docx w:outlineLvl), 0-8 = TOC levels 1-9 — resolved
   *  through the paragraph-style cascade (heading styles carry it). Drives TOC
   *  generation under the field's `\u` switch and robust heading detection when
   *  styles use opaque ids. Absent = body text (no outline level). */
  outlineLevel?: number;
  /** This paragraph ENDS a section (OOXML puts sectPr ON a paragraph). `props`
   *  describes the section being terminated; blocks after it belong to the next
   *  break's section, or to `Document.section` (the final/body sectPr).
   *  `| undefined` so a setParaStyle patch can remove the break. */
  sectionBreak?: { type: "nextPage"; props: SectionPatch } | undefined;
  /** Ctrl+Shift+Enter: this paragraph starts the next newspaper column
   *  (no-op in single-column sections, like Word). */
  columnBreakBefore?: boolean;
  /** Table-of-contents entry pointing at a heading paragraph. The page number
   *  is PAINT-ONLY (engine post-pass) so it can never go stale; Ctrl+click
   *  jumps to the target. `| undefined` so regeneration can clear it. */
  tocEntry?: { targetId: string; level: number } | undefined;
  /** Explicit tab stops (docx w:tabs), sorted by `posPx`. A `\t` in run text
   *  advances to the next stop past the current x; past the last explicit stop
   *  (or with none) the layout falls back to a fixed default interval. */
  tabStops?: TabStop[];
}

export type TabAlign = "left" | "center" | "right" | "decimal";
export type TabLeader = "none" | "dot" | "dash" | "underscore";

export interface TabStop {
  /** Position from the start content edge, after the paragraph's start indent —
   *  the LEFT edge for an LTR paragraph, the RIGHT edge for an RTL (`w:bidi`) one. */
  posPx: number;
  /** Text alignment at the stop (default "left"). */
  align?: TabAlign;
  /** Filler drawn from the previous content to the stop (default "none"). */
  leader?: TabLeader;
}

/** Style-homogeneous span of text. Adjacent equal-styled runs are merged on every edit. */
export interface Run {
  text: string;
  style: CharStyle;
}

export interface Paragraph {
  kind: "paragraph";
  id: string; // stable — layout cache & undo key on this, never on index
  revision: number; // bumped by ops; invalidates the pretext prepare-cache
  runs: Run[];
  style: ParaStyle;
  /** Membership in a generic field's result region (see Document.fields). A
   *  contiguous run of blocks sharing a fieldId IS the field's result; export
   *  wraps them in the field's begin/instrText/separate…end. Absent = ordinary
   *  content. Like `tocEntry`/run `sdtPath`, an invisible flat marker — no layout
   *  effect. `| undefined` so an op can clear it. */
  fieldId?: string | undefined;
  /** Block-level content-control ancestry (OOXML block-level w:sdt), outer→inner.
   *  A contiguous run of blocks sharing this path is wrapped by the control(s);
   *  this represents controls around whole paragraphs/tables (incl. run-less
   *  blocks). Inline controls inside the block put their ids on the runs'
   *  `CharStyle.sdtPath` instead. Mirrors `fieldId`; invisible marker. Treat as
   *  IMMUTABLE. `| undefined` (never `[]`). */
  sdtPath?: string[] | undefined;
}

export interface ImageBlock {
  kind: "image";
  id: string;
  revision: number;
  /** Runtime image URL (a session-local `blob:` or an inline `data:` URL). NOT
   *  portable — a `blob:` URL dies with the tab and means nothing to the server.
   *  Persistence/replication addresses the bytes by `mediaId` instead and
   *  rehydrates `src` on load (see shared/persist/media + serialize). */
  src: string;
  /** Content address of the image bytes (`sha256(bytes)` hex). The portable,
   *  stable handle stored in snapshots and ops; absent only for legacy images
   *  not yet registered in a MediaStore (then `src` may carry an inline data:
   *  URL as a fallback). */
  mediaId?: string;
  widthPx: number;
  heightPx: number;
  align: "left" | "center" | "right";
  /** 'block' (default): occupies vertical space like a paragraph.
   *  'square': floats at the left/right margin (per align) and following text
   *  flows around it — pretext's per-line maxWidth makes this affordable. */
  wrap?: "block" | "square";
  /** Absolutely-positioned anchored image (DOCX wp:anchor + wp:wrapNone) that
   *  sits BEHIND (behind=true) or in front of the text. Unlike `wrap`, it does
   *  NOT occupy vertical flow space and does NOT reflow surrounding text — it is
   *  painted at `offset{X,Y}Px` from the `relFrom{H,V}` origin. Mutually
   *  exclusive with `wrap`; absent = ordinary flow image. */
  anchor?: {
    behind: boolean;
    offsetXPx: number;
    offsetYPx: number;
    relFromH: "page" | "margin" | "column" | "leftMargin" | "rightMargin" | "character";
    relFromV: "page" | "margin" | "paragraph" | "line" | "topMargin" | "bottomMargin";
    decorative?: boolean;
    /** Stacking order among anchored images in the SAME layer (behind/front):
     *  higher paints later (on top). Drives "bring to front"/"send to back";
     *  maps to/from OOXML wp:anchor @relativeHeight. Default 0. */
    z?: number;
  };
  /** Field result membership — see Paragraph.fieldId. */
  fieldId?: string | undefined;
  /** Block-level content-control ancestry — see Paragraph.sdtPath. */
  sdtPath?: string[] | undefined;
}

/** Cells hold Blocks: paragraphs (first-class editing targets, located through
 *  model/text.ts), images, and nested tables (rendered; their inner cells are
 *  read-only — the paragraph locator goes one level deep). */
/** One resolved cell-edge border. The importer collapses the OOXML cascade
 *  (table style → w:tblBorders → w:tcBorders, plus inside/outside selection)
 *  down to a concrete per-edge spec, so paint just draws what it's given. */
export interface CellBorder {
  /** CSS color, e.g. "#000000". */
  color: string;
  /** Line width in px (OOXML w:sz is eighths of a point → px). */
  widthPx: number;
  /** Default "single". "none" is expressed by omitting the edge entirely. */
  style?: "single" | "double" | "dashed" | "dotted";
}

/** Resolved per-edge borders for a cell. An omitted edge draws no line. */
export interface CellBorders {
  top?: CellBorder;
  right?: CellBorder;
  bottom?: CellBorder;
  left?: CellBorder;
}

/** Table-level default borders (OOXML w:tblPr/w:tblBorders). Extends the per-cell
 *  edges with the two INTERIOR edges Word resolves between adjacent cells. The
 *  importer cascades these onto individual cells for layout/paint; the table keeps
 *  them so export re-emits the table-level defaults at tblPr level rather than only
 *  the baked per-cell copies. */
export interface TableBorders extends CellBorders {
  /** Interior horizontal edge between vertically-adjacent cells (w:insideH). */
  insideH?: CellBorder;
  /** Interior vertical edge between horizontally-adjacent cells (w:insideV). */
  insideV?: CellBorder;
}

/** Inner cell padding in px, resolved from the OOXML cell-margin cascade
 *  (w:tcMar over the table's w:tblCellMar over Word's defaults). Word's default
 *  is 0 top/bottom and ~7.2px (108 twips) left/right — vertical padding is NOT
 *  symmetric with horizontal, which is why a fixed all-sides pad makes rows too
 *  tall. Absent = the engine's Word-matching default. */
export interface CellMargin {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export interface TableCell {
  id: string;
  blocks: Block[];
  /** Horizontal merge: this cell covers N columns (default 1). */
  colSpan?: number;
  /** Vertical merge: this cell covers N rows (default 1). The cell lives in its
   *  TOP row; the rows it spans into simply omit a cell for that grid column
   *  (HTML rowspan semantics), so the importer must drop w:vMerge="continue"
   *  cells and bump this on the "restart" cell. */
  rowSpan?: number;
  /** Resolved background fill (CSS color) from w:shd. Absent = no fill. */
  shading?: string;
  /** Resolved per-edge borders. Absent = renderer's default light grid (so
   *  native/unstyled tables keep a visible grid); present = draw exactly these
   *  edges, where an omitted edge means "no border on that side". */
  borders?: CellBorders;
  /** Resolved inner padding (px) from the w:tcMar/w:tblCellMar cascade. Absent =
   *  engine default (Word's 0 vertical, ~7.2px horizontal). */
  margin?: CellMargin;
  /** Preferred cell width (OOXML w:tcW). Only consulted when the table is in an
   *  autofit mode, where it clamps the cell's content-derived min/max width up to
   *  this preference (Word semantics). `abs` = px (from dxa); `pct` = px resolved
   *  per layout from a percentage of the table width. Absent = content-only. */
  preferredWidth?: { px: number; type: "abs" | "pct" };
  /** Vertical alignment of the cell's content within its box (OOXML w:tcPr/w:vAlign).
   *  Absent = "top" (the historical default). For "center"/"bottom" the layout offsets
   *  the block stack by the slack between content height and the cell's height — most
   *  visible in a tall (rowSpan) or fixed-height cell with short content. */
  vAlign?: "top" | "center" | "bottom";
}

export interface TableRow {
  cells: TableCell[];
}

/** Which conditional bands of a table style this table activates (OOXML w:tblLook).
 *  Word's default turns on the header row and row banding. */
export interface TableCondOverrides {
  firstRow?: boolean;
  lastRow?: boolean;
  firstCol?: boolean;
  lastCol?: boolean;
  bandRows?: boolean;
  bandCols?: boolean;
}

export interface TableBlock {
  kind: "table";
  id: string;
  revision: number;
  rows: TableRow[];
  /** Column widths as fractions of the content width (sum = 1). Absent = equal.
   *  In autofit modes this is the last-known snapshot (used as the export grid
   *  hint and the value the "Fixed" path / column drag edits); layout recomputes
   *  the painted widths from cell content and does not depend on it. */
  colFractions?: number[];
  /** Column sizing strategy (OOXML w:tblLayout + w:tblW). Absent = "fixed", the
   *  historical behavior: columns are `colFractions × content width`. The autofit
   *  modes derive column widths from cell content at layout time —
   *  "autofitContents" lets the table shrink below the content width to fit its
   *  content; "autofitWindow" always fills the available width. */
  widthMode?: "fixed" | "autofitContents" | "autofitWindow";
  /** Preferred TOTAL table width (OOXML w:tblPr/w:tblW), honored in "fixed" mode
   *  only. Absent = span the full content width (the historical behavior). `pct` is
   *  0..100 of the content width; `px` is absolute CSS px. Resolved and clamped to
   *  [floor, contentWidth] at layout time; the autofit modes ignore it. */
  preferredWidth?: { type: "pct" | "px"; value: number };
  /** Horizontal alignment of the table within the content width (OOXML w:tblPr/w:jc),
   *  applied whenever the table is narrower than the band. Absent = "left". */
  align?: "left" | "center" | "right";
  /** Table-level default borders (OOXML w:tblPr/w:tblBorders), including interior
   *  edges. The importer cascades these onto each cell so layout/paint stay
   *  per-cell; kept here so export hoists the table-level defaults back to tblPr
   *  level instead of relying only on the baked per-cell copies. Absent = none. */
  defaultBorders?: TableBorders;
  /** Table-level default shading fill (OOXML w:tblPr/w:shd) — a CSS color applied to
   *  every cell unless the cell overrides it. Absent = no table-level fill. */
  defaultShading?: string;
  /** Table-level default cell margins (OOXML w:tblPr/w:tblCellMar), the base each
   *  cell's own w:tcMar overrides per side. Absent = Word's defaults. */
  defaultCellMargin?: CellMargin;
  /** Table-style reference (OOXML w:tblStyle → Document.tableStyles). The effective
   *  per-cell formatting is baked onto the cells; this is kept for re-editing and
   *  round-trip. Absent = no table style (direct cell formatting only). */
  styleId?: string | undefined;
  /** Which conditional bands of the referenced style are active (w:tblLook). */
  condOverrides?: TableCondOverrides;
  /** Field result membership — see Paragraph.fieldId. */
  fieldId?: string | undefined;
  /** Block-level content-control ancestry — see Paragraph.sdtPath. */
  sdtPath?: string[] | undefined;
}

/** A display (block) equation — Word's `m:oMathPara`. Stands on its own line(s)
 *  like an image; the MathML AST lives directly on the block (no registry). */
export interface EquationBlock {
  kind: "equation";
  id: string;
  revision: number;
  /** The equation as a Presentation-MathML AST (the canonical form). */
  equation: MathEquation;
  /** Horizontal placement of the equation. Default "center" (Word display math). */
  align?: "left" | "center" | "right";
  /** Field result membership — see Paragraph.fieldId. */
  fieldId?: string | undefined;
  /** Block-level content-control ancestry — see Paragraph.sdtPath. */
  sdtPath?: string[] | undefined;
}

export type Block = Paragraph | ImageBlock | TableBlock | EquationBlock;

/** OOXML page-number / list format (the field `\* <fmt>` switch). */
export type PageNumFmt = "arabic" | "roman" | "Roman" | "alpha" | "Alpha";
/** IF comparison operators. */
export type IfOp = "=" | "<>" | "<" | ">" | "<=" | ">=";

/** Typed, parsed form of a BUILT-IN field's definition — drives the field
 *  constructor UI and the evaluator. `FieldDef.instruction` stays the verbatim
 *  round-trip source of truth; `spec` is derived from it and re-synthesized back.
 *  TOC is NOT here — it keeps its own `tocEntry`/`tocInstruction` path. */
export type FieldSpec =
  | { type: "PAGE"; numFmt?: PageNumFmt }
  | { type: "NUMPAGES"; numFmt?: PageNumFmt }
  | { type: "DATE"; format: string }
  | { type: "TIME"; format: string }
  | { type: "IF"; operandA: string; op: IfOp; operandB: string; trueRuns: Run[]; falseRuns: Run[] };

/** A generic OOXML field tracked in the model: its verbatim instruction plus a
 *  classification. CUSTOM (host-resolvable) fields and the BUILT-IN fields the
 *  editor understands (PAGE/NUMPAGES/DATE/TIME/IF) are both tracked this way; TOC
 *  keeps its own `tocEntry`/`tocInstruction` path. The field's result is the
 *  contiguous run of BLOCKS carrying its `id` as `Block.fieldId` (region fields)
 *  OR the contiguous RUNS carrying it as `CharStyle.fieldId` (inline fields). */
export interface FieldDef {
  id: string;
  /** Verbatim w:instrText (e.g. ` MYCHART "sales-2026" `, ` PAGE \* roman `), re-emitted on export. */
  instruction: string;
  /** Field keyword, uppercased — the first instruction token (e.g. "PAGE", "MYCHART"). */
  name: string;
  kind: "builtin" | "custom";
  /** Parsed spec for built-in fields the editor understands; absent for opaque
   *  custom (host-resolved) fields and built-ins we only round-trip verbatim. */
  spec?: FieldSpec;
}

/** The six margin-band stories: default header/footer plus the Word variants
 *  ("Different first page" / "Different odd & even"). Container ops, the
 *  paragraph locator, and the band layout all key on these names. */
export type BandContainer =
  | "header"
  | "footer"
  | "headerFirst"
  | "headerEven"
  | "footerFirst"
  | "footerEven";

export const BAND_CONTAINERS: readonly BandContainer[] = [
  "header",
  "footer",
  "headerFirst",
  "headerEven",
  "footerFirst",
  "footerEven",
];

/** One newspaper column's geometry (px). When a section supplies `cols`, it
 *  overrides the equal-width division implied by `columns.count`/`gapPx`. The
 *  array length MUST equal `columns.count`; `spaceAfterPx` is the gap to the
 *  NEXT column (ignored on the last entry). */
export interface ColumnEntry {
  widthPx: number;
  spaceAfterPx: number;
}

/** One edge of a page border (w:pgBorders child). `style` maps to @w:val. */
export interface PageBorderEdge {
  style: "single" | "double" | "dashed" | "dotted" | "thick" | "none";
  /** w:sz is eighth-points (px = sz / 6 at 96 dpi). */
  widthPx: number;
  /** "#rrggbb"; OOXML "auto" maps to "#000000". */
  color: string;
  /** w:space (points) — offset from the page edge / text, converted to px. */
  spacePx?: number;
}

/** w:sectPr/w:pgBorders — page border box. Absent edges are not drawn. */
export interface PageBorders {
  top?: PageBorderEdge;
  right?: PageBorderEdge;
  bottom?: PageBorderEdge;
  left?: PageBorderEdge;
  /** w:pgBorders/@w:offsetFrom — "page" (from page edge) or "text" (from margin). */
  offsetFrom?: "page" | "text";
}

/** Per-section overrides carried on a section-break paragraph. Absent fields
 *  inherit from `Document.section` — Word's "link to previous" for bands, and
 *  shared page geometry unless the user changes it for one section. */
export interface SectionPatch {
  pageWidthPx?: number;
  pageHeightPx?: number;
  marginPx?: { top: number; right: number; bottom: number; left: number };
  /** `null` = explicitly single-column; absent = inherit `Document.section`.
   *  `sep` draws a separator line between columns; `cols` overrides the
   *  equal-width division with explicit per-column widths. */
  columns?: { count: number; gapPx: number; sep?: boolean; cols?: ColumnEntry[] } | null;
  /** Restart page numbering at this section (absent = continue counting). */
  pageNumberStart?: number;
  /** w:background/@w:color — page fill ("#rrggbb"). See `SectionProps`. */
  pageColorHex?: string;
  /** w:sectPr/w:pgBorders — page border box. */
  pageBorders?: PageBorders;
  /** w:pgMar/@w:header — distance (px) from the page TOP to the header's top edge. */
  headerDistancePx?: number;
  /** w:pgMar/@w:footer — distance (px) from the page BOTTOM to the footer's bottom
   *  edge (ECMA-376). The footer grows upward from there. */
  footerDistancePx?: number;
  header?: Block[];
  footer?: Block[];
  headerFirst?: Block[];
  headerEven?: Block[];
  footerFirst?: Block[];
  footerEven?: Block[];
}

export interface SectionProps {
  pageWidthPx: number;
  pageHeightPx: number;
  marginPx: { top: number; right: number; bottom: number; left: number };
  /** Newspaper columns: the content box divides into `count` boxes separated by
   *  `gapPx`; flow fills column 1 top-to-bottom, then column 2, then the next
   *  page. Absent = single column. `sep` draws a line between columns; `cols`
   *  overrides the equal-width division with explicit per-column widths. */
  columns?: { count: number; gapPx: number; sep?: boolean; cols?: ColumnEntry[] };
  /** Restart page numbering at this section's first page ({page} tokens).
   *  Absent = continue counting from the previous page. */
  pageNumberStart?: number;
  /** w:background/@w:color — page fill behind everything ("#rrggbb"). Absent =
   *  white/no fill. NOTE: OOXML w:background is document-global (one element on
   *  w:document); stored per-section for a uniform model/dialog, but export
   *  reads it only from `Document.section`. */
  pageColorHex?: string;
  /** w:sectPr/w:pgBorders — page border box. Absent = no border. */
  pageBorders?: PageBorders;
  /** w:pgMar/@w:header — distance (px) from the page TOP to the header's top edge
   *  (header grows down). Absent = center the band in the top margin. */
  headerDistancePx?: number;
  /** w:pgMar/@w:footer — distance (px) from the page BOTTOM to the footer's bottom
   *  edge (footer grows up). Absent = center the band in the bottom margin. */
  footerDistancePx?: number;
  /** Header/footer are full block stories (paragraphs, images, tables) laid out
   *  by the same engine into the margin bands. {page}/{pages} tokens in run text
   *  are substituted per page ({page:roman|Roman|alpha|Alpha} formats). The
   *  First/Even variants override the default on the section's first page and
   *  on even-numbered pages respectively (Word's band variant model). */
  header?: Block[];
  footer?: Block[];
  headerFirst?: Block[];
  headerEven?: Block[];
  footerFirst?: Block[];
  footerEven?: Block[];
}

export interface Document {
  section: SectionProps;
  blocks: Block[];
  /** Named styles (Word's style gallery) — see model/stylesheet.ts. */
  stylesheet?: import("./stylesheet").Stylesheet;
  /** List definitions keyed by id (docx numId space) — see model/lists.ts. */
  lists?: Record<string, import("./lists").ListDefinition>;
  /** Table styles keyed by id (docx styleId space) — see model/tableStyles.ts. */
  tableStyles?: Record<string, import("./tableStyles").TableStyle>;
  /** Footnote bodies keyed by ref id (docx footnotes.xml space). Each note is
   *  a paragraph story laid out in the page-bottom footnote area; notes render
   *  on whatever page their reference run lands on. */
  footnotes?: Record<string, Paragraph[]>;
  /** Content-control properties keyed by sdt id. Runs carry inline membership via
   *  `CharStyle.sdtPath`; blocks carry block-level membership via `Block.sdtPath`.
   *  Every id appearing on any path has an entry here. */
  sdts?: Record<string, SdtProps>;
  /** Bookmark name → its character RANGE (docx w:bookmarkStart/End). `start`/`end`
   *  are positions (block id + UTF-16 offset) that may span paragraphs; a point
   *  bookmark has start === end. The block may live in the body, a table cell, or
   *  a header/footer band. Targets for in-document anchor links ("#name" — TOC
   *  entries, cross-references) and the Bookmarks panel. */
  bookmarks?: Record<string, BookmarkRange>;
  /** The document's `TOC` field instruction (e.g. ` TOC \o "1-3" \h \z \u `),
   *  captured on import so export re-emits it verbatim (honoring its switches)
   *  instead of a hardcoded default. Absent when the document has no TOC field. */
  tocInstruction?: string;
  /** Block id of the paragraph holding the (empty/placeholder) `TOC` field, captured
   *  on import so a headless render can BUILD the entries at the right spot. Absent
   *  when there's no TOC field, or when the TOC already has entries (then the
   *  existing tocEntry blocks mark the location). */
  tocAnchorBlockId?: string;
  /** Generic custom fields keyed by id (blocks carry the membership via
   *  `fieldId`). Populated on import for non-built-in fields; export re-emits each
   *  region as a real complex field; the editor refreshes it via `resolveField`.
   *  Absent when the document has no custom fields. */
  fields?: Record<string, FieldDef>;
}

export interface BookmarkRange {
  start: DocPosition;
  end: DocPosition;
}
