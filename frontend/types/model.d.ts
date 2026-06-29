// Public document-model types for @forevka/wordcanvas. Hand-written (like
// wordcanvas.d.ts) so the published surface stays self-contained and stable;
// mirrors shared/src/model — the editor, builder, and exporters all consume
// this same plain-data shape.

export interface CharStyle {
  fontFamily: string;
  fontSizePx: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  /** CSS color, e.g. "#202124". */
  color: string;
  /** Hidden text (OOXML w:vanish) — preserved but never laid out or painted. */
  hidden?: boolean;
  letterSpacingPx?: number;
  /** Background highlight (Word's text highlight). */
  highlightColor?: string | undefined;
  /** Sub/superscript. */
  verticalAlign?: "sub" | "super" | undefined;
  /** Hyperlink target; linked runs paint blue+underlined. */
  link?: string | undefined;
  /** Footnote reference id — this run is the marker. */
  footnoteRef?: string | undefined;
  /** Inline content-control ancestry, outer→inner (properties live in
   *  Document.sdts). Runs sharing a path prefix are nested in the same outer
   *  control(s); block-level controls use Block.sdtPath. */
  sdtPath?: string[] | undefined;
  /** Character-style reference (w:rStyle → a type==="character" NamedStyle).
   *  Reference only — concrete formatting stays baked on the run. */
  charStyleId?: string | undefined;
  /** Inline equation payload (OOXML inline m:oMath) — the run is a single U+FFFC
   *  that renders as the laid-out MathML. Block equations use EquationBlock. */
  equation?: MathEquation | undefined;
  /** Explicit right-to-left run (OOXML w:rPr/w:rtl). Forces a bidi-RTL embedding
   *  regardless of the run's characters. Absent = resolve from Unicode bidi classes. */
  rtl?: boolean | undefined;
  /** All-caps display (OOXML w:caps) — letters render uppercased; the model text is
   *  unchanged and the transform is offset-transparent (caret/measurement included). */
  caps?: boolean | undefined;
  /** Small-capitals display (OOXML w:smallCaps) — letters render uppercased, with the
   *  originally-lowercase ones drawn smaller. Takes precedence over `caps`. */
  smallCaps?: boolean | undefined;
}

export interface TabStop {
  /** Position from the left content edge, px. */
  posPx: number;
  align?: "left" | "center" | "right" | "decimal";
  leader?: "none" | "dot" | "dash" | "underscore";
}

export interface ParaStyle {
  align: "left" | "center" | "right" | "justify";
  /** Base writing direction (OOXML w:bidi). "rtl" lays the paragraph out
   *  right-to-left: align "left"/"right" read as START/END (mirrored), and
   *  left/right indents swap to start/end. Absent = "ltr". */
  direction?: "ltr" | "rtl";
  /** Line height multiplier. */
  lineHeight: number;
  spaceBeforePx: number;
  spaceAfterPx: number;
  indentFirstLinePx: number;
  indentLeftPx: number;
  indentRightPx?: number;
  keepWithNext?: boolean;
  keepLinesTogether?: boolean;
  /** Suppress before/after spacing between adjacent same-style paragraphs (docx
   *  w:contextualSpacing) — Word's default for list styles. */
  contextualSpacing?: boolean;
  /** This paragraph starts a new page. */
  pageBreakBefore?: boolean;
  /** List membership: definition id + level 0..8. */
  list?: { listId: string; level: number } | undefined;
  /** Named style reference into Document.stylesheet (e.g. "Heading1"). */
  namedStyle?: string;
  /** Effective outline level (OOXML w:outlineLvl), 0-8 = TOC levels 1-9. Absent =
   *  body text. Makes a paragraph a TOC entry without a heading style. */
  outlineLevel?: number;
  columnBreakBefore?: boolean;
  tabStops?: TabStop[];
  /** Paragraph borders (OOXML w:pBdr) — a box around the paragraph; each edge
   *  reuses the table CellBorder value type. */
  borders?: ParaBorders | undefined;
  /** Paragraph shading fill (OOXML paragraph-level w:shd), a CSS color. */
  shading?: string | undefined;
}

/** Resolved per-edge paragraph borders (OOXML w:pBdr). An omitted edge draws no
 *  line. `between` is the inter-paragraph rule. */
export interface ParaBorders {
  top?: CellBorder;
  right?: CellBorder;
  bottom?: CellBorder;
  left?: CellBorder;
  between?: CellBorder;
}

/** Style-homogeneous span of text. */
export interface Run {
  text: string;
  style: CharStyle;
}

export interface Paragraph {
  kind: "paragraph";
  /** Stable unique id. */
  id: string;
  revision: number;
  runs: Run[];
  style: ParaStyle;
}

export interface ImageBlock {
  kind: "image";
  id: string;
  revision: number;
  /** Image URL — use data: URLs for portable documents. */
  src: string;
  /** Content address of the bytes (sha256 hex), when registered. */
  mediaId?: string;
  widthPx: number;
  heightPx: number;
  align: "left" | "center" | "right";
  /** 'block' (default): own line. 'square': floats per align, text wraps. */
  wrap?: "block" | "square";
}

export interface CellBorder {
  color: string;
  widthPx: number;
  style?: "single" | "double" | "dashed" | "dotted";
}

export interface CellBorders {
  top?: CellBorder;
  right?: CellBorder;
  bottom?: CellBorder;
  left?: CellBorder;
}

export interface CellMargin {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export interface TableCell {
  id: string;
  blocks: Block[];
  colSpan?: number;
  rowSpan?: number;
  /** Background fill (CSS color). */
  shading?: string;
  borders?: CellBorders;
  margin?: CellMargin;
  /** Preferred cell width (w:tcW); clamps content-derived min/max in autofit modes. */
  preferredWidth?: { px: number; type: "abs" | "pct" };
}

export interface TableRow {
  cells: TableCell[];
}

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
  /** Column widths as fractions of content width (sum = 1). Absent = equal. In
   *  autofit modes this is the last-known snapshot / export grid hint. */
  colFractions?: number[];
  /** Column sizing strategy (w:tblLayout + w:tblW). Absent = "fixed". */
  widthMode?: "fixed" | "autofitContents" | "autofitWindow";
  /** Table-style reference (→ Document.tableStyles); effective cell formatting is baked. */
  styleId?: string | undefined;
  /** Which conditional bands of the referenced style are active (w:tblLook). */
  condOverrides?: TableCondOverrides;
}

// ── Math (MathML AST) — mirrors shared/src/model/math.ts ─────────────────────
export type MathVariant =
  | "normal" | "bold" | "italic" | "bold-italic" | "double-struck" | "fraktur"
  | "bold-fraktur" | "script" | "bold-script" | "sans-serif" | "bold-sans-serif"
  | "sans-serif-italic" | "sans-serif-bold-italic" | "monospace";

export interface MathRow { type: "row"; children: MathNode[]; }
export interface MathIdent { type: "ident"; text: string; variant?: MathVariant; }
export interface MathNumber { type: "number"; text: string; }
export interface MathOperator { type: "op"; text: string; stretchy?: boolean; form?: "prefix" | "infix" | "postfix"; }
export interface MathText { type: "text"; text: string; }
export interface MathSpace { type: "space"; widthEm?: number; }
export interface MathFrac { type: "frac"; num: MathNode; den: MathNode; bevelled?: boolean; thickness?: "0" | "normal"; }
export interface MathScript { type: "script"; base: MathNode; sub?: MathNode; sup?: MathNode; }
export interface MathRadical { type: "radical"; radicand: MathNode; index?: MathNode; }
export interface MathFenced { type: "fenced"; open: string; close: string; separators?: string; child: MathNode; }
export interface MathLimit { type: "limit"; base: MathNode; under?: MathNode; over?: MathNode; accent?: boolean; }
export interface MathNary { type: "nary"; op: string; sub?: MathNode; sup?: MathNode; body: MathNode; hideOp?: boolean; }
export interface MathMatrix { type: "matrix"; rows: MathNode[][]; colAlign?: ("left" | "center" | "right")[]; }
export interface MathPhantom { type: "phantom"; child: MathNode; }
export interface MathUnknown { type: "unknown"; omml?: string; mathml?: string; }
export type MathNode =
  | MathRow | MathIdent | MathNumber | MathOperator | MathText | MathSpace
  | MathFrac | MathScript | MathRadical | MathFenced | MathLimit | MathNary
  | MathMatrix | MathPhantom | MathUnknown;
export interface MathEquation { root: MathRow; display: boolean; }

/** A display (block) equation — Word's m:oMathPara. */
export interface EquationBlock {
  kind: "equation";
  id: string;
  revision: number;
  equation: MathEquation;
  align?: "left" | "center" | "right";
  fieldId?: string | undefined;
  sdtPath?: string[] | undefined;
}

export type Block = Paragraph | ImageBlock | TableBlock | EquationBlock;

/** Conditional-format slots of a table style (OOXML w:tblStylePr types). */
export type TableCond =
  | "wholeTable"
  | "firstRow" | "lastRow" | "firstCol" | "lastCol"
  | "band1Horz" | "band2Horz" | "band1Vert" | "band2Vert"
  | "nwCell" | "neCell" | "swCell" | "seCell";

export interface TableCondProps {
  char?: Partial<CharStyle>;
  para?: Partial<ParaStyle>;
  shading?: string;
  borders?: CellBorders;
  margin?: CellMargin;
}

export interface TableStyle {
  id: string;
  name: string;
  basedOn?: string;
  conds: Partial<Record<TableCond, TableCondProps>>;
  rowBandSize?: number;
  colBandSize?: number;
}

export type BandContainer = "header" | "footer" | "headerFirst" | "headerEven" | "footerFirst" | "footerEven";

export interface SectionProps {
  pageWidthPx: number;
  pageHeightPx: number;
  marginPx: { top: number; right: number; bottom: number; left: number };
  /** Newspaper columns. Absent = single column. */
  columns?: { count: number; gapPx: number };
  pageNumberStart?: number;
  headerDistancePx?: number;
  footerDistancePx?: number;
  /** Header/footer block stories. {page}/{pages} tokens in run text are
   *  substituted per page at layout time. */
  header?: Block[];
  footer?: Block[];
  headerFirst?: Block[];
  headerEven?: Block[];
  footerFirst?: Block[];
  footerEven?: Block[];
}

export type NamedStyleType = "paragraph" | "character";

export interface NamedStyle {
  /** Shares the docx styleId space ("Normal", "Heading1", …). */
  id: string;
  /** Display name. */
  name: string;
  /** Paragraph vs character style. Optional; untyped styles read as "paragraph". */
  type?: NamedStyleType;
  basedOn?: string;
  char: Partial<CharStyle>;
  para: Partial<ParaStyle>;
}

export interface Stylesheet {
  styles: NamedStyle[];
  defaultStyleId: string;
}

export type ListNumberFormat = "bullet" | "decimal" | "lowerLetter" | "upperLetter" | "lowerRoman" | "upperRoman";

export interface ListLevel {
  format: ListNumberFormat;
  /** Marker pattern; %N is level N-1's counter (e.g. "%1."). Ignored for bullets. */
  text: string;
  bulletChar?: string;
  indentLeftPx: number;
  hangingPx: number;
  start: number;
  markerStyle?: Partial<CharStyle>;
}

export interface ListDefinition {
  id: string;
  /** Up to 9 levels (0..8). */
  levels: ListLevel[];
}

export interface DocPosition {
  blockId: string;
  offset: number;
}

export interface BookmarkRange {
  start: DocPosition;
  end: DocPosition;
}

export type SdtType = "richText" | "plainText" | "checkbox" | "dropDown" | "comboBox" | "date";

export interface SdtProps {
  type: SdtType;
  alias?: string;
  tag?: string;
  placeholder?: boolean;
  listItems?: { display: string; value: string }[];
  dateFormat?: string;
  checked?: boolean;
  lockContent?: boolean;
  lockControl?: boolean;
}

export interface Document {
  section: SectionProps;
  blocks: Block[];
  stylesheet?: Stylesheet;
  /** List definitions keyed by id. */
  lists?: Record<string, ListDefinition>;
  /** Table styles keyed by id (docx styleId space). */
  tableStyles?: Record<string, TableStyle>;
  /** Footnote bodies keyed by ref id. */
  footnotes?: Record<string, Paragraph[]>;
  /** Content-control properties keyed by sdt id (runs/blocks carry membership via sdtPath). */
  sdts?: Record<string, SdtProps>;
  /** Bookmark name → character range. */
  bookmarks?: Record<string, BookmarkRange>;
  /** The document's `TOC` field instruction (e.g. ` TOC \o "1-3" \h `), captured
   *  on import. Absent when the document has no TOC field. */
  tocInstruction?: string;
  /** Block id of the paragraph holding the (empty/placeholder) `TOC` field, captured
   *  on import so a headless render can build the entries at the right spot. */
  tocAnchorBlockId?: string;
}
