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
  /** Line height multiplier (used when `lineRule` is absent/"auto"). */
  lineHeight: number;
  /** Fixed line-spacing rule (docx w:lineRule). Absent = `lineHeight` is a
   *  multiplier. "exact" = the line is exactly `lineHeightPx` tall (taller content
   *  clips); "atLeast" = at least `lineHeightPx`, growing for a taller line. */
  lineRule?: "exact" | "atLeast";
  /** Fixed line height in px — meaningful only alongside `lineRule`. */
  lineHeightPx?: number;
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

/** A DrawingML preset shape drawn in the flow like an image, with an optional solid
 *  fill + (optionally dashed) solid stroke and rotation. See ShapeBlock in the model. */
export type ShapePreset =
  | "rect"
  | "roundRect"
  | "ellipse"
  | "triangle"
  | "diamond"
  | "rightArrow"
  | "leftArrow"
  | "line";
export type ShapeDash = "solid" | "dash" | "dot" | "dashDot" | "lgDash";
export type ShapeFill = { color: string } | { none: true };
export type ShapeStroke = { color: string; widthPt: number; dash?: ShapeDash } | { none: true };

/** One segment of a freeform custom-geometry path (a:custGeom / a:path). Points are
 *  fractions of the shape box (0–1). */
export type ShapePathSegment =
  | { type: "moveTo"; x: number; y: number }
  | { type: "lineTo"; x: number; y: number }
  | { type: "cubicBezierTo"; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { type: "close" };

/** A freeform custom geometry (a:custGeom). Segments trace a path in normalized box
 *  coordinates (0–1). */
export interface ShapePath {
  segments: ShapePathSegment[];
}

/** A drawing shape's text box body (OOXML wps:txbx) — a nested paragraph flow
 *  rendered read-only inside the shape box. */
export interface ShapeTextBody {
  blocks: Paragraph[];
}

/** A grouped-shapes container (OOXML wpg:wgp): set on `ShapeBlock.group`, it draws
 *  its child drawings composed under the group transform instead of a preset path;
 *  resizing/moving the group transforms the children as one object. */
export interface ShapeGroup {
  children: ShapeGroupChild[];
  /** a:chOff — origin of the child coordinate space (px). */
  childOffsetXPx: number;
  childOffsetYPx: number;
  /** a:chExt — extent of the child coordinate space (px). */
  childExtentXPx: number;
  childExtentYPx: number;
}

/** One member of a {@link ShapeGroup}, at its local rect within the child
 *  coordinate space; `shape` is a leaf shape or a nested group. */
export interface ShapeGroupChild {
  xPx: number;
  yPx: number;
  shape: ShapeBlock;
}

export interface ShapeBlock {
  kind: "shape";
  id: string;
  revision: number;
  /** The vector geometry drawn inside the box (a:prstGeom@prst), or a freeform
   *  path when `custom` is present (a:custGeom); `adjust` carries the raw a:avLst
   *  guide values (name → number) for parametric presets. */
  geometry: { preset: ShapePreset; adjust?: Record<string, number>; custom?: ShapePath };
  /** When present, this shape is a GROUP container (OOXML wpg:wgp) that draws its
   *  children composed under the group transform instead of a preset path. */
  group?: ShapeGroup;
  /** Interior fill; absent = the theme-neutral default. */
  fill?: ShapeFill;
  /** Outline; absent = the default outline. */
  stroke?: ShapeStroke;
  /** Text box body (OOXML wps:txbx) rendered read-only inside the box; absent = none. */
  text?: ShapeTextBody;
  widthPx: number;
  heightPx: number;
  align: "left" | "center" | "right";
  /** Clockwise rotation in degrees (a:xfrm@rot); absent = none. */
  rotation?: number;
  /** 'block' (default): own line. 'square': floats per align, text wraps beside. */
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
  /** Uniform size multiplier (drag-to-resize), default 1. Clamped to [0.25, 4].
   *  Round-trips to .docx as the equation paragraph's run font size. */
  scale?: number | undefined;
  fieldId?: string | undefined;
  sdtPath?: string[] | undefined;
}

/** An embedder-defined custom block, drawn on canvas by a registered block type
 *  (see `registerBlockType`). Atomic + non-editable; holds only JSON-serializable
 *  `data`. Has no native OOXML (export is lossy by default). */
export interface CustomBlock {
  kind: "custom";
  id: string;
  revision: number;
  customType: string;
  data: unknown;
  fieldId?: string | undefined;
  sdtPath?: string[] | undefined;
}

export type Block = Paragraph | ImageBlock | TableBlock | EquationBlock | CustomBlock | ShapeBlock;

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

export interface DocSelection {
  anchor: DocPosition;
  focus: DocPosition;
  /** Preserved X column for Up/Down caret movement. */
  goalX?: number;
}

export interface BookmarkRange {
  start: DocPosition;
  end: DocPosition;
}

/** OOXML page-number / list format (the field `\* <fmt>` switch). */
export type PageNumFmt = "arabic" | "roman" | "Roman" | "alpha" | "Alpha";
/** IF-field comparison operators. */
export type IfOp = "=" | "<>" | "<" | ">" | "<=" | ">=";

/** Typed, parsed form of a BUILT-IN field's definition (drives the field UI + the
 *  evaluator). TOC keeps its own path and is not here. */
export type FieldSpec =
  | { type: "PAGE"; numFmt?: PageNumFmt }
  | { type: "NUMPAGES"; numFmt?: PageNumFmt }
  | { type: "DATE"; format: string }
  | { type: "TIME"; format: string }
  | { type: "IF"; operandA: string; op: IfOp; operandB: string; trueRuns: Run[]; falseRuns: Run[] };

/** A generic OOXML field tracked in the model (custom host-resolved + built-in).
 *  Its result is the blocks/runs carrying its `id` as `fieldId`. */
export interface FieldDef {
  id: string;
  /** Verbatim w:instrText, re-emitted on export. */
  instruction: string;
  /** Field keyword, uppercased (the first instruction token). */
  name: string;
  kind: "builtin" | "custom";
  /** Parsed spec for built-ins the editor understands; absent for opaque fields. */
  spec?: FieldSpec;
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
