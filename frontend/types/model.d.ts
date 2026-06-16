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
  /** Content-control membership (properties live in Document.sdts). */
  sdtId?: string | undefined;
}

export interface TabStop {
  /** Position from the left content edge, px. */
  posPx: number;
  align?: "left" | "center" | "right" | "decimal";
  leader?: "none" | "dot" | "dash" | "underscore";
}

export interface ParaStyle {
  align: "left" | "center" | "right" | "justify";
  /** Line height multiplier. */
  lineHeight: number;
  spaceBeforePx: number;
  spaceAfterPx: number;
  indentFirstLinePx: number;
  indentLeftPx: number;
  indentRightPx?: number;
  keepWithNext?: boolean;
  keepLinesTogether?: boolean;
  /** This paragraph starts a new page. */
  pageBreakBefore?: boolean;
  /** List membership: definition id + level 0..8. */
  list?: { listId: string; level: number } | undefined;
  /** Named style reference into Document.stylesheet (e.g. "Heading1"). */
  namedStyle?: string;
  columnBreakBefore?: boolean;
  tabStops?: TabStop[];
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
}

export interface TableRow {
  cells: TableCell[];
}

export interface TableBlock {
  kind: "table";
  id: string;
  revision: number;
  rows: TableRow[];
  /** Column widths as fractions of content width (sum = 1). Absent = equal. */
  colFractions?: number[];
}

export type Block = Paragraph | ImageBlock | TableBlock;

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

export interface NamedStyle {
  /** Shares the docx styleId space ("Normal", "Heading1", …). */
  id: string;
  /** Display name. */
  name: string;
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
  /** Footnote bodies keyed by ref id. */
  footnotes?: Record<string, Paragraph[]>;
  /** Content-control properties keyed by sdtId. */
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
