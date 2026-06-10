// Layer 1: Document model — single source of truth. Pure data, no DOM/canvas imports.

export interface CharStyle {
  fontFamily: string;
  fontSizePx: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
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
  /** Content-control membership (OOXML w:sdt): contiguous runs sharing an
   *  sdtId form one inline control; properties live in `Document.sdts`.
   *  `| undefined` so removing a control can strip the marker. */
  sdtId?: string | undefined;
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
  /** Position from the left content edge (after the paragraph's left indent). */
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
}

export interface TableRow {
  cells: TableCell[];
}

export interface TableBlock {
  kind: "table";
  id: string;
  revision: number;
  rows: TableRow[];
  /** Column widths as fractions of the content width (sum = 1). Absent = equal. */
  colFractions?: number[];
}

export type Block = Paragraph | ImageBlock | TableBlock;

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

/** Per-section overrides carried on a section-break paragraph. Absent fields
 *  inherit from `Document.section` — Word's "link to previous" for bands, and
 *  shared page geometry unless the user changes it for one section. */
export interface SectionPatch {
  pageWidthPx?: number;
  pageHeightPx?: number;
  marginPx?: { top: number; right: number; bottom: number; left: number };
  /** `null` = explicitly single-column; absent = inherit `Document.section`. */
  columns?: { count: number; gapPx: number } | null;
  /** Restart page numbering at this section (absent = continue counting). */
  pageNumberStart?: number;
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
   *  page. Absent = single column. */
  columns?: { count: number; gapPx: number };
  /** Restart page numbering at this section's first page ({page} tokens).
   *  Absent = continue counting from the previous page. */
  pageNumberStart?: number;
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
  /** Footnote bodies keyed by ref id (docx footnotes.xml space). Each note is
   *  a paragraph story laid out in the page-bottom footnote area; notes render
   *  on whatever page their reference run lands on. */
  footnotes?: Record<string, Paragraph[]>;
  /** Content-control properties keyed by sdtId (runs carry the membership). */
  sdts?: Record<string, SdtProps>;
  /** Bookmark name → id of the block it sits in (docx w:bookmarkStart). Targets
   *  for in-document anchor links ("#name" — TOC entries, cross-references). */
  bookmarks?: Record<string, string>;
}
