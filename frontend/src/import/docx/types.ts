// .docx import — shared types.
//
// The pipeline is: zip → XML parts → IR (intermediate representation) → model.
// The IR is a faithful-but-minimal decode of document.xml: the parser keeps
// everything it understands (even what the model can't hold yet); mapToModel
// decides what survives and emits an ImportWarning for every lossy decision.

import type { Document } from "@cw/shared";

export type ImportPhase = "unzip" | "styles" | "parse" | "map";

export type ImportErrorCode = "NOT_ZIP" | "ENCRYPTED" | "NO_DOCUMENT_PART" | "MALFORMED_XML";

export class ImportError extends Error {
  readonly code: ImportErrorCode;
  constructor(code: ImportErrorCode, message: string) {
    super(message);
    this.name = "ImportError";
    this.code = code;
  }
}

export interface ImportWarning {
  code: string;
  message: string;
}

/** Deduplicating warning collector — a 70-page doc with 400 tabs should say
 *  "tabs converted" once, not 400 times. */
export class WarningSink {
  readonly list: ImportWarning[] = [];
  private readonly seen = new Set<string>();
  add(code: string, message: string): void {
    if (this.seen.has(code)) return;
    this.seen.add(code);
    this.list.push({ code, message });
  }
}

export interface ImportResult {
  doc: Document;
  /** Every lossy mapping decision — surfaced, not swallowed. */
  warnings: ImportWarning[];
  /** blob: URLs created for embedded media; caller revokes when the doc is discarded. */
  mediaUrls: string[];
}

// ---------------------------------------------------------------------------
// Worker protocol

export interface ToWorker {
  id: number;
  buf: ArrayBuffer; // transferred, not cloned
}

export type FromWorker =
  | { id: number; type: "progress"; phase: ImportPhase; pct: number }
  | { id: number; type: "done"; result: ImportResult }
  | { id: number; type: "error"; code: ImportErrorCode; message: string };

// ---------------------------------------------------------------------------
// IR — what documentParser produces. Units stay as OOXML units (twips,
// half-points); conversion to px happens in mapToModel via units.ts.

/** Direct run formatting decoded from w:rPr. Absent field = "inherit" —
 *  resolved through the style cascade in milestone 2, from defaults today. */
export interface IRRunProps {
  /** w:rStyle reference — recorded now, resolved when StyleResolver lands. */
  styleId?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  /** Raw w:color value: hex without '#' ("FF0000") or "auto". */
  color?: string;
  /** w:sz — half-points. */
  sizeHalfPoints?: number;
  /** w:rFonts w:ascii. */
  fontAscii?: string;
  /** w:rFonts w:asciiTheme — theme font slot ("minorHAnsi", "majorHAnsi", …). */
  fontThemeAscii?: string;
  /** w:color w:themeColor — theme color slot ("accent1", "text1", …). */
  colorTheme?: string;
  /** w:highlight w:val — named highlight color ("yellow", "green", …). */
  highlight?: string;
  /** w:vertAlign w:val — "superscript" | "subscript" | "baseline". */
  vertAlign?: string;
  /** w:vanish — hidden text (Word shows it only with ¶ marks on). */
  vanish?: boolean;
  /** Hyperlink membership (set on runs inside a w:hyperlink). Resolved to a URL
   *  in mapToModel: relId via the part's rels (external target), or anchor for
   *  an in-document bookmark (kept as "#name"). */
  linkRelId?: string;
  linkAnchor?: string;
  /** w:footnoteReference w:id — this run is a footnote marker. mapToModel sets
   *  the run text to the sequential note number and CharStyle.footnoteRef. */
  footnoteId?: string;
}

/** w:sdtPr — content-control properties, decoded faithfully (mapToModel turns
 *  these into the model's SdtProps; runs inside the control carry the sdtId). */
export interface IRSdtProps {
  type: "richText" | "plainText" | "checkbox" | "dropDown" | "comboBox" | "date";
  alias?: string;
  tag?: string;
  /** w:showingPlcHdr — the content is currently the gray placeholder. */
  placeholder?: boolean;
  listItems?: { display: string; value: string }[];
  dateFormat?: string;
  checked?: boolean;
  lockContent?: boolean;
  lockControl?: boolean;
}

export type IRInline =
  | { kind: "run"; text: string; props: IRRunProps; sdtId?: string }
  /** w:br / w:cr — soft line break (model has none; mapToModel splits the
   *  paragraph). page=true for w:br w:type="page" (→ pageBreakBefore on the
   *  follower); column=true for w:br w:type="column" (→ columnBreakBefore). */
  | { kind: "break"; page?: boolean; column?: boolean }
  /** w:drawing / w:pict — becomes a block-level ImageBlock (model has no inline images). */
  | {
      kind: "image";
      relId: string;
      widthEmu?: number;
      heightEmu?: number;
      anchored: boolean;
      /** For wp:anchor: square = text wraps around (maps to ImageBlock.wrap);
       *  block = wrap mode the model can't express (none/topAndBottom). */
      anchorWrap?: "square" | "block";
      /** wp:positionH/wp:align when present. */
      anchorAlign?: "left" | "right" | "center";
    };

export interface IRParaProps {
  /** w:pStyle reference — recorded now, resolved when StyleResolver lands. */
  styleId?: string;
  align?: "left" | "center" | "right" | "justify";
  spaceBeforeTwips?: number;
  spaceAfterTwips?: number;
  /** Multiplier — only set when lineRule is "auto" (or absent). */
  lineHeight?: number;
  indentLeftTwips?: number;
  /** w:ind/@w:right|@w:end — right-edge indent. */
  indentRightTwips?: number;
  /** Negative = hanging indent (w:hanging). */
  indentFirstLineTwips?: number;
  /** w:keepNext — maps onto ParaStyle.keepWithNext. */
  keepWithNext?: boolean;
  /** w:keepLines — maps onto ParaStyle.keepLinesTogether. */
  keepLinesTogether?: boolean;
  /** w:pPr/w:tabs — raw tab stops (pos in twips; val/leader raw OOXML names). */
  tabStops?: { posTwips: number; val?: string; leader?: string }[];
  /** w:pageBreakBefore — this paragraph starts a new page. */
  pageBreakBefore?: boolean;
  /** w:numPr — list membership. numId "0" / null = explicitly NOT a list
   *  (overrides an inherited list from the paragraph style). */
  list?: { numId: string; level: number } | null;
  /** w:pPr/w:sectPr — this paragraph ENDS a section. "page" (nextPage/odd/even)
   *  implies the following content starts a new page; "continuous" doesn't. */
  sectionBreak?: "page" | "continuous";
  /** The ending section's page size (twips), when its w:sectPr declares pgSz.
   *  A "page" break only forces a new page when this differs from the document's
   *  page size — geometry-preserving breaks (footer/header switches, which these
   *  generated reports emit by the dozen) flow instead of leaving half-empty
   *  pages, matching Word. */
  sectionPgSize?: { w: number; h: number };
  /** The ending section's margins (twips) — for a SectionPatch when geometry differs. */
  sectionMarginTwips?: { top: number; right: number; bottom: number; left: number };
  /** The ending section's newspaper columns (count > 1). */
  sectionColumns?: { count: number; spaceTwips?: number };
  /** The ending section's w:pgNumType w:start (page-number restart). */
  sectionPageNumberStart?: number;
  /** The ending section's sectPr declares its own header/footer references —
   *  when such a section is flowed (not page-broken) its bands aren't applied. */
  sectionHasBands?: boolean;
  /** w:pPr/w:rPr — the paragraph MARK's run formatting. Word styles empty
   *  paragraphs (and the ¶ itself) with this; we use it for empty-run style. */
  markRunProps?: IRRunProps;
}

export interface IRParagraph {
  kind: "paragraph";
  props: IRParaProps;
  inlines: IRInline[];
  /** w:bookmarkStart names anchored in this paragraph (TOC/cross-ref targets). */
  bookmarks?: string[];
}

/** One raw border edge (w:top/w:left/… inside w:tblBorders or w:tcBorders).
 *  Kept raw (OOXML units) until mapToModel collapses the cascade to px. */
export interface IRRawBorder {
  /** w:val — "single", "double", "dashed", "nil"/"none", … */
  val: string;
  /** w:sz — eighths of a point. */
  sizeEighthPt?: number;
  /** w:color — hex without '#', or "auto". */
  color?: string;
}

export interface IRBorders {
  top?: IRRawBorder;
  left?: IRRawBorder;
  bottom?: IRRawBorder;
  right?: IRRawBorder;
  /** Interior horizontal/vertical edges (w:tblBorders only). */
  insideH?: IRRawBorder;
  insideV?: IRRawBorder;
}

/** Cell margins in twips, per side. Any side may be absent (inherits the next
 *  level of the cascade: cell → table → Word default). w:tcMar / w:tblCellMar. */
export interface IRCellMargin {
  top?: number;
  right?: number;
  bottom?: number;
  left?: number;
}

export interface IRTableCell {
  /** Full block content: paragraphs, images (inside paragraphs), nested tables. */
  blocks: IRBlock[];
  /** w:gridSpan — columns this cell covers (1 = normal). Maps to TableCell.colSpan. */
  gridSpan: number;
  /** w:vMerge continuation — this cell is swallowed by the cell above. */
  vMergeContinue: boolean;
  /** w:tcPr/w:tcBorders. */
  borders?: IRBorders;
  /** A w:tcBorders element is present — even if empty (all edges nil/absent), the
   *  author explicitly chose this cell's borders, so it must NOT fall back to the
   *  renderer's default grid. */
  bordersSpecified?: boolean;
  /** w:tcPr/w:shd → CSS fill. */
  shd?: string;
  /** w:tcPr/w:tcMar — per-side inner padding override (twips). */
  marginTwips?: IRCellMargin;
}

export interface IRTableRow {
  cells: IRTableCell[];
}

export interface IRTable {
  kind: "table";
  rows: IRTableRow[];
  /** w:tblGrid/w:gridCol widths — become TableBlock.colFractions. */
  colWidthsTwips?: number[];
  /** w:tblPr/w:tblStyle — table style id (its borders/shd are the cascade base). */
  styleId?: string;
  /** w:tblPr/w:tblBorders. */
  borders?: IRBorders;
  /** A w:tblBorders element is present — even if empty, the table explicitly
   *  declares its borders (e.g. "no borders"), suppressing the default grid. */
  bordersSpecified?: boolean;
  /** w:tblPr/w:shd → CSS fill applied to every cell unless overridden. */
  shd?: string;
  /** w:tblPr/w:tblCellMar — table-wide cell-margin default (twips), the base each
   *  cell's own w:tcMar overrides per side. */
  cellMarginTwips?: IRCellMargin;
}

export type IRBlock = IRParagraph | IRTable;

// ---------------------------------------------------------------------------
// Numbering (numbering.xml) — raw decode; mapToModel converts to model lists.

export interface IRListLevel {
  /** Raw w:numFmt w:val ("decimal", "bullet", "lowerRoman", "none", …). */
  format: string;
  /** Raw w:lvlText w:val — "%1." pattern, or the bullet glyph for bullets. */
  lvlText: string;
  start: number;
  indentLeftTwips?: number;
  hangingTwips?: number;
  /** w:lvl/w:rPr — marker glyph styling (esp. the bullet font). */
  markerRunProps?: IRRunProps;
}

export interface IRListDefinition {
  /** numId (the id space paragraphs reference via w:numPr/w:numId). */
  id: string;
  levels: IRListLevel[];
}

export interface IRSection {
  pageWidthTwips?: number;
  pageHeightTwips?: number;
  marginTwips?: { top: number; right: number; bottom: number; left: number };
  /** w:pgMar/@w:header and @w:footer — band distances from the page edges (twips). */
  headerDistTwips?: number;
  footerDistTwips?: number;
  /** w:headerReference / w:footerReference by type — r:id into the document
   *  part's rels; the referenced parts are parsed separately. Variants are
   *  gated downstream (first by w:titlePg, even by settings evenAndOdd). */
  headerRefs?: BandRefs;
  footerRefs?: BandRefs;
  /** w:titlePg — this section has a distinct first-page header/footer. */
  titlePg?: boolean;
  /** w:cols — newspaper columns (only count > 1 is meaningful). */
  columns?: { count: number; spaceTwips?: number };
  /** w:pgNumType w:start — restart page numbering at this value. */
  pageNumberStart?: number;
}

/** Header/footer relationship ids by Word variant. */
export interface BandRefs {
  default?: string;
  first?: string;
  even?: string;
}

export interface IRDocument {
  blocks: IRBlock[];
  section: IRSection | null;
  /** Content controls found in the body (header/footer parsing extends it). */
  sdts: Record<string, IRSdtProps>;
}
