// .docx import — shared types.
//
// The pipeline is: zip → XML parts → IR (intermediate representation) → model.
// The IR is a faithful-but-minimal decode of document.xml: the parser keeps
// everything it understands (even what the model can't hold yet); mapToModel
// decides what survives and emits an ImportWarning for every lossy decision.

import type { Document, FieldDef, MathRow } from "@cw/shared";

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
  /** Raw embedded image bytes, populated only when runImport is called with
   *  { collectMediaBytes: true } — the Node/backend path, which content-addresses
   *  them into its own media store instead of minting blob: URLs. Each `src`
   *  matches the ImageBlock.src the mapper assigned, so the caller can swap in a
   *  mediaId. Empty (and mediaUrls carry blob: URLs) in the browser path. */
  media: ImportMedia[];
}

/** A collected embedded image (Node/backend path). */
export interface ImportMedia {
  /** The synthetic src the mapper put on the ImageBlock (e.g. "cw-media:0"). */
  src: string;
  bytes: Uint8Array;
  mime: string;
}

/** Sink the media store routes raw image bytes to instead of creating a blob:
 *  URL. Returns the src string to stamp on the ImageBlock. */
export interface MediaCollector {
  add(bytes: Uint8Array, mime: string): string;
}

export interface RunImportOpts {
  onProgress?: (phase: ImportPhase, pct: number) => void;
  /** Collect raw image bytes (and skip blob: URL creation) for Node/backend
   *  callers that content-address media themselves. */
  collectMediaBytes?: boolean;
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
  /** w:rtl — explicit right-to-left run direction. */
  rtl?: boolean;
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
 *  these into the model's SdtProps; runs inside the control carry the sdtPath). */
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
  | { kind: "run"; text: string; props: IRRunProps; sdtPath?: string[]; fieldId?: string }
  /** Inline equation (w:p-level m:oMath) — becomes a single-U+FFFC run carrying
   *  the MathML on CharStyle.equation. `props` carries run-like metadata the
   *  w:hyperlink backfill tags on (linkRelId/linkAnchor) so a linked equation
   *  keeps its hyperlink. */
  | { kind: "mathInline"; root: MathRow; props?: IRRunProps }
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
      /** Set for wrapNone anchors (image sits behind/in-front of text, not in
       *  the flow). Maps to ImageBlock.anchor — positioned absolutely, no flow
       *  height, no text reflow. */
      anchorFloat?: {
        behind: boolean;
        offsetXEmu: number;
        offsetYEmu: number;
        relFromH: "page" | "margin" | "column" | "leftMargin" | "rightMargin" | "character";
        relFromV: "page" | "margin" | "paragraph" | "line" | "topMargin" | "bottomMargin";
        decorative?: boolean;
        /** wp:anchor @relativeHeight — stacking order within the layer. */
        z?: number;
      };
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
  /** w:outlineLvl — 0-8. Heading styles set it; drives TOC `\u`/heading detection. */
  outlineLevel?: number;
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
  sectionColumns?: { count: number; spaceTwips?: number; sep?: boolean; cols?: { wTwips: number; spaceTwips: number }[] };
  /** The ending section's w:pgNumType w:start (page-number restart). */
  sectionPageNumberStart?: number;
  /** The ending section's w:pgMar header/footer band distances (twips). */
  sectionHeaderDistTwips?: number;
  sectionFooterDistTwips?: number;
  /** The ending section's w:pgBorders (raw, mapped downstream). */
  sectionPgBorders?: IRPageBorders;
  /** The ending section's sectPr declares its own header/footer references —
   *  when such a section is flowed (not page-broken) its bands aren't applied. */
  sectionHasBands?: boolean;
  /** w:pPr/w:rPr — the paragraph MARK's run formatting. Word styles empty
   *  paragraphs (and the ¶ itself) with this; we use it for empty-run style. */
  markRunProps?: IRRunProps;
  /** w:bidi — paragraph base direction is RTL. */
  direction?: "ltr" | "rtl";
}

/** A w:bookmarkStart / w:bookmarkEnd seen while walking a paragraph: its w:id
 *  (pairs start↔end, possibly across paragraphs), w:name (start only), and the
 *  UTF-16 offset within the paragraph text where it sits. */
export interface BookmarkMarker {
  id: string;
  name?: string;
  kind: "start" | "end";
  offset: number;
}

export interface IRParagraph {
  kind: "paragraph";
  props: IRParaProps;
  inlines: IRInline[];
  /** w:bookmarkStart names anchored in this paragraph (TOC/cross-ref heuristics). */
  bookmarks?: string[];
  /** Start/end markers with offsets — resolved to model ranges in mapToModel. */
  bookmarkMarkers?: BookmarkMarker[];
  /** Custom-field result membership — mapped onto Block.fieldId. */
  fieldId?: string;
  /** Block-level content-control ancestry (outer→inner) — mapped onto Block.sdtPath. */
  sdtPath?: string[];
  /** Set when this paragraph holds a `TOC` field — its verbatim instrText. Lets a
   *  headless render anchor a freshly-built TOC at this exact block (mapToModel
   *  records the resulting block id), instead of fragile ordinal counting. */
  tocField?: string;
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
  /** w:tcPr/w:tcW — preferred cell width. `abs` carries twips (dxa); `pct` carries
   *  a 0..1 fraction of the table width (w:w is fiftieths of a percent). Consulted
   *  only in autofit modes (→ TableCell.preferredWidth). */
  preferredWidth?: { type: "abs"; twips: number } | { type: "pct"; frac: number };
}

export interface IRTableRow {
  cells: IRTableCell[];
}

export interface IRTable {
  kind: "table";
  rows: IRTableRow[];
  /** w:tblGrid/w:gridCol widths — become TableBlock.colFractions. */
  colWidthsTwips?: number[];
  /** w:tblPr/w:tblLayout + w:tblW → column-sizing strategy. Absent = fixed. */
  widthMode?: "fixed" | "autofitContents" | "autofitWindow";
  /** w:tblPr/w:tblW preferred TOTAL width on a fixed-layout table (→ TableBlock.preferredWidth).
   *  `preferredWidthTwips` = absolute dxa; `preferredWidthPct` = 0..100 percent. */
  preferredWidthTwips?: number;
  preferredWidthPct?: number;
  /** w:tblPr/w:jc → table alignment within the content width (→ TableBlock.align). */
  align?: "left" | "center" | "right";
  /** w:tblPr/w:tblStyle — table style id (its borders/shd are the cascade base). */
  styleId?: string;
  /** w:tblPr/w:tblLook — which conditional bands of the table style are active. */
  look?: { firstRow?: boolean; lastRow?: boolean; firstCol?: boolean; lastCol?: boolean; bandRows?: boolean; bandCols?: boolean };
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
  /** Custom-field result membership — mapped onto Block.fieldId. */
  fieldId?: string;
  /** Block-level content-control ancestry (outer→inner) — mapped onto Block.sdtPath. */
  sdtPath?: string[];
}

/** w:oMathPara — a display (block) equation. The OMML is converted to a MathML
 *  AST at parse time (see import/docx/documentParser → ommlToMathml). */
export interface IRMath {
  kind: "math";
  root: MathRow;
  display: boolean;
  /** m:oMathParaPr/m:jc — mapped onto EquationBlock.align. */
  align?: "left" | "center" | "right";
  /** Custom-field result membership — mapped onto Block.fieldId. */
  fieldId?: string;
  /** Block-level content-control ancestry — mapped onto Block.sdtPath. */
  sdtPath?: string[];
}

export type IRBlock = IRParagraph | IRTable | IRMath;

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
  /** w:cols — newspaper columns (only count > 1 is meaningful). `sep` = a
   *  separator line; `cols` = explicit per-column widths (w:equalWidth="0"). */
  columns?: {
    count: number;
    spaceTwips?: number;
    sep?: boolean;
    cols?: { wTwips: number; spaceTwips: number }[];
  };
  /** w:pgNumType w:start — restart page numbering at this value. */
  pageNumberStart?: number;
  /** w:pgBorders — page border box (raw twips/eighth-points, mapped downstream). */
  pageBorders?: IRPageBorders;
}

/** w:pgBorders edge as parsed (w:sz eighth-points, w:space points). */
export interface IRPageBorderEdge {
  style: string;
  sz?: number;
  space?: number;
  color?: string;
}
export interface IRPageBorders {
  top?: IRPageBorderEdge;
  right?: IRPageBorderEdge;
  bottom?: IRPageBorderEdge;
  left?: IRPageBorderEdge;
  offsetFrom?: "page" | "text";
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
  /** Custom (non-built-in) complex fields captured in the body, keyed by id. */
  fields?: Record<string, FieldDef>;
  /** w:document/w:background @w:color — document-global page fill ("RRGGBB"). */
  pageColorHex?: string;
}
