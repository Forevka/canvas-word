// .docx import — shared types.
//
// The pipeline is: zip → XML parts → IR (intermediate representation) → model.
// The IR is a faithful-but-minimal decode of document.xml: the parser keeps
// everything it understands (even what the model can't hold yet); mapToModel
// decides what survives and emits an ImportWarning for every lossy decision.

import type { Document } from "../../model/document";

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
  /** w:vanish — hidden text (Word shows it only with ¶ marks on). */
  vanish?: boolean;
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
   *  paragraph). page=true for w:br w:type="page": the following content
   *  starts a new page (maps to ParaStyle.pageBreakBefore). */
  | { kind: "break"; page?: boolean }
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
  /** Negative = hanging indent (w:hanging). */
  indentFirstLineTwips?: number;
  /** w:keepNext — maps onto ParaStyle.keepWithNext. */
  keepWithNext?: boolean;
  /** w:pageBreakBefore — this paragraph starts a new page. */
  pageBreakBefore?: boolean;
  /** w:pPr/w:sectPr — this paragraph ENDS a section. "page" (nextPage/odd/even)
   *  implies the following content starts a new page; "continuous" doesn't. */
  sectionBreak?: "page" | "continuous";
  /** w:pPr/w:rPr — the paragraph MARK's run formatting. Word styles empty
   *  paragraphs (and the ¶ itself) with this; we use it for empty-run style. */
  markRunProps?: IRRunProps;
}

export interface IRParagraph {
  kind: "paragraph";
  props: IRParaProps;
  inlines: IRInline[];
}

export interface IRTableCell {
  /** Full block content: paragraphs, images (inside paragraphs), nested tables. */
  blocks: IRBlock[];
  /** w:gridSpan — columns this cell covers (1 = normal). Maps to TableCell.colSpan. */
  gridSpan: number;
  /** w:vMerge continuation — this cell is swallowed by the cell above. */
  vMergeContinue: boolean;
}

export interface IRTableRow {
  cells: IRTableCell[];
}

export interface IRTable {
  kind: "table";
  rows: IRTableRow[];
  /** w:tblGrid/w:gridCol widths — become TableBlock.colFractions. */
  colWidthsTwips?: number[];
}

export type IRBlock = IRParagraph | IRTable;

export interface IRSection {
  pageWidthTwips?: number;
  pageHeightTwips?: number;
  marginTwips?: { top: number; right: number; bottom: number; left: number };
  /** w:headerReference / w:footerReference (default type) — r:id into the
   *  document part's rels; the referenced parts are parsed separately. */
  headerRelId?: string;
  footerRelId?: string;
}

export interface IRDocument {
  blocks: IRBlock[];
  section: IRSection | null;
  /** Content controls found in the body (header/footer parsing extends it). */
  sdts: Record<string, IRSdtProps>;
}
