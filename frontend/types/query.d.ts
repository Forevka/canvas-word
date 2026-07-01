// Public type surface for @forevka/wordcanvas/query — the document query + edit
// API (the rough analog of .NET's WordprocessingDocument access). Hand-written,
// self-contained (model types come from ./model, shared with wordcanvas.d.ts).
//
//   import { findParagraphs, DocumentEditor, getPages } from "@forevka/wordcanvas/query";
//   const ed = new DocumentEditor(doc);
//   for (const m of findParagraphs(ed.doc, "DRAFT")) ed.setParagraphText(m.paragraph.id, "FINAL");

import type {
  BandContainer,
  Block,
  CharStyle,
  Document,
  ImageBlock,
  Paragraph,
  ParaStyle,
  SectionProps,
  TableBlock,
} from "./model";

export type {
  Block,
  CharStyle,
  Document,
  DocPosition,
  EquationBlock,
  ImageBlock,
  Paragraph,
  ParaStyle,
  Run,
  SectionProps,
  TableBlock,
} from "./model";

// ---------------------------------------------------------------------------
// Traversal / query

/** A top-level story: the body or one of the six header/footer band stories.
 *  Note bodies report `"body"` and set `note` instead. */
export type Container = "body" | BandContainer;

/** Where a visited block sits. `cell`/`note` are set when the block is nested
 *  below the top level of its story. */
export interface BlockContext {
  container: Container;
  /** Set when the block lives inside a table cell (innermost cell if nested). */
  cell?: { tableId: string; row: number; col: number };
  /** Set when the block belongs to a footnote/endnote body. */
  note?: { kind: "footnote" | "endnote"; id: string };
}

export type BlockVisitor = (block: Block, ctx: BlockContext) => void;

/** Which stories `walk` descends into. All default to `true`. */
export interface WalkOptions {
  bands?: boolean;
  notes?: boolean;
  cells?: boolean;
}

export interface ParagraphMatch {
  paragraph: Paragraph;
  /** The paragraph's concatenated run text (what was tested). */
  text: string;
  context: BlockContext;
}

export interface ResolvedSection {
  index: number;
  props: SectionProps;
  startBlock: number;
  endBlock: number;
  breakType: "nextPage" | "evenPage" | "oddPage";
}

/** Visit every block, descending (by default) into table cells, header/footer
 *  bands, and note bodies. */
export declare function walk(doc: Document, visit: BlockVisitor, options?: WalkOptions): void;
/** Plain text of a block (tables join cells with tabs, rows with newlines). */
export declare function textOf(block: Block): string;
export declare function getParagraphs(doc: Document, options?: WalkOptions): Paragraph[];
export declare function getTables(doc: Document, options?: WalkOptions): TableBlock[];
export declare function getImages(doc: Document, options?: WalkOptions): ImageBlock[];
/** Find paragraphs by substring or RegExp; each match carries its context. */
export declare function findParagraphs(doc: Document, pattern: string | RegExp, options?: WalkOptions): ParagraphMatch[];
export declare function getBlockById(doc: Document, id: string, options?: WalkOptions): Block | undefined;
export declare function getParagraphById(doc: Document, id: string, options?: WalkOptions): Paragraph | undefined;
export declare function getTableById(doc: Document, id: string, options?: WalkOptions): TableBlock | undefined;
export declare function getImageById(doc: Document, id: string, options?: WalkOptions): ImageBlock | undefined;
/** Enumerate sections (page geometry + the block range each covers). */
export declare function getSections(doc: Document): ResolvedSection[];
export declare function resolveSections(doc: Document): ResolvedSection[];

// ---------------------------------------------------------------------------
// Edit facade

export interface InsertParagraphOptions {
  position?: "before" | "after";
  style?: ParaStyle;
  runStyle?: CharStyle;
}

/** Ergonomic, headless editing facade over the operation engine, with undo/redo.
 *  Every edit swaps `doc` for a new immutable value (structural sharing). */
export declare class DocumentEditor {
  constructor(doc: Document);
  get doc(): Document;
  get canUndo(): boolean;
  get canRedo(): boolean;
  get lastInsertedId(): string | null;

  setParagraphText(blockId: string, text: string, style?: CharStyle): this;
  insertText(blockId: string, offset: number, text: string, style?: CharStyle): this;
  deleteText(blockId: string, start: number, end: number): this;
  replaceText(blockId: string, start: number, end: number, text: string, style?: CharStyle): this;
  setParagraphStyle(blockId: string, patch: Partial<ParaStyle>): this;
  removeBlock(blockId: string): this;
  insertParagraph(refBlockId: string, text: string, options?: InsertParagraphOptions): this;

  undo(): boolean;
  redo(): boolean;

  find(pattern: string | RegExp): ParagraphMatch[];
  getParagraph(blockId: string): Paragraph | undefined;
}

// ---------------------------------------------------------------------------
// Page query (layout-backed)

export interface PageInfo {
  index: number;
  number: number;
  widthPx: number;
  heightPx: number;
  marginPx: { top: number; right: number; bottom: number; left: number };
  contentTopPx: number;
  contentBottomPx: number;
  pageColorHex?: string;
  blockIds: string[];
}

export interface GetPagesOptions {
  fontRegistry?: unknown;
  engineOptions?: {
    cjkFallback?: string | null;
    cjkLocale?: string;
    arabicFallback?: string | null;
    hebrewFallback?: string | null;
  };
}

/** Lay the document out and return a serializable per-page map. Page numbers can
 *  shift after edits, so re-run after mutating. Requires text measurement
 *  (pretext) to be initialized, like rendering/export. */
export declare function getPages(doc: Document, options?: GetPagesOptions): PageInfo[];
/** The displayed page number a top-level block first appears on, or null. */
export declare function pageOfBlock(doc: Document, blockId: string, options?: GetPagesOptions): number | null;
