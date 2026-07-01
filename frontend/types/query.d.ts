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
  BookmarkRange,
  CharStyle,
  Document,
  DocPosition,
  DocSelection,
  FieldDef,
  ImageBlock,
  NamedStyle,
  Paragraph,
  ParaStyle,
  SdtProps,
  SectionProps,
  TableBlock,
} from "./model";

export type {
  Block,
  BookmarkRange,
  CharStyle,
  Document,
  DocPosition,
  DocSelection,
  EquationBlock,
  FieldDef,
  FieldSpec,
  ImageBlock,
  NamedStyle,
  Paragraph,
  ParaStyle,
  Run,
  SdtProps,
  SdtType,
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
// Content controls (SDTs) — the primary templating surface. Membership is an
// ordered ancestry path, so controls nest; these expose the flat list (by
// id/tag/alias), the nesting tree, and the wrapped blocks/text.

/** A content control's id paired with its properties. */
export interface SdtMatch {
  id: string;
  props: SdtProps;
}

/** A content control as a node in the nesting tree. */
export interface SdtNode extends SdtMatch {
  /** The control directly wrapping this one, or undefined for a top-level control. */
  parentId?: string;
  /** Direct child controls, in first-appearance order. */
  childIds: string[];
  /** Full ancestry outer→inner ending at this id (`[this]` for a root). */
  path: string[];
  /** Nesting depth: 0 for a root (== path.length - 1). */
  depth: number;
}

/** A content control's properties by id, or undefined. */
export declare function getSdt(doc: Document, id: string): SdtProps | undefined;
/** Every content control, as `{ id, props }`, in `doc.sdts` insertion order. */
export declare function getSdts(doc: Document): SdtMatch[];
/** Controls whose machine-readable tag (w:tag) equals `tag` (a tag is not unique). */
export declare function getSdtsByTag(doc: Document, tag: string): SdtMatch[];
/** Controls whose title/alias (w:alias) equals `alias`. */
export declare function getSdtsByAlias(doc: Document, alias: string): SdtMatch[];
/** Every content control as a tree node (parent/child links, path, depth). */
export declare function getSdtNodes(doc: Document): SdtNode[];
/** The top-level controls — those not nested inside any other. */
export declare function getSdtRoots(doc: Document): SdtNode[];
/** The controls nested directly (one level) inside `id`. */
export declare function getSdtChildren(doc: Document, id: string): SdtNode[];
/** The controls wrapping `id`, outermost→innermost (excluding `id`). */
export declare function getSdtAncestors(doc: Document, id: string): SdtNode[];
/** Every control nested anywhere below `id` (depth-first, pre-order). */
export declare function getSdtDescendants(doc: Document, id: string): SdtNode[];
/** The blocks a block-level control wraps (their `sdtPath` includes `id`). */
export declare function getSdtBlocks(doc: Document, id: string, options?: WalkOptions): Block[];
/** The plain text a control encloses (block-level whole blocks + inline runs;
 *  nested controls' text included). Blocks join with newlines. */
export declare function sdtText(doc: Document, id: string, options?: WalkOptions): string;

/** A content control's value in the shape its `type` implies. `text` is always the
 *  enclosed text; `checked` is set for checkboxes; `selected` for dropDown/comboBox
 *  (the VALUE of the matching listItem, or the raw text for combo free entry). */
export interface SdtValue {
  type: SdtProps["type"];
  text: string;
  checked?: boolean;
  selected?: string;
}
/** Read a control's typed value, or undefined if it does not exist. */
export declare function getSdtValue(doc: Document, id: string): SdtValue | undefined;

// ---------------------------------------------------------------------------
// Fields / bookmarks / notes / lists / styles / block location

/** A field definition by id, or undefined. */
export declare function getField(doc: Document, id: string): FieldDef | undefined;
/** Every tracked field (custom + built-in), in document order. */
export declare function getFields(doc: Document): FieldDef[];
/** Fields whose keyword matches `name`, case-insensitively. */
export declare function getFieldsByName(doc: Document, name: string): FieldDef[];
/** The blocks forming a region field's result (their `fieldId` === id). */
export declare function getFieldBlocks(doc: Document, id: string, options?: WalkOptions): Block[];

export interface BookmarkEntry {
  name: string;
  range: BookmarkRange;
}
/** A bookmark's character range by name, or undefined. */
export declare function getBookmark(doc: Document, name: string): BookmarkRange | undefined;
/** Every bookmark as `{ name, range }`. */
export declare function getBookmarks(doc: Document): BookmarkEntry[];

export interface NoteStory {
  id: string;
  paragraphs: Paragraph[];
}
/** Footnote stories keyed by ref id. */
export declare function getFootnotes(doc: Document): NoteStory[];
/** Endnote stories keyed by ref id. */
export declare function getEndnotes(doc: Document): NoteStory[];

export interface ListItem {
  paragraph: Paragraph;
  /** 0-based list level (OOXML w:ilvl). */
  level: number;
  /** Resolved marker text ("1.", "a.", "•", …); "" for a markerless/unknown list. */
  marker: string;
  context: BlockContext;
}
/** Paragraphs bound to a list definition, in body reading order, each with its
 *  resolved marker (mirrors the layout engine's numbering pass, body-only). */
export declare function getListItems(doc: Document, listId: string): ListItem[];

/** Enumerate the stylesheet (Word's style gallery). */
export declare function getStyles(doc: Document): NamedStyle[];
/** A named style by id, or undefined. */
export declare function getStyleById(doc: Document, styleId: string): NamedStyle | undefined;

/** "Where is this block?" — the resolved container/cell/note context of a block by
 *  id, or undefined if no block carries the id. */
export declare function blockPath(doc: Document, id: string, options?: WalkOptions): BlockContext | undefined;

// ---------------------------------------------------------------------------
// Range / text addressing

/** The text a selection covers. A collapsed/single-block selection slices that
 *  block's text (any story); a multi-block selection is supported across top-level
 *  body blocks (joined by newlines), else "". Endpoints are ordered automatically. */
export declare function rangeText(doc: Document, selection: DocSelection): string;
/** The first `DocPosition` where `needle` appears in a paragraph's text (reading
 *  order), or undefined — target an edit without computing offsets by hand. */
export declare function positionOfText(doc: Document, needle: string, options?: WalkOptions): DocPosition | undefined;

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

  /** Merge a patch onto a content control's properties (the control must exist).
   *  The control `type` is not patchable — it is always preserved. */
  setSdtProps(id: string, patch: Partial<Omit<SdtProps, "type">>): this;
  /** Set a checkbox control's state (throws if `id` is not a checkbox). */
  setCheckbox(id: string, checked: boolean): this;
  /** Fill a single-paragraph content control's text (inline or block-level),
   *  preserving nesting and clearing any placeholder. */
  setSdtText(id: string, text: string): this;
  /** Select a value on a dropDown/comboBox control (by listItem value then display;
   *  comboBox allows free text). Throws for other control kinds. */
  setSdtValue(id: string, value: string): this;
  /** Remove a content control, unwrapping it (strip its id from member paths and
   *  delete its props; nested controls survive). `deleteContents` removes the
   *  wrapped content too. Block-level members outside the top-level body throw. */
  removeSdt(id: string, options?: { deleteContents?: boolean }): this;

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
/** Where a top-level block sits in a `getPages` map: its page `index` and 0-based
 *  `order` among that page's block ids, or null. Pure over the page map (no layout). */
export declare function indexOnPage(pages: PageInfo[], blockId: string): { index: number; order: number } | null;
