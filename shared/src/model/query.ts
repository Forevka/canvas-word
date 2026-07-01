// Document query API — a read-only traversal layer over the flat `Document`
// model, the rough analog of .NET's `WordprocessingDocument` descendant queries
// (`Body.Descendants<Paragraph>()`, find-by-text, section enumeration). Pure and
// DOM-free like the rest of the shared core; it never mutates. Page-level queries
// ("what's on page N") are NOT here — pages exist only after layout, so they live
// in the layout package.

import type { Block, BookmarkRange, Document, FieldDef, ImageBlock, Paragraph, Run, SdtProps, TableBlock, TableCell } from "./document";
import { BAND_CONTAINERS } from "./document";
import type { Container } from "./ops";
import { fullSdtChain } from "./sdt";
import { markerText, type ListDefinition } from "./lists";
import type { NamedStyle } from "./stylesheet";
import type { DocPosition, DocSelection } from "./position";
import { textOfRuns } from "./text";
import { resolveSections, type ResolvedSection } from "./sections";

/** A table cell coordinate (which table, which row/col). */
export interface CellRef {
  tableId: string;
  row: number;
  col: number;
}

/** Where a visited block sits. `cell`/`note` are set when the block is nested
 *  below the top level of its story. */
export interface BlockContext {
  /** The top-level story this block belongs to (body or a header/footer band).
   *  Note bodies report `"body"` and set `note` instead. */
  container: Container;
  /** Set when the block lives inside a table cell (INNERMOST cell if nested). */
  cell?: CellRef;
  /** The full cell ancestry outer→inner, present ONLY for a block nested in a
   *  TABLE-WITHIN-A-CELL (≥2 levels deep) — so single-cell blocks keep the simple
   *  `{ container, cell }` shape. Use this to answer "which outer table/cell". */
  cellPath?: CellRef[];
  /** Set when the block belongs to a footnote/endnote body. */
  note?: { kind: "footnote" | "endnote"; id: string };
}

export type BlockVisitor = (block: Block, ctx: BlockContext) => void;

/** Which stories `walk` descends into. All default to `true`. */
export interface WalkOptions {
  /** Header/footer band stories on `doc.section`. */
  bands?: boolean;
  /** Footnote/endnote bodies. */
  notes?: boolean;
  /** Descend into table cells (and any tables nested within them). */
  cells?: boolean;
}

type ResolvedWalkOptions = Required<WalkOptions>;

function walkStory(
  blocks: Block[],
  ctx: BlockContext,
  opts: ResolvedWalkOptions,
  visit: BlockVisitor,
  cellStack: CellRef[] = [],
): void {
  for (const block of blocks) {
    visit(block, ctx);
    if (opts.cells && block.kind === "table") {
      for (let row = 0; row < block.rows.length; row++) {
        const cells = block.rows[row]!.cells;
        for (let col = 0; col < cells.length; col++) {
          const cell: CellRef = { tableId: block.id, row, col };
          const stack = [...cellStack, cell];
          // Spread ctx so a table nested in a note body keeps its `note` membership.
          // `cellPath` is attached only when nested ≥2 deep, so single-cell blocks
          // keep the historical `{ container, cell }` shape.
          const childCtx: BlockContext = { ...ctx, cell };
          if (stack.length >= 2) childCtx.cellPath = stack;
          walkStory(cells[col]!.blocks, childCtx, opts, visit, stack);
        }
      }
    }
  }
}

/** Visit every block in the document, descending (by default) into table cells,
 *  header/footer bands, and note bodies. The single traversal primitive the rest
 *  of this module builds on. */
export function walk(doc: Document, visit: BlockVisitor, options: WalkOptions = {}): void {
  const opts: ResolvedWalkOptions = { bands: true, notes: true, cells: true, ...options };
  walkStory(doc.blocks, { container: "body" }, opts, visit);
  if (opts.bands) {
    for (const band of BAND_CONTAINERS) {
      const blocks = doc.section[band];
      if (blocks) walkStory(blocks, { container: band }, opts, visit);
    }
    // Bands can also be attached to an EARLIER section via its section-break
    // paragraph's props (see effectiveSection); those stories live nowhere else,
    // so visit any explicitly present here. Absent props inherit doc.section,
    // already visited above — so this never double-counts.
    for (const block of doc.blocks) {
      if (block.kind !== "paragraph" || !block.style.sectionBreak) continue;
      for (const band of BAND_CONTAINERS) {
        const blocks = block.style.sectionBreak.props[band];
        if (blocks) walkStory(blocks, { container: band }, opts, visit);
      }
    }
  }
  if (opts.notes) {
    for (const [id, paras] of Object.entries(doc.footnotes ?? {})) {
      walkStory(paras, { container: "body", note: { kind: "footnote", id } }, opts, visit);
    }
    for (const [id, paras] of Object.entries(doc.endnotes ?? {})) {
      walkStory(paras, { container: "body", note: { kind: "endnote", id } }, opts, visit);
    }
  }
}

/** A table cell's text — its blocks' text joined by newlines, empties dropped. The
 *  single definition of cell-text joining (shared by `textOf` and the editor). */
export function textOfCell(cell: TableCell): string {
  return cell.blocks.map(textOf).filter((t) => t.length > 0).join("\n");
}

/** Where a visited run sits: its paragraph's `BlockContext`, plus the paragraph,
 *  the run's index, and its FULL enclosing content-control chain (block ancestry
 *  ++ inline ancestry, outer→inner). */
export interface RunContext extends BlockContext {
  block: Paragraph;
  runIndex: number;
  /** `block.sdtPath ++ run.style.sdtPath`, outer→inner; empty if none. */
  sdtChain: string[];
}

export type RunVisitor = (run: Run, ctx: RunContext) => void;

/** Visit every run in the document (only paragraphs have runs), in reading order,
 *  with its paragraph context and full sdt chain. The run-level companion to
 *  `walk` — the primitive behind inline content-control / field / bookmark text
 *  resolution. Descends into the same stories as `walk` (options apply). */
export function walkRuns(doc: Document, visit: RunVisitor, options?: WalkOptions): void {
  walk(doc, (block, ctx) => {
    if (block.kind !== "paragraph") return;
    block.runs.forEach((run, runIndex) => {
      visit(run, { ...ctx, block, runIndex, sdtChain: fullSdtChain(block, run.style) });
    });
  }, options);
}

/** The plain text of a block. Tables join cells with tabs and rows with newlines;
 *  images and equations have no text. */
export function textOf(block: Block): string {
  switch (block.kind) {
    case "paragraph":
      return textOfRuns(block.runs);
    case "table":
      // Preserve paragraph boundaries inside a cell (newline), columns with tabs,
      // rows with newlines. Empty blocks (images/equations) drop out so they don't
      // inject blank lines.
      return block.rows.map((row) => row.cells.map(textOfCell).join("\t")).join("\n");
    default:
      return "";
  }
}

/** Every paragraph in the document (body, cells, bands, and notes by default). */
export function getParagraphs(doc: Document, options?: WalkOptions): Paragraph[] {
  const out: Paragraph[] = [];
  walk(doc, (b) => {
    if (b.kind === "paragraph") out.push(b);
  }, options);
  return out;
}

/** Every table in the document. */
export function getTables(doc: Document, options?: WalkOptions): TableBlock[] {
  const out: TableBlock[] = [];
  walk(doc, (b) => {
    if (b.kind === "table") out.push(b);
  }, options);
  return out;
}

/** Every image in the document. */
export function getImages(doc: Document, options?: WalkOptions): ImageBlock[] {
  const out: ImageBlock[] = [];
  walk(doc, (b) => {
    if (b.kind === "image") out.push(b);
  }, options);
  return out;
}

export interface ParagraphMatch {
  paragraph: Paragraph;
  /** The paragraph's concatenated run text (what was tested). */
  text: string;
  context: BlockContext;
}

/** Find paragraphs whose text matches. A string matches by substring; a RegExp is
 *  tested against the whole paragraph text (a global regex is reset each test so
 *  its `lastIndex` state can't cause missed matches). */
export function findParagraphs(doc: Document, pattern: string | RegExp, options?: WalkOptions): ParagraphMatch[] {
  const test = typeof pattern === "string"
    ? (t: string) => t.includes(pattern)
    : (t: string) => {
        pattern.lastIndex = 0;
        return pattern.test(t);
      };
  const out: ParagraphMatch[] = [];
  walk(doc, (block, context) => {
    if (block.kind !== "paragraph") return;
    const text = textOfRuns(block.runs);
    if (test(text)) out.push({ paragraph: block, text, context });
  }, options);
  return out;
}

/** The first block anywhere in the document with this id, or undefined. */
export function getBlockById(doc: Document, id: string, options?: WalkOptions): Block | undefined {
  let found: Block | undefined;
  walk(doc, (b) => {
    if (found === undefined && b.id === id) found = b;
  }, options);
  return found;
}

export function getParagraphById(doc: Document, id: string, options?: WalkOptions): Paragraph | undefined {
  const b = getBlockById(doc, id, options);
  return b?.kind === "paragraph" ? b : undefined;
}

export function getTableById(doc: Document, id: string, options?: WalkOptions): TableBlock | undefined {
  const b = getBlockById(doc, id, options);
  return b?.kind === "table" ? b : undefined;
}

export function getImageById(doc: Document, id: string, options?: WalkOptions): ImageBlock | undefined {
  const b = getBlockById(doc, id, options);
  return b?.kind === "image" ? b : undefined;
}

/** Enumerate the document's sections (page size, margins, columns, header/footer
 *  distances…) with the top-level block range each one covers. A single-section
 *  document returns one entry. Thin friendly alias over `resolveSections`. */
export function getSections(doc: Document): ResolvedSection[] {
  return resolveSections(doc);
}

// ---------------------------------------------------------------------------
// Content controls (OOXML w:sdt) — the primary templating surface.
//
// Properties for every control live in `doc.sdts` keyed by id. Membership is an
// ORDERED ANCESTRY PATH (outer→inner): block-level controls put their ids on
// `Block.sdtPath`, inline controls on `run.style.sdtPath`; a run's full enclosing
// chain is `block.sdtPath ++ run.style.sdtPath`. Because membership is a path,
// controls NEST — these helpers expose both the flat list (by id/tag/alias) and
// the nesting tree (roots/children/ancestors/descendants), plus the wrapped
// blocks and text.

/** A content control's id paired with its properties. */
export interface SdtMatch {
  id: string;
  props: SdtProps;
}

/** A content control as a node in the nesting tree. */
export interface SdtNode extends SdtMatch {
  /** The control DIRECTLY wrapping this one, or undefined for a top-level control. */
  parentId?: string;
  /** Direct child controls, in first-appearance order. */
  childIds: string[];
  /** Full ancestry outer→inner ENDING at this id (`[this]` for a root,
   *  `[outer, this]` when nested one level, …). */
  path: string[];
  /** Nesting depth: 0 for a root, 1 for its children, … (== `path.length - 1`). */
  depth: number;
}

/** A content control's properties by id, or undefined. */
export function getSdt(doc: Document, id: string): SdtProps | undefined {
  return doc.sdts?.[id];
}

/** Every content control, as `{ id, props }`, in `doc.sdts` insertion order
 *  (import adds controls as it encounters them — close to, but not guaranteed to
 *  be, document reading order). */
export function getSdts(doc: Document): SdtMatch[] {
  return Object.entries(doc.sdts ?? {}).map(([id, props]) => ({ id, props }));
}

/** Controls whose machine-readable tag (w:tag) equals `tag`. A tag is not unique
 *  (a template can repeat a field), so this returns every match. */
export function getSdtsByTag(doc: Document, tag: string): SdtMatch[] {
  return getSdts(doc).filter((s) => s.props.tag === tag);
}

/** Controls whose title/alias (w:alias) equals `alias`. */
export function getSdtsByAlias(doc: Document, alias: string): SdtMatch[] {
  return getSdts(doc).filter((s) => s.props.alias === alias);
}

/** Derive parent→child control relationships from every membership path present
 *  in the document (block `sdtPath`s + full run chains). In a path `[a, b, c]`,
 *  `a` wraps `b` wraps `c`. */
function sdtRelations(doc: Document): { parentOf: Map<string, string>; childrenOf: Map<string, string[]> } {
  const parentOf = new Map<string, string>();
  const childrenOf = new Map<string, string[]>();
  const seenEdge = new Set<string>();
  const record = (path: string[] | undefined): void => {
    if (!path || path.length < 2) return;
    for (let i = 1; i < path.length; i++) {
      const parent = path[i - 1]!;
      const child = path[i]!;
      if (!parentOf.has(child)) parentOf.set(child, parent);
      const edge = `${parent} ${child}`;
      if (seenEdge.has(edge)) continue;
      seenEdge.add(edge);
      const kids = childrenOf.get(parent) ?? [];
      kids.push(child);
      childrenOf.set(parent, kids);
    }
  };
  walk(doc, (block) => {
    record(block.sdtPath);
    if (block.kind === "paragraph") {
      for (const run of block.runs) record(fullSdtChain(block, run.style));
    }
  });
  return { parentOf, childrenOf };
}

/** Build every declared control's tree node, keyed by id. Nodes come from
 *  `doc.sdts` (the property source of truth); relationships come from the paths.
 *  Path ids missing from `doc.sdts` are ignored (defensive — the model declares
 *  every path id). */
function sdtNodeMap(doc: Document): Map<string, SdtNode> {
  const sdts = doc.sdts ?? {};
  const has = (id: string): boolean => Object.prototype.hasOwnProperty.call(sdts, id);
  const { parentOf, childrenOf } = sdtRelations(doc);
  // Ancestry via parent walk, cycle-guarded; stop at the first parent not declared.
  const pathOf = (id: string): string[] => {
    const chain: string[] = [id];
    const seen = new Set<string>([id]);
    for (let cur = parentOf.get(id); cur !== undefined && has(cur) && !seen.has(cur); cur = parentOf.get(cur)) {
      chain.unshift(cur);
      seen.add(cur);
    }
    return chain;
  };
  const nodes = new Map<string, SdtNode>();
  for (const [id, props] of Object.entries(sdts)) {
    const rawParent = parentOf.get(id);
    const parentId = rawParent !== undefined && has(rawParent) ? rawParent : undefined;
    const childIds = (childrenOf.get(id) ?? []).filter(has);
    const path = pathOf(id);
    nodes.set(id, {
      id,
      props,
      ...(parentId !== undefined ? { parentId } : {}),
      childIds,
      path,
      depth: path.length - 1,
    });
  }
  return nodes;
}

/** Every content control as a tree node (parent/child links, path, depth). */
export function getSdtNodes(doc: Document): SdtNode[] {
  return [...sdtNodeMap(doc).values()];
}

/** The top-level controls — those not nested inside any other. */
export function getSdtRoots(doc: Document): SdtNode[] {
  return getSdtNodes(doc).filter((n) => n.parentId === undefined);
}

/** The controls nested DIRECTLY (one level) inside `id`. Empty if `id` has no
 *  children or does not exist. */
export function getSdtChildren(doc: Document, id: string): SdtNode[] {
  const nodes = sdtNodeMap(doc);
  const self = nodes.get(id);
  if (!self) return [];
  return self.childIds.map((c) => nodes.get(c)).filter((n): n is SdtNode => n !== undefined);
}

/** The controls wrapping `id`, outermost→innermost (excluding `id` itself). */
export function getSdtAncestors(doc: Document, id: string): SdtNode[] {
  const nodes = sdtNodeMap(doc);
  const self = nodes.get(id);
  if (!self) return [];
  return self.path.slice(0, -1).map((a) => nodes.get(a)).filter((n): n is SdtNode => n !== undefined);
}

/** Every control nested ANYWHERE below `id` (depth-first, pre-order). */
export function getSdtDescendants(doc: Document, id: string): SdtNode[] {
  const nodes = sdtNodeMap(doc);
  const out: SdtNode[] = [];
  const visit = (nid: string): void => {
    for (const child of nodes.get(nid)?.childIds ?? []) {
      const c = nodes.get(child);
      if (!c) continue;
      out.push(c);
      visit(child);
    }
  };
  if (nodes.has(id)) visit(id);
  return out;
}

/** The blocks a block-level control wraps (their `sdtPath` includes `id`). Empty
 *  for a purely inline control (its content is runs — see `sdtText`). Blocks
 *  wrapped by controls NESTED under `id` are included too, since membership is by
 *  path containment. */
export function getSdtBlocks(doc: Document, id: string, options?: WalkOptions): Block[] {
  const out: Block[] = [];
  walk(doc, (b) => {
    if (b.sdtPath?.includes(id)) out.push(b);
  }, options);
  return out;
}

/** The plain text a control encloses — the "read the value" half of the template
 *  round-trip. Covers block-level membership (whole blocks whose `sdtPath`
 *  includes `id`) and inline membership (the matching runs of a paragraph whose
 *  `style.sdtPath` includes `id`). Text inside NESTED controls is naturally
 *  included (membership is by path containment). Blocks join with newlines. */
export function sdtText(doc: Document, id: string, options?: WalkOptions): string {
  const parts: string[] = [];
  walk(doc, (b) => {
    if (b.sdtPath?.includes(id)) {
      const t = textOf(b);
      if (t.length > 0) parts.push(t);
      return;
    }
    if (b.kind === "paragraph") {
      const runs = b.runs.filter((r) => r.style.sdtPath?.includes(id));
      if (runs.length > 0) parts.push(textOfRuns(runs));
    }
  }, options);
  return parts.join("\n");
}

/** A content control's value, read in the shape its `type` implies — the typed
 *  companion to `sdtText`. `text` is always the enclosed text; `checked` is set
 *  for checkboxes; `selected` is set for dropDown/comboBox — the VALUE of the
 *  listItem whose display matches the current text. When nothing matches, a combo
 *  box (free text) reports the raw text, while a dropdown (a closed option set)
 *  leaves `selected` undefined. Undefined when the control does not exist. */
export interface SdtValue {
  type: SdtProps["type"];
  text: string;
  checked?: boolean;
  selected?: string;
}

export function getSdtValue(doc: Document, id: string): SdtValue | undefined {
  const props = doc.sdts?.[id];
  if (!props) return undefined;
  const value: SdtValue = { type: props.type, text: sdtText(doc, id) };
  if (props.type === "checkbox") value.checked = props.checked ?? false;
  if (props.type === "dropDown" || props.type === "comboBox") {
    const match = props.listItems?.find((i) => i.display === value.text);
    // A dropdown is a closed set: an unmatched value is not a valid selection, so
    // leave it undefined. A combo box accepts free text, so echo the raw text.
    if (match) value.selected = match.value;
    else if (props.type === "comboBox") value.selected = value.text;
  }
  return value;
}

// ---------------------------------------------------------------------------
// Fields (OOXML complex fields) — definitions live in `doc.fields`; the result
// region is the blocks (or runs) carrying the id as `fieldId`. Covers custom
// (host-resolved) fields and the built-ins the editor understands (PAGE/DATE/…).

/** A field definition by id, or undefined. */
export function getField(doc: Document, id: string): FieldDef | undefined {
  return doc.fields?.[id];
}

/** Every tracked field definition (custom + built-in), in `doc.fields` order. */
export function getFields(doc: Document): FieldDef[] {
  return Object.values(doc.fields ?? {});
}

/** Fields whose keyword matches `name`, case-insensitively (e.g. "PAGE", "MYCHART"). */
export function getFieldsByName(doc: Document, name: string): FieldDef[] {
  const want = name.toUpperCase();
  return getFields(doc).filter((f) => f.name.toUpperCase() === want);
}

/** The blocks forming a region field's result (their `fieldId` === id). Empty for
 *  a purely inline field (its result is runs, not blocks). */
export function getFieldBlocks(doc: Document, id: string, options?: WalkOptions): Block[] {
  const out: Block[] = [];
  walk(doc, (b) => {
    if (b.fieldId === id) out.push(b);
  }, options);
  return out;
}

// ---------------------------------------------------------------------------
// Bookmarks — name → character range (may span paragraphs). Resolving a range to
// its text needs cross-block offset math (see `rangeText`, a follow-up); these
// expose the ranges themselves.

export interface BookmarkEntry {
  name: string;
  range: BookmarkRange;
}

/** A bookmark's character range by name, or undefined. */
export function getBookmark(doc: Document, name: string): BookmarkRange | undefined {
  return doc.bookmarks?.[name];
}

/** Every bookmark as `{ name, range }`, in `doc.bookmarks` order. */
export function getBookmarks(doc: Document): BookmarkEntry[] {
  return Object.entries(doc.bookmarks ?? {}).map(([name, range]) => ({ name, range }));
}

// ---------------------------------------------------------------------------
// Footnote / endnote stories, keyed by their reference id (the value a run's
// `footnoteRef`/`endnoteRef` points at).

export interface NoteStory {
  id: string;
  paragraphs: Paragraph[];
}

export function getFootnotes(doc: Document): NoteStory[] {
  return Object.entries(doc.footnotes ?? {}).map(([id, paragraphs]) => ({ id, paragraphs }));
}

export function getEndnotes(doc: Document): NoteStory[] {
  return Object.entries(doc.endnotes ?? {}).map(([id, paragraphs]) => ({ id, paragraphs }));
}

// ---------------------------------------------------------------------------
// List items — the paragraphs bound to one list definition, in body reading
// order, each with its resolved marker ("1.", "a.", "•", …). Mirrors the layout
// engine's numbering pass (per-level counters; incrementing a level resets deeper
// ones) but purely, so callers get markers without a layout. Counting runs over
// the BODY (incl. table cells) only — exactly like the engine — so list markers in
// header/footer bands are not part of this sequence.

export interface ListItem {
  paragraph: Paragraph;
  /** 0-based list level (OOXML w:ilvl). */
  level: number;
  /** Resolved marker text ("1.", "a.", "•", …); "" for a markerless/unknown list. */
  marker: string;
  context: BlockContext;
}

export function getListItems(doc: Document, listId: string): ListItem[] {
  const def: ListDefinition | undefined = doc.lists?.[listId];
  const stack: number[] = [];
  const out: ListItem[] = [];
  walk(doc, (block, context) => {
    if (block.kind !== "paragraph") return;
    const ref = block.style.list;
    if (!ref || ref.listId !== listId) return;
    const lvl = def?.levels[Math.min(ref.level, def.levels.length - 1)];
    stack[ref.level] = (stack[ref.level] ?? (lvl?.start ?? 1) - 1) + 1;
    stack.length = ref.level + 1; // incrementing a level resets deeper ones
    const marker = def ? markerText(def, ref.level, stack) : "";
    out.push({ paragraph: block, level: ref.level, marker, context });
  }, { bands: false, notes: false });
  return out;
}

// ---------------------------------------------------------------------------
// Stylesheet enumeration (Word's style gallery).

export function getStyles(doc: Document): NamedStyle[] {
  return doc.stylesheet?.styles ?? [];
}

export function getStyleById(doc: Document, styleId: string): NamedStyle | undefined {
  return doc.stylesheet?.styles.find((s) => s.id === styleId);
}

// ---------------------------------------------------------------------------
// "Where is this block?" — the resolved container/cell/note context of a block by
// id (the same `BlockContext` walk reports), for a nested block. Undefined when no
// block carries the id.

export function blockPath(doc: Document, id: string, options?: WalkOptions): BlockContext | undefined {
  let found: BlockContext | undefined;
  walk(doc, (b, ctx) => {
    if (found === undefined && b.id === id) found = ctx;
  }, options);
  return found;
}

// ---------------------------------------------------------------------------
// Range / text addressing. Offsets are UTF-16 code-unit indices into a block's
// text (a paragraph's concatenated run text; a table's `textOf`).

/** The text a block addresses by offset — a paragraph's run text, or `textOf` for
 *  other kinds. */
function blockText(block: Block): string {
  return block.kind === "paragraph" ? textOfRuns(block.runs) : textOf(block);
}

/** The text covered by a selection. A collapsed or single-block selection slices
 *  that block's text (works in any story). A multi-block selection is supported
 *  across TOP-LEVEL BODY blocks — first/last blocks sliced at the caret, whole
 *  blocks between, joined by newlines; it returns "" if an endpoint is not a
 *  top-level body block (a cross-cell/band/note range is out of scope). Endpoints
 *  are ordered automatically, so anchor/focus direction does not matter. */
export function rangeText(doc: Document, selection: DocSelection): string {
  const { anchor, focus } = selection;
  if (anchor.blockId === focus.blockId) {
    const b = getBlockById(doc, anchor.blockId);
    if (!b) return "";
    const lo = Math.min(anchor.offset, focus.offset);
    const hi = Math.max(anchor.offset, focus.offset);
    return blockText(b).slice(lo, hi);
  }
  const index = new Map(doc.blocks.map((b, i) => [b.id, i] as const));
  const ai = index.get(anchor.blockId);
  const fi = index.get(focus.blockId);
  if (ai === undefined || fi === undefined) return "";
  const [start, end, si, ei] = ai <= fi ? [anchor, focus, ai, fi] : [focus, anchor, fi, ai];
  const parts: string[] = [];
  for (let i = si; i <= ei; i++) {
    const full = blockText(doc.blocks[i]!);
    if (i === si) parts.push(full.slice(start.offset));
    else if (i === ei) parts.push(full.slice(0, end.offset));
    else parts.push(full);
  }
  return parts.join("\n");
}

/** The first `DocPosition` where `needle` (a substring) appears in a paragraph's
 *  text, scanning in reading order — the offset side of an edit, so callers can
 *  target `insertText`/`replaceText` without computing offsets by hand. Undefined
 *  when not found. */
export function positionOfText(doc: Document, needle: string, options?: WalkOptions): DocPosition | undefined {
  let found: DocPosition | undefined;
  walk(doc, (b) => {
    if (found !== undefined || b.kind !== "paragraph") return;
    const at = textOfRuns(b.runs).indexOf(needle);
    if (at >= 0) found = { blockId: b.id, offset: at };
  }, options);
  return found;
}
