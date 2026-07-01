// Document query API — a read-only traversal layer over the flat `Document`
// model, the rough analog of .NET's `WordprocessingDocument` descendant queries
// (`Body.Descendants<Paragraph>()`, find-by-text, section enumeration). Pure and
// DOM-free like the rest of the shared core; it never mutates. Page-level queries
// ("what's on page N") are NOT here — pages exist only after layout, so they live
// in the layout package.

import type { Block, Document, ImageBlock, Paragraph, SdtProps, TableBlock } from "./document";
import { BAND_CONTAINERS } from "./document";
import type { Container } from "./ops";
import { fullSdtChain } from "./sdt";
import { textOfRuns } from "./text";
import { resolveSections, type ResolvedSection } from "./sections";

/** Where a visited block sits. `cell`/`note` are set when the block is nested
 *  below the top level of its story. */
export interface BlockContext {
  /** The top-level story this block belongs to (body or a header/footer band).
   *  Note bodies report `"body"` and set `note` instead. */
  container: Container;
  /** Set when the block lives inside a table cell (innermost cell if nested). */
  cell?: { tableId: string; row: number; col: number };
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
): void {
  for (const block of blocks) {
    visit(block, ctx);
    if (opts.cells && block.kind === "table") {
      for (let row = 0; row < block.rows.length; row++) {
        const cells = block.rows[row]!.cells;
        for (let col = 0; col < cells.length; col++) {
          // Spread ctx so a table nested in a note body keeps its `note` membership.
          walkStory(cells[col]!.blocks, { ...ctx, cell: { tableId: block.id, row, col } }, opts, visit);
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
      return block.rows
        .map((row) =>
          row.cells
            .map((cell) => cell.blocks.map(textOf).filter((t) => t.length > 0).join("\n"))
            .join("\t"),
        )
        .join("\n");
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

/** Every content control, as `{ id, props }`, in `doc.sdts` order. */
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
