// Geometry queries over the LayoutTree — consumed by the input layer (selection,
// caret movement) and the wiring layer (selection highlight rects). Read-only.
//
// Internals: a flattened, document-ordered line list is built lazily per tree
// (WeakMap cache), and per-fragment grapheme-cluster advances are measured lazily
// with one cached prefix-measureText pass — most fragments are never hit-tested.

import type { DocPosition, DocSelection } from "@cw/shared";
import type { InlineFragment, LayoutTree, LineBox, PlacedBlock, PlacedTableCell } from "./layoutTree";
import { charStyleToFont } from "./metrics";
import { widthScale } from "../paint/paintStyle";

export interface CaretRect {
  pageIndex: number;
  x: number;
  y: number;
  height: number;
}

export interface Rect {
  pageIndex: number;
  x: number;
  y: number;
  width: number;
  height: number;
}

// ---------------------------------------------------------------------------
// Flattened line list (document order), cached per tree

interface LineEntry {
  pageIndex: number;
  block: PlacedBlock;
  line: LineBox;
  /** Paragraph-text offsets covered by this line (empty line: 0..0). */
  startOffset: number;
  endOffset: number;
  /** The table cell this line lives in, if any — shared by reference across all
   *  of the cell's lines so hit-testing can pin a click to its cell. */
  cell?: PlacedTableCell;
}

interface TreeIndex {
  entries: LineEntry[];
  /** blockId -> indexes into entries, in order. */
  byBlock: Map<string, number[]>;
  /** blockId -> document-order ordinal, for cross-block position comparison. */
  blockOrder: Map<string, number>;
}

/** Story-editing scope: limit every query to ONE page's margin band. The same
 *  band block ids repeat on every page, so an unscoped lookup would be
 *  ambiguous — the scope pins queries to the page being edited. */
export interface GeoScope {
  band: "header" | "footer";
  pageIndex: number;
}

const indexCache = new WeakMap<LayoutTree, TreeIndex>();
const bandIndexCache = new WeakMap<LayoutTree, Map<string, TreeIndex>>();

function buildIndex(
  pageBlocks: Iterable<{ pageIndex: number; blocks: PlacedBlock[] }>,
  indexBandTables: boolean,
): TreeIndex {
  const entries: LineEntry[] = [];
  const byBlock = new Map<string, number[]>();
  const blockOrder = new Map<string, number>();
  const indexBlock = (pageIndex: number, block: PlacedBlock, depth = 0, cell?: PlacedTableCell): void => {
    if (block.table) {
      // Cell paragraphs are regular PlacedBlocks — index them in cell order and
      // every caret/selection query works inside tables with no special cases.
      // (Band tables stay unindexed; NESTED tables render but their inner cells
      // are read-only — the paragraph locator only goes one level deep.)
      if (!indexBandTables || depth > 0) return;
      for (const row of block.table.rows) {
        for (const c of row.cells) for (const cb of c.blocks) indexBlock(pageIndex, cb, depth + 1, c);
      }
      return;
    }
    if (block.image) return; // images carry no lines
    if (!blockOrder.has(block.blockId)) blockOrder.set(block.blockId, blockOrder.size);
    let list = byBlock.get(block.blockId);
    if (!list) {
      list = [];
      byBlock.set(block.blockId, list);
    }
    for (const line of block.lines) {
      // Fragments are stored in VISUAL order, so on a bidi-reordered (RTL/mixed)
      // line fragments[0]/[last] are the visually-, not logically-, first/last —
      // their offsets can run backwards. Scan for the logical min/max so line
      // ownership (entryOf), the caret's line pick, and selection rects stay
      // correct on multi-run RTL lines. (Single-fragment & LTR lines are
      // unaffected: min/max == first/last there.)
      let startOffset = line.emptyOffset ?? 0;
      let endOffset = line.emptyOffset ?? 0;
      if (line.fragments.length > 0) {
        startOffset = Infinity;
        endOffset = -Infinity;
        for (const f of line.fragments) {
          if (f.startOffset < startOffset) startOffset = f.startOffset;
          if (f.endOffset > endOffset) endOffset = f.endOffset;
        }
      }
      list.push(entries.length);
      entries.push({
        pageIndex,
        block,
        line,
        startOffset,
        endOffset,
        ...(cell ? { cell } : {}),
      });
    }
  };
  for (const pb of pageBlocks) {
    for (const block of pb.blocks) indexBlock(pb.pageIndex, block);
  }
  return { entries, byBlock, blockOrder };
}

function treeIndex(tree: LayoutTree): TreeIndex {
  let idx = indexCache.get(tree);
  if (idx) return idx;
  idx = buildIndex(
    tree.pages.map((p) => ({ pageIndex: p.index, blocks: p.blocks })),
    true,
  );
  indexCache.set(tree, idx);
  return idx;
}

function getIndex(tree: LayoutTree, scope?: GeoScope): TreeIndex {
  if (!scope) return treeIndex(tree);
  let perTree = bandIndexCache.get(tree);
  if (!perTree) {
    perTree = new Map();
    bandIndexCache.set(tree, perTree);
  }
  const key = `${scope.band}:${scope.pageIndex}`;
  let idx = perTree.get(key);
  if (!idx) {
    // Band tables ARE indexed (imported footers are routinely a table of text
    // beside a page-number paragraph) — their cell paragraphs edit like body cells.
    const blocks = tree.pages[scope.pageIndex]?.[scope.band] ?? [];
    idx = buildIndex([{ pageIndex: scope.pageIndex, blocks }], true);
    perTree.set(key, idx);
  }
  return idx;
}

export function comparePositions(
  tree: LayoutTree,
  a: DocPosition,
  b: DocPosition,
  scope?: GeoScope,
): number {
  const order = getIndex(tree, scope).blockOrder;
  const oa = order.get(a.blockId) ?? 0;
  const ob = order.get(b.blockId) ?? 0;
  return oa !== ob ? oa - ob : a.offset - b.offset;
}

// ---------------------------------------------------------------------------
// Lazy per-fragment cluster advances

interface ClusterInfo {
  /** Local cluster boundary offsets: [0, ...,] text.length — length n+1. */
  boundaries: number[];
  /** Cumulative advance at each boundary — length n+1, advances[0] === 0. */
  advances: number[];
}

const clusterCache = new WeakMap<InlineFragment, ClusterInfo>();
const graphemeSegmenter = new Intl.Segmenter(undefined, { granularity: "grapheme" });
const advCtx = document.createElement("canvas").getContext("2d")!;

function clustersOf(frag: InlineFragment): ClusterInfo {
  let info = clusterCache.get(frag);
  if (info) return info;
  // Inline equation: one atomic cluster spanning the whole box width, so the
  // caret lands only before/after the equation (never "inside" the sentinel).
  if (frag.equation) {
    info = { boundaries: [0, frag.text.length], advances: [0, frag.width] };
    clusterCache.set(frag, info);
    return info;
  }
  advCtx.font = charStyleToFont(frag.style);
  // letter/wordSpacing are supported in all evergreen canvases; must match paint.
  const ctx2 = advCtx as CanvasRenderingContext2D & { letterSpacing: string; wordSpacing: string };
  ctx2.letterSpacing = `${frag.style.letterSpacingPx ?? 0}px`;
  ctx2.wordSpacing = `${frag.wordSpacingPx ?? 0}px`;
  const boundaries: number[] = [0];
  const advances: number[] = [0];
  // Character width scaling (w:w): the painted glyphs are stretched horizontally,
  // so scale the cluster advances too, keeping caret hit-testing aligned with paint.
  const scale = widthScale(frag.style);
  for (const seg of graphemeSegmenter.segment(frag.text)) {
    const end = seg.index + seg.segment.length;
    boundaries.push(end);
    // Prefix measurement keeps kerning/ligature effects consistent with the
    // single fillText the paint layer issues for the whole fragment.
    advances.push(advCtx.measureText(frag.text.slice(0, end)).width * scale);
  }
  info = { boundaries, advances };
  clusterCache.set(frag, info);
  return info;
}

/** Block-relative x centers of every rendered U+0020 in a fragment, for the
 *  formatting-marks overlay (the space dots). Reuses the cluster-advance cache
 *  (which folds in letter/word spacing), so each dot sits exactly mid-glyph —
 *  consistent with where the caret lands either side of the space. */
export function spaceMarkXs(frag: InlineFragment): number[] {
  const { boundaries, advances } = clustersOf(frag);
  const xs: number[] = [];
  for (let k = 0; k + 1 < boundaries.length; k++) {
    if (frag.text.slice(boundaries[k]!, boundaries[k + 1]!) === " ") {
      xs.push(frag.x + (advances[k]! + advances[k + 1]!) / 2);
    }
  }
  return xs;
}

/** Model-offset delta (relative to startOffset) of a fragment-text-local index.
 *  Identity unless pretext collapsed whitespace inside the fragment. */
const localToModel = (frag: InlineFragment, local: number): number =>
  frag.offsetMap ? (frag.offsetMap[local] ?? frag.offsetMap[frag.offsetMap.length - 1]!) : local;

/** A bidi-reordered line: at least one fragment carries an embedding level (set
 *  by the engine only when it reordered the line). LTR-only lines keep `level`
 *  undefined and take the original, monotonic-x fast path. */
const isBidiLine = (line: LineBox): boolean => line.fragments.some((f) => f.level !== undefined);
const isRtlFrag = (frag: InlineFragment): boolean => ((frag.level ?? 0) & 1) === 1;

/** Block-local x of `offset` WITHIN one fragment, honoring its direction: in an
 *  RTL fragment the logical prefix occupies the RIGHT side, so the boundary after
 *  k logical clusters sits at (frag.x + width − prefixAdvance). */
function xWithinFrag(frag: InlineFragment, offset: number): number {
  const { boundaries, advances } = clustersOf(frag);
  const modelLocal = offset - frag.startOffset;
  let adv = advances[advances.length - 1]!;
  for (let k = 0; k < boundaries.length; k++) {
    if (localToModel(frag, boundaries[k]!) >= modelLocal) {
      adv = advances[k]!;
      break;
    }
  }
  return isRtlFrag(frag) ? frag.x + frag.width - adv : frag.x + adv;
}

/** X of `offset` within a bidi-reordered line. Fragments are in VISUAL order, so
 *  the offset can live anywhere in the array; find its owning fragment, else snap
 *  to the visually-nearest fragment edge. */
function xAtOffsetBidi(entry: LineEntry, offset: number): number {
  const { block, line } = entry;
  let bestX = block.x + line.fragments[0]!.x;
  let bestDist = Infinity;
  for (const frag of line.fragments) {
    if (offset >= frag.startOffset && offset <= frag.endOffset) {
      return block.x + xWithinFrag(frag, offset);
    }
    // Track the nearest boundary for the eaten-gap / out-of-range fallback.
    const d = offset < frag.startOffset ? frag.startOffset - offset : offset - frag.endOffset;
    if (d < bestDist) {
      bestDist = d;
      const edgeOff = offset < frag.startOffset ? frag.startOffset : frag.endOffset;
      bestX = block.x + xWithinFrag(frag, edgeOff);
    }
  }
  return bestX;
}

/** X of `offset` within a line, in page coordinates. Clamps into the line. */
function xAtOffset(entry: LineEntry, offset: number): number {
  const { block, line } = entry;
  if (line.fragments.length === 0) return block.x;
  if (isBidiLine(line)) return xAtOffsetBidi(entry, offset);
  const first = line.fragments[0]!;
  if (offset <= first.startOffset) return block.x + first.x;
  for (const frag of line.fragments) {
    if (offset > frag.endOffset) continue;
    if (offset < frag.startOffset) return block.x + frag.x; // offset in an eaten gap
    const { boundaries, advances } = clustersOf(frag);
    const modelLocal = offset - frag.startOffset;
    for (let k = 0; k < boundaries.length; k++) {
      if (localToModel(frag, boundaries[k]!) >= modelLocal) return block.x + frag.x + advances[k]!;
    }
    return block.x + frag.x + advances[advances.length - 1]!;
  }
  const lastFrag = line.fragments[line.fragments.length - 1]!;
  const { advances } = clustersOf(lastFrag);
  return block.x + lastFrag.x + advances[advances.length - 1]!;
}

/** Cluster-boundary offset nearest page-x within ONE fragment, honoring its
 *  direction (RTL measures the cursor distance from the fragment's right edge). */
function offsetWithinFrag(frag: InlineFragment, fx: number): number {
  const { boundaries, advances } = clustersOf(frag);
  const rtl = isRtlFrag(frag);
  const target = rtl ? frag.width - fx : fx; // distance along the logical reading run
  for (let k = 0; k + 1 < advances.length; k++) {
    if (target <= advances[k + 1]!) {
      const mid = (advances[k]! + advances[k + 1]!) / 2;
      const boundary = target < mid ? boundaries[k]! : boundaries[k + 1]!;
      return frag.startOffset + localToModel(frag, boundary);
    }
  }
  return frag.endOffset;
}

/** Hit-test x within a bidi-reordered line. Fragments are in visual (L→R) order,
 *  so walk them by x; map within the owning fragment honoring its direction. */
function offsetAtXBidi(entry: LineEntry, x: number): number {
  const { block, line } = entry;
  const local = x - block.x;
  const frags = line.fragments;
  for (let fi = 0; fi < frags.length; fi++) {
    const frag = frags[fi]!;
    const next = frags[fi + 1];
    const fragEnd = frag.x + frag.width;
    if (local > fragEnd) {
      if (next && local >= next.x) continue;
      // In the visual gap after this fragment: snap to the nearer edge. The
      // logical offset at a fragment's visual-right edge depends on direction.
      const here = isRtlFrag(frag) ? frag.startOffset : frag.endOffset;
      if (next) {
        const mid = (fragEnd + next.x) / 2;
        const there = isRtlFrag(next) ? next.endOffset : next.startOffset;
        return local < mid ? here : there;
      }
      return here;
    }
    if (local <= frag.x) {
      // Before the first/this fragment visually: its visual-left edge offset.
      return isRtlFrag(frag) ? frag.endOffset : frag.startOffset;
    }
    return offsetWithinFrag(frag, local - frag.x);
  }
  return entry.endOffset;
}

/** Nearest cluster-boundary offset at page-x within a line (rounds to midpoint). */
function offsetAtX(entry: LineEntry, x: number): number {
  const { block, line } = entry;
  if (line.fragments.length === 0) return entry.startOffset;
  if (isBidiLine(line)) return offsetAtXBidi(entry, x);
  const local = x - block.x;
  const first = line.fragments[0]!;
  if (local <= first.x) return first.startOffset;
  for (let fi = 0; fi < line.fragments.length; fi++) {
    const frag = line.fragments[fi]!;
    const next = line.fragments[fi + 1];
    const fragEnd = frag.x + frag.width;
    if (local > fragEnd) {
      if (next && local >= next.x) continue;
      // In the gap after this fragment (or past the last one): nearest edge.
      if (next) {
        const mid = (fragEnd + next.x) / 2;
        return local < mid ? frag.endOffset : next.startOffset;
      }
      return frag.endOffset;
    }
    const { boundaries, advances } = clustersOf(frag);
    const fx = local - frag.x;
    for (let k = 0; k + 1 < advances.length; k++) {
      if (fx <= advances[k + 1]!) {
        const mid = (advances[k]! + advances[k + 1]!) / 2;
        const boundary = fx < mid ? boundaries[k]! : boundaries[k + 1]!;
        return frag.startOffset + localToModel(frag, boundary);
      }
    }
    return frag.endOffset;
  }
  return entry.endOffset;
}

/** The line entry a position renders on (offsets in eaten line-break whitespace
 *  get previous-line affinity, matching where the text visually ended). */
function entryOf(idx: TreeIndex, pos: DocPosition): LineEntry | null {
  const lineIdxs = idx.byBlock.get(pos.blockId);
  if (!lineIdxs || lineIdxs.length === 0) return null;
  let chosen = idx.entries[lineIdxs[0]!]!;
  for (const li of lineIdxs) {
    const e = idx.entries[li]!;
    if (e.startOffset <= pos.offset) chosen = e;
    else break;
  }
  return chosen;
}

function entryIndexOf(idx: TreeIndex, pos: DocPosition): number {
  const lineIdxs = idx.byBlock.get(pos.blockId);
  if (!lineIdxs || lineIdxs.length === 0) return -1;
  let chosen = lineIdxs[0]!;
  for (const li of lineIdxs) {
    if (idx.entries[li]!.startOffset <= pos.offset) chosen = li;
    else break;
  }
  return chosen;
}

// ---------------------------------------------------------------------------
// Public queries

/** Point on a page -> nearest document position. Vertically-containing lines
 *  are ranked by horizontal distance — table cells put several lines at the
 *  same y, so pure line-by-y would always land in the leftmost column. */
export function hitTest(
  tree: LayoutTree,
  pageIndex: number,
  x: number,
  y: number,
  scope?: GeoScope,
): DocPosition | null {
  const idx = getIndex(tree, scope);

  // If the point lands inside a table cell, the caret belongs in THAT cell —
  // never a neighbour. A short line in the clicked cell can sit farther from x
  // than a longer line in the next column, so the x-distance ranking below
  // would otherwise jump columns. Pin candidates to the containing cell.
  let targetCell: PlacedTableCell | undefined;
  for (const e of idx.entries) {
    if (e.pageIndex !== pageIndex || !e.cell) continue;
    const c = e.cell;
    if (x >= c.x && x < c.x + c.width && y >= c.y && y < c.y + c.height) {
      targetCell = c;
      break;
    }
  }

  const xDistance = (e: LineEntry): number => {
    const first = e.line.fragments[0];
    const last = e.line.fragments[e.line.fragments.length - 1];
    const left = e.block.x + (first ? first.x : 0);
    const right = first && last ? e.block.x + last.x + last.width : left + 1;
    return x < left ? left - x : x > right ? x - right : 0;
  };

  let containing: LineEntry | null = null; // best line whose vertical span holds y
  let before: LineEntry | null = null; // nearest line fully above y on this page
  let after: LineEntry | null = null; // first line below y on this page
  let earlier: LineEntry | null = null; // last line on an earlier page

  for (const e of idx.entries) {
    if (e.pageIndex > pageIndex) break;
    if (e.pageIndex < pageIndex) {
      earlier = e;
      continue;
    }
    if (targetCell && e.cell !== targetCell) continue; // pinned to the clicked cell
    const top = e.block.y + e.line.y;
    const bottom = top + e.line.height;
    if (y >= top && y < bottom) {
      if (containing === null || xDistance(e) < xDistance(containing)) containing = e;
    } else if (bottom <= y) {
      if (before === null || top >= before.block.y + before.line.y) before = e;
    } else if (after === null) {
      after = e;
    }
  }

  const best = containing ?? before ?? after ?? earlier ?? idx.entries[0] ?? null;
  if (!best) return null;
  return { blockId: best.block.blockId, offset: offsetAtX(best, x) };
}

export function caretRect(tree: LayoutTree, pos: DocPosition, scope?: GeoScope): CaretRect | null {
  const e = entryOf(getIndex(tree, scope), pos);
  if (!e) return null;
  return {
    pageIndex: e.pageIndex,
    x: xAtOffset(e, Math.min(pos.offset, Math.max(e.endOffset, e.startOffset))),
    y: e.block.y + e.line.y,
    height: e.line.height,
  };
}

/** Per-line highlight rectangles; handles multi-block and multi-page selections. */
export function selectionRects(tree: LayoutTree, sel: DocSelection, scope?: GeoScope): Rect[] {
  const idx = getIndex(tree, scope);
  const [from, to] =
    comparePositions(tree, sel.anchor, sel.focus, scope) <= 0
      ? [sel.anchor, sel.focus]
      : [sel.focus, sel.anchor];
  const fromIdx = entryIndexOf(idx, from);
  const toIdx = entryIndexOf(idx, to);
  if (fromIdx < 0 || toIdx < 0) return [];

  const rects: Rect[] = [];
  for (let i = fromIdx; i <= toIdx; i++) {
    const e = idx.entries[i]!;
    const y = e.block.y + e.line.y;
    const lo = i === fromIdx ? from.offset : e.startOffset;
    const hi = i === toIdx ? to.offset : e.endOffset;
    if (isBidiLine(e.line)) {
      // A logical range can be visually discontiguous: emit one rect per
      // fragment-portion it covers (each fragment is single-direction).
      for (const frag of e.line.fragments) {
        const a = Math.max(lo, frag.startOffset);
        const b = Math.min(hi, frag.endOffset);
        if (b <= a) continue;
        const xa = e.block.x + xWithinFrag(frag, a);
        const xb = e.block.x + xWithinFrag(frag, b);
        const x1 = Math.min(xa, xb);
        const x2 = Math.max(xa, xb);
        if (x2 <= x1) continue;
        rects.push({ pageIndex: e.pageIndex, x: x1, y, width: x2 - x1, height: e.line.height });
      }
      continue;
    }
    const x1 = i === fromIdx ? xAtOffset(e, from.offset) : xAtOffset(e, e.startOffset);
    // Interior lines extend a touch past the text end (Word marks the line break).
    const x2 = i === toIdx ? xAtOffset(e, to.offset) : xAtOffset(e, e.endOffset) + 4;
    if (x2 <= x1) continue;
    rects.push({
      pageIndex: e.pageIndex,
      x: x1,
      y,
      width: x2 - x1,
      height: e.line.height,
    });
  }
  return rects;
}

/** Next caret stop VISUALLY left (-1) or right (+1) of `pos` within its own line,
 *  for bidi-aware arrow movement. Returns null on a non-bidi line (the caller uses
 *  the cheap logical step) or when `pos` is already at the line's visual edge (the
 *  caller crosses to the adjacent line). Walks every cluster boundary on the line
 *  and picks the nearest stop strictly beyond the caret's current x. */
export function visualCaretStep(
  tree: LayoutTree,
  pos: DocPosition,
  dir: -1 | 1,
  scope?: GeoScope,
): DocPosition | null {
  const e = entryOf(getIndex(tree, scope), pos);
  if (!e || !isBidiLine(e.line)) return null;
  const curX = xAtOffset(e, pos.offset);
  let bestOff = -1;
  let bestX = dir === 1 ? Infinity : -Infinity;
  const consider = (offset: number, x: number): void => {
    if (dir === 1 ? x > curX + 0.01 && x < bestX : x < curX - 0.01 && x > bestX) {
      bestX = x;
      bestOff = offset;
    }
  };
  for (const frag of e.line.fragments) {
    const { boundaries } = clustersOf(frag);
    for (const b of boundaries) {
      const offset = frag.startOffset + localToModel(frag, b);
      consider(offset, xAtOffset(e, offset));
    }
  }
  return bestOff < 0 ? null : { blockId: e.block.blockId, offset: bestOff };
}

/** Home/End need line geometry, not just the model. */
export function lineEdges(
  tree: LayoutTree,
  pos: DocPosition,
  scope?: GeoScope,
): { home: DocPosition; end: DocPosition } | null {
  const e = entryOf(getIndex(tree, scope), pos);
  if (!e) return null;
  return {
    home: { blockId: e.block.blockId, offset: e.startOffset },
    end: { blockId: e.block.blockId, offset: e.endOffset },
  };
}

export function positionOnAdjacentLine(
  tree: LayoutTree,
  pos: DocPosition,
  direction: "up" | "down",
  goalX: number,
  scope?: GeoScope,
): DocPosition | null {
  const idx = getIndex(tree, scope);
  const i = entryIndexOf(idx, pos);
  if (i < 0) return null;
  const j = direction === "up" ? i - 1 : i + 1;
  const e = idx.entries[j];
  if (!e) return null;
  return { blockId: e.block.blockId, offset: offsetAtX(e, goalX) };
}

// ---------------------------------------------------------------------------
// Object (image) and table-structure hit-testing

export interface ObjectHit {
  blockId: string;
  rect: Rect;
}

function scanImages(
  blocks: PlacedBlock[],
  visit: (b: PlacedBlock) => boolean, // return true to stop
): boolean {
  for (const block of blocks) {
    if (block.image && visit(block)) return true;
    if (block.table) {
      for (const row of block.table.rows) {
        for (const cell of row.cells) if (scanImages(cell.blocks, visit)) return true;
      }
    }
  }
  return false;
}

/** Topmost image under a page point matching `want` — including images inside
 *  table cells. page.blocks is in paint order, so keeping the LAST match yields
 *  the visually-topmost image. (Band objects stay read-only.) */
function hitImage(
  tree: LayoutTree,
  pageIndex: number,
  x: number,
  y: number,
  want: (im: { behind?: boolean }) => boolean,
): ObjectHit | null {
  const page = tree.pages[pageIndex];
  if (!page) return null;
  let hit: ObjectHit | null = null;
  scanImages(page.blocks, (block) => {
    const im = block.image!;
    if (!want(im)) return false;
    const { width, height } = im;
    if (x >= block.x && x <= block.x + width && y >= block.y && y <= block.y + height) {
      hit = { blockId: block.blockId, rect: { pageIndex, x: block.x, y: block.y, width, height } };
    }
    return false;
  });
  return hit;
}

/** Foreground image (inline or in-front-of-text anchor) under a page point.
 *  Behind-text backgrounds are excluded — those select only via the gaps (see
 *  hitTestSelectableObject), so clicks on the text in front fall through. */
export function hitTestObject(tree: LayoutTree, pageIndex: number, x: number, y: number): ObjectHit | null {
  return hitImage(tree, pageIndex, x, y, (im) => im.behind !== true);
}

/** Behind-text image under a page point (full-page backgrounds). */
export function hitTestBehindObject(tree: LayoutTree, pageIndex: number, x: number, y: number): ObjectHit | null {
  return hitImage(tree, pageIndex, x, y, (im) => im.behind === true);
}

/** True when (x,y) lands on the painted glyphs of some text line on the page.
 *  Lets a behind-text image stay selectable in its empty regions without ever
 *  stealing a click meant for the text drawn on top of it. */
export function pointOnText(tree: LayoutTree, pageIndex: number, x: number, y: number, scope?: GeoScope): boolean {
  const idx = getIndex(tree, scope);
  for (const e of idx.entries) {
    if (e.pageIndex !== pageIndex) continue;
    const top = e.block.y + e.line.y;
    const bottom = top + e.line.height;
    if (y < top || y >= bottom) continue;
    const first = e.line.fragments[0];
    const last = e.line.fragments[e.line.fragments.length - 1];
    if (!first || !last) continue; // blank line — no glyphs, let the click pass through
    const left = e.block.x + first.x;
    const right = e.block.x + last.x + last.width;
    if (x >= left && x <= right) return true;
  }
  return false;
}

/** Display-equation block under a page point (incl. equations inside table cells).
 *  Equations aren't object-selectable like images (no resize frame); this is used
 *  by the context menu to offer "Edit Equation…". Returns the topmost match. */
export function hitTestEquation(tree: LayoutTree, pageIndex: number, x: number, y: number): ObjectHit | null {
  const page = tree.pages[pageIndex];
  if (!page) return null;
  let hit: ObjectHit | null = null;
  const scan = (blocks: PlacedBlock[]): void => {
    for (const block of blocks) {
      if (block.equation) {
        const { width, height } = block.equation;
        if (x >= block.x && x <= block.x + width && y >= block.y && y <= block.y + height) {
          hit = { blockId: block.blockId, rect: { pageIndex, x: block.x, y: block.y, width, height } };
        }
      }
      if (block.table) for (const row of block.table.rows) for (const cell of row.cells) scan(cell.blocks);
    }
  };
  scan(page.blocks);
  return hit;
}

/** Object to select for a click/right-click at a page point: a foreground image
 *  wins anywhere it's hit; a behind-text image is selectable only where no text
 *  covers it (so text on top stays clickable). */
export function hitTestSelectableObject(
  tree: LayoutTree,
  pageIndex: number,
  x: number,
  y: number,
  scope?: GeoScope,
): ObjectHit | null {
  const fg = hitTestObject(tree, pageIndex, x, y);
  if (fg) return fg;
  const eq = hitTestEquation(tree, pageIndex, x, y);
  if (eq) return eq;
  if (pointOnText(tree, pageIndex, x, y, scope)) return null;
  return hitTestBehindObject(tree, pageIndex, x, y);
}

/** Where the selection frame for an image OR equation block lives right now
 *  (post-relayout) — drives the object frame for both. */
export function objectRect(tree: LayoutTree, blockId: string): Rect | null {
  for (const page of tree.pages) {
    let rect: Rect | null = null;
    scanImages(page.blocks, (block) => {
      if (block.blockId !== blockId) return false;
      rect = {
        pageIndex: page.index,
        x: block.x,
        y: block.y,
        width: block.image!.width,
        height: block.image!.height,
      };
      return true;
    });
    if (rect) return rect;
  }
  // Equations aren't image blocks; scan them separately (incl. table cells).
  for (const page of tree.pages) {
    const found = scanEquationRect(page.blocks, blockId, page.index);
    if (found) return found;
  }
  return null;
}

function scanEquationRect(blocks: PlacedBlock[], blockId: string, pageIndex: number): Rect | null {
  for (const block of blocks) {
    if (block.equation && block.blockId === blockId) {
      return { pageIndex, x: block.x, y: block.y, width: block.equation.width, height: block.equation.height };
    }
    if (block.table) {
      for (const row of block.table.rows) {
        for (const cell of row.cells) {
          const r = scanEquationRect(cell.blocks, blockId, pageIndex);
          if (r) return r;
        }
      }
    }
  }
  return null;
}

/** Inline equation under a page point: the character RANGE of the equation run
 *  (one U+FFFC), so a click can select it like a small inline object. */
export function inlineEquationAt(
  tree: LayoutTree,
  pageIndex: number,
  x: number,
  y: number,
  scope?: GeoScope,
): { blockId: string; start: number; end: number } | null {
  const idx = getIndex(tree, scope);
  for (const e of idx.entries) {
    if (e.pageIndex !== pageIndex) continue;
    const top = e.block.y + e.line.y;
    if (y < top || y >= top + e.line.height) continue;
    for (const frag of e.line.fragments) {
      if (!frag.equation) continue;
      const fx = e.block.x + frag.x;
      if (x >= fx && x <= fx + frag.width) {
        return { blockId: frag.blockId, start: frag.startOffset, end: frag.endOffset };
      }
    }
  }
  return null;
}

/** Hyperlink under a page point, if the fragment there carries one. */
export function linkAt(
  tree: LayoutTree,
  pageIndex: number,
  x: number,
  y: number,
  scope?: GeoScope,
): string | null {
  const idx = getIndex(tree, scope);
  for (const e of idx.entries) {
    if (e.pageIndex !== pageIndex) continue;
    const top = e.block.y + e.line.y;
    if (y < top || y >= top + e.line.height) continue;
    for (const frag of e.line.fragments) {
      const fx = e.block.x + frag.x;
      if (x >= fx && x <= fx + frag.width) return frag.style.link ?? null;
    }
  }
  return null;
}

export interface ColumnBoundaryHit {
  tableId: string;
  /** Boundary between columns boundaryIndex and boundaryIndex+1. */
  boundaryIndex: number;
  x: number;
  tableX: number;
  tableWidth: number;
  /** This table chunk's vertical extent on the hit page — drives the hover guide. */
  pageIndex: number;
  tableY: number;
  tableHeight: number;
  /** The table's CURRENTLY RENDERED per-column widths — lets a resize that exits
   *  autofit pin the on-screen proportions instead of a stale fractions snapshot. */
  colWidths: number[];
}

const COLUMN_GRIP_PX = 6;

/** Interior column boundary of a body table within grip distance of a point.
 *  Boundaries come from colWidths, not row cells — merged cells make row-cell
 *  edges unreliable as column markers. */
export function hitTestColumnBoundary(
  tree: LayoutTree,
  pageIndex: number,
  x: number,
  y: number,
): ColumnBoundaryHit | null {
  const page = tree.pages[pageIndex];
  if (!page) return null;
  for (const block of page.blocks) {
    const t = block.table;
    if (!t) continue;
    if (y < t.y || y > t.y + t.height) continue;
    let edge = t.x;
    for (let ci = 0; ci + 1 < t.colWidths.length; ci++) {
      edge += t.colWidths[ci]!;
      if (Math.abs(x - edge) <= COLUMN_GRIP_PX) {
        return {
          tableId: block.blockId, boundaryIndex: ci, x: edge, tableX: t.x, tableWidth: t.width,
          pageIndex, tableY: t.y, tableHeight: t.height, colWidths: t.colWidths,
        };
      }
    }
  }
  return null;
}

export interface RowBoundaryHit {
  tableId: string;
  /** MODEL index of the row whose bottom edge is being dragged. */
  rowIndex: number;
  /** Boundary y (page-local doc coords) and the row's top edge — drag delta is
   *  measured against the rendered height (y − rowTop). */
  y: number;
  rowTop: number;
  /** This table chunk's horizontal extent on the hit page — drives the hover guide. */
  pageIndex: number;
  tableX: number;
  tableWidth: number;
}

const ROW_GRIP_PX = 6;

/** Horizontal boundary (a row's bottom edge) of a body table within grip distance
 *  of a point. Mirrors {@link hitTestColumnBoundary}; the caller checks columns
 *  FIRST so a cell-corner pixel prefers the column grip. Every row's bottom edge is
 *  grabbable (including the last — unlike columns, a row drag sets that one row's
 *  height rather than redistributing). The returned rowIndex is the MODEL index
 *  (chunk offset + placed row), so a continuation page / repeated header maps back
 *  to the right model row. */
export function hitTestRowBoundary(
  tree: LayoutTree,
  pageIndex: number,
  x: number,
  y: number,
): RowBoundaryHit | null {
  const page = tree.pages[pageIndex];
  if (!page) return null;
  for (const block of page.blocks) {
    const t = block.table;
    if (!t) continue;
    if (x < t.x || x > t.x + t.width) continue;
    for (let ri = 0; ri < t.rows.length; ri++) {
      const row = t.rows[ri]!;
      const edge = row.y + row.height;
      if (Math.abs(y - edge) <= ROW_GRIP_PX) {
        return {
          tableId: block.blockId,
          rowIndex: block.firstLineIndex + ri,
          y: edge,
          rowTop: row.y,
          pageIndex,
          tableX: t.x,
          tableWidth: t.width,
        };
      }
    }
  }
  return null;
}

/** A point's table cell, in GRID coordinates (body tables only). Returns the
 *  ORIGIN (top-left grid slot) of the merged cell the point lands in, so a point
 *  anywhere inside a colSpan/rowSpan cell maps to one stable coordinate — a drag
 *  across a spanning cell (e.g. a full-width header) stays in that single cell
 *  instead of flipping columns and tripping a false cross-cell selection. */
export interface CellHit {
  tableId: string;
  row: number;
  col: number;
}

export function hitTestCell(tree: LayoutTree, pageIndex: number, x: number, y: number): CellHit | null {
  const page = tree.pages[pageIndex];
  if (!page) return null;
  for (const block of page.blocks) {
    const t = block.table;
    if (!t) continue;
    if (x < t.x || x > t.x + t.width || y < t.y || y > t.y + t.height) continue;
    // Placed cell boxes tile the table (a merged cell covers its whole
    // rectangle, including rowSpan holes), so the box containing the point is
    // the merged cell itself — its origin is the stable grid coordinate.
    for (const row of t.rows) {
      for (const cell of row.cells) {
        if (x >= cell.x && x < cell.x + cell.width && y >= cell.y && y < cell.y + cell.height) {
          return { tableId: block.blockId, row: cell.originRow, col: cell.originCol };
        }
      }
    }
  }
  return null;
}

/** Pixel rectangles covering a NORMALIZED grid rectangle [r0..r1]×[c0..c1] of a
 *  table (one per page chunk it spans). The caller normalizes against the model
 *  grid first so the rectangle aligns to whole merged cells. */
export function cellRangeRects(
  tree: LayoutTree,
  tableId: string,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): Rect[] {
  const rects: Rect[] = [];
  for (let pageIndex = 0; pageIndex < tree.pages.length; pageIndex++) {
    for (const block of tree.pages[pageIndex]!.blocks) {
      const t = block.table;
      if (!t || block.blockId !== tableId) continue;
      const chunkFirst = block.firstLineIndex;
      const chunkLast = chunkFirst + t.rows.length - 1;
      const lo = Math.max(r0, chunkFirst);
      const hi = Math.min(r1, chunkLast);
      if (lo > hi) continue;
      const topRow = t.rows[lo - chunkFirst]!;
      const botRow = t.rows[hi - chunkFirst]!;
      let xLeft = t.x;
      for (let i = 0; i < c0 && i < t.colWidths.length; i++) xLeft += t.colWidths[i]!;
      let width = 0;
      for (let i = c0; i <= c1 && i < t.colWidths.length; i++) width += t.colWidths[i]!;
      rects.push({ pageIndex, x: xLeft, y: topRow.y, width, height: botRow.y + botRow.height - topRow.y });
    }
  }
  return rects;
}

/** First/last positions in the document — or in the scoped band story. */
export function documentEdges(tree: LayoutTree, scope?: GeoScope): { start: DocPosition; end: DocPosition } | null {
  const idx = getIndex(tree, scope);
  const first = idx.entries[0];
  const last = idx.entries[idx.entries.length - 1];
  if (!first || !last) return null;
  return {
    start: { blockId: first.block.blockId, offset: first.startOffset },
    end: { blockId: last.block.blockId, offset: last.endOffset },
  };
}
