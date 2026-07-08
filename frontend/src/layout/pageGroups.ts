// Page-group model — the safe unit for "reorder pages". A PAGE GROUP is a
// contiguous range of top-level body blocks delimited by a HARD structural
// break: a group starts at block 0, at any paragraph carrying pageBreakBefore,
// or at a section start (the previous block is a paragraph carrying
// sectionBreak). These are exactly the boundaries the layout engine turns into
// a forced page break (see engine.ts), so a group seam always coincides with a
// page boundary and reordering only ever re-sequences WHOLE blocks — it never
// splits a paragraph/table or rewrites content.
//
// A multi-page chapter with no internal hard break is ONE group (its pages move
// together); a report with a page/section break between each logical page has
// one group per page (the common case → behaves like a PDF's page thumbnails).
//
// Everything here is pure model logic (no DOM, no layout) except the optional
// page-index mapping, which reads a LayoutTree only to tell the UI which
// rendered pages belong to each group.

import type { Block, Document, Op, Paragraph } from "@cw/shared";
import type { LayoutTree } from "./layoutTree";

export interface PageGroup {
  /** 0-based position of this group in the document's group list. */
  index: number;
  /** Inclusive first/last top-level block indices into doc.blocks. */
  startBlock: number;
  endBlock: number;
  /** Block ids in this group, in order. */
  blockIds: string[];
  /** LayoutTree page indices this group occupies (in order). Empty unless a
   *  tree was passed to computePageGroups. */
  pageIndices: number[];
}

const isPara = (b: Block | undefined): b is Paragraph => !!b && b.kind === "paragraph";
const hasSectionBreak = (b: Block | undefined): boolean => isPara(b) && !!b.style.sectionBreak;
const hasPageBreakBefore = (b: Block | undefined): boolean => isPara(b) && b.style.pageBreakBefore === true;

/** The top-level block indices at which a new page group begins. */
export function groupStartIndices(doc: Document): number[] {
  const starts: number[] = [];
  const blocks = doc.blocks;
  for (let i = 0; i < blocks.length; i++) {
    if (i === 0 || hasPageBreakBefore(blocks[i]) || hasSectionBreak(blocks[i - 1])) starts.push(i);
  }
  return starts;
}

/** Split the document into page groups — the movable units for reordering. Pass
 *  the current LayoutTree to (a) fill each unit's pageIndices and (b) MERGE model
 *  groups that don't actually begin a fresh page.
 *
 *  The model's hard-break boundaries (pageBreakBefore / section start) USUALLY
 *  coincide with page boundaries, but not always: a section break followed by an
 *  effectively-hidden paragraph is layout-skipped, so the "next" group renders on
 *  the SAME page as the previous one (e.g. an Executive Summary section whose
 *  tables spill onto the following section's first page). Those groups are one
 *  visual page and must move together — otherwise dragging one leaves the other's
 *  content behind. So with a tree, a group that isn't the first-placed block of
 *  any page is merged into the unit before it. */
export function computePageGroups(doc: Document, tree?: LayoutTree): PageGroup[] {
  const starts = groupStartIndices(doc);
  const groups: PageGroup[] = [];
  for (let g = 0; g < starts.length; g++) {
    const startBlock = starts[g]!;
    const endBlock = (g + 1 < starts.length ? starts[g + 1]! : doc.blocks.length) - 1;
    groups.push({
      index: g,
      startBlock,
      endBlock,
      blockIds: doc.blocks.slice(startBlock, endBlock + 1).map((b) => b.id),
      pageIndices: [],
    });
  }
  if (!tree) return groups;
  const units = mergeToVisualPages(tree, groups);
  assignPages(tree, units);
  return units;
}

/** Merge each model group that does not begin a fresh page into the unit before
 *  it, so a movable unit is exactly what renders together as a page-run. */
function mergeToVisualPages(tree: LayoutTree, groups: PageGroup[]): PageGroup[] {
  const pageFirst = new Set<string>();
  for (const page of tree.pages) {
    const id = page.blocks[0]?.blockId;
    if (id !== undefined) pageFirst.add(id);
  }
  const units: PageGroup[] = [];
  for (const g of groups) {
    const prev = units[units.length - 1];
    if (prev && !pageFirst.has(g.blockIds[0]!)) {
      prev.endBlock = g.endBlock;
      prev.blockIds.push(...g.blockIds);
    } else {
      units.push({ index: units.length, startBlock: g.startBlock, endBlock: g.endBlock, blockIds: [...g.blockIds], pageIndices: [] });
    }
  }
  return units;
}

/** Attach each rendered page to the unit that owns its first placed body block.
 *  After merging, every page's first block belongs to exactly one unit. */
function assignPages(tree: LayoutTree, groups: PageGroup[]): void {
  const groupOf = new Map<string, number>();
  groups.forEach((g, gi) => g.blockIds.forEach((id) => groupOf.set(id, gi)));
  for (const page of tree.pages) {
    const firstId = page.blocks[0]?.blockId;
    const gi = firstId !== undefined ? groupOf.get(firstId) : undefined;
    if (gi !== undefined) groups[gi]!.pageIndices.push(page.index);
  }
}

/** Seam-normalization ops for a reordered body (given as `ids` in final order,
 *  with `requiredStartIds` the block ids that MUST begin a fresh page — every
 *  group start plus the moved content's first block). For each required start
 *  except the one at document position 0, guarantee a leading hard break: a
 *  section break on the preceding block already does it, otherwise pin
 *  pageBreakBefore on the (paragraph) first block. Redundant breaks are inert
 *  (the engine no-ops pageBreakBefore at page top), and the very first block's
 *  now-inert break is stripped so a reordered doc exports clean. We never touch
 *  section-break paragraphs — the reorder leaves those in place — so no other
 *  section's geometry moves.
 *
 *  This is a MODEL-only heuristic. It cannot see a break governed by an
 *  effectively-hidden first paragraph that layout-skips and swallows; that
 *  surfaces as a group rendering zero pages, which the "Organize Pages" UI
 *  repairs with a layout-verified pass (pinPageBreakBeforeCmd). */
function seamBreakOps(byId: Map<string, Block>, ids: string[], requiredStartIds: Set<string>): Op[] {
  const posOf = new Map(ids.map((id, i) => [id, i] as const));
  const pin = new Set<string>();
  const strip = new Set<string>();
  for (const id of requiredStartIds) {
    const i = posOf.get(id);
    if (i === undefined) continue;
    const b = byId.get(id);
    if (i === 0) {
      if (isPara(b) && b.style.pageBreakBefore === true) strip.add(id);
      continue;
    }
    if (hasSectionBreak(byId.get(ids[i - 1]!))) continue; // section break already starts a page
    if (isPara(b) && b.style.pageBreakBefore === true) continue; // already breaks
    if (isPara(b)) pin.add(id);
  }
  const ops: Op[] = [];
  for (const id of pin) ops.push({ type: "setParaStyle", blockId: id, patch: { pageBreakBefore: true } });
  for (const id of strip) if (!pin.has(id)) ops.push({ type: "setParaStyle", blockId: id, patch: { pageBreakBefore: false } });
  return ops;
}

/** Ops to pin an explicit pageBreakBefore on a paragraph (the layout-verified
 *  repair for a group that merged onto the previous page after a reorder —
 *  usually a section-chain disruption or a skipped hidden first paragraph).
 *  No-op if the block isn't a paragraph or already carries the break. */
export function pinPageBreakBeforeOps(doc: Document, blockId: string): Op[] {
  const b = doc.blocks.find((x) => x.id === blockId);
  if (!isPara(b) || b.style.pageBreakBefore === true) return [];
  return [{ type: "setParaStyle", blockId, patch: { pageBreakBefore: true } }];
}

/** Ops that move the contiguous block range [firstId…lastId] (a movable unit) so
 *  it lands immediately before `anchorId` (or at the document end when null).
 *  Emits a remove+insert batch plus seam-normalization breaks. Returns [] for an
 *  invalid range, a drop into itself, or a no-op (already in that position).
 *
 *  Section-boundary safety: if the range ends in a section-break paragraph, that
 *  terminator is LEFT in place — moving it would relocate the section boundary
 *  and reflow OTHER sections (e.g. split a Table of Contents that isn't being
 *  moved). Only the content before it travels and adopts the destination's
 *  section geometry. */
export function reorderPageGroupsOps(doc: Document, firstId: string, lastId: string, anchorId: string | null): Op[] {
  const blocks = doc.blocks;
  const fromI = blocks.findIndex((b) => b.id === firstId);
  const toI = blocks.findIndex((b) => b.id === lastId);
  if (fromI < 0 || toI < fromI) return [];

  const byId = new Map<string, Block>(blocks.map((b) => [b.id, b] as const));
  let movedIds = blocks.slice(fromI, toI + 1).map((b) => b.id);
  const last = byId.get(movedIds[movedIds.length - 1]!);
  if (isPara(last) && !!last.style.sectionBreak) movedIds = movedIds.slice(0, -1);
  if (movedIds.length === 0) return [];

  const movedSet = new Set(movedIds);
  if (anchorId !== null && movedSet.has(anchorId)) return []; // dropped into itself
  const afterRemoval = blocks.filter((b) => !movedSet.has(b.id)).map((b) => b.id);
  const insertPos = anchorId === null ? afterRemoval.length : afterRemoval.indexOf(anchorId);
  if (insertPos < 0) return [];

  const resultIds = afterRemoval.slice();
  resultIds.splice(insertPos, 0, ...movedIds);
  if (resultIds.length === blocks.length && resultIds.every((id, i) => id === blocks[i]!.id)) return []; // unchanged

  const ops: Op[] = [];
  for (const id of movedIds) ops.push({ type: "removeBlock", blockId: id });
  movedIds.forEach((id, k) => ops.push({ type: "insertBlock", index: insertPos + k, block: byId.get(id)!, where: "body" }));

  // Boundaries that must still begin a fresh page: the ORIGINAL group starts (so
  // a page stays distinct even if it loses its section-break predecessor), the
  // group starts of the reordered body, and the moved content's first block.
  const required = new Set<string>(groupStartIndices(doc).map((i) => blocks[i]!.id));
  for (let i = 0; i < resultIds.length; i++) {
    const b = byId.get(resultIds[i]!);
    const prev = i > 0 ? byId.get(resultIds[i - 1]!) : undefined;
    if (i === 0 || hasPageBreakBefore(b) || hasSectionBreak(prev)) required.add(resultIds[i]!);
  }
  required.add(movedIds[0]!);
  ops.push(...seamBreakOps(byId, resultIds, required));
  return ops;
}

/** Ops that delete the contiguous block range [firstId…lastId] (a movable unit).
 *  Returns [] when that would empty the document. Strips a trailing section break
 *  the deletion would leave as the document's last block (which would show as an
 *  empty section), and a now-inert leading break when the first block is removed. */
export function deletePageGroupOps(doc: Document, firstId: string, lastId: string): Op[] {
  const blocks = doc.blocks;
  const fromI = blocks.findIndex((b) => b.id === firstId);
  const toI = blocks.findIndex((b) => b.id === lastId);
  if (fromI < 0 || toI < fromI) return [];
  const movedIds = blocks.slice(fromI, toI + 1).map((b) => b.id);
  if (movedIds.length >= blocks.length) return []; // never empty the document
  const removed = new Set(movedIds);
  const ops: Op[] = [];
  for (const id of movedIds) ops.push({ type: "removeBlock", blockId: id });

  const remaining = blocks.filter((b) => !removed.has(b.id));
  const lastB = remaining[remaining.length - 1];
  if (hasSectionBreak(lastB)) ops.push({ type: "setParaStyle", blockId: lastB!.id, patch: { sectionBreak: undefined } });
  if (fromI === 0) {
    const first = remaining[0];
    if (isPara(first) && first.style.pageBreakBefore === true)
      ops.push({ type: "setParaStyle", blockId: first.id, patch: { pageBreakBefore: false } });
  }
  return ops;
}
