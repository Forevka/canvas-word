// Layout tree: turns the laid-out geometry (LayoutTree) into navigable nodes —
// pages → placed blocks → lines → fragments, and tables → rows → cells → blocks.
// Every node carries a literal page-coordinate rect so hovering highlights its
// exact painted region (fragment-precise), and placed blocks keep their blockId so
// the canvas→tree reverse highlight + caret outline still work.

import type { LayoutTree, LineBox, Page, PlacedBlock } from "../../layout/layoutTree";
import type { TreeNode } from "./types";
import { previewText, shortId } from "./types";

const r0 = (n: number): number => Math.round(n);

/** Bounding rect (page coords) of a placed block — image/table box or paragraph
 *  line-stack extent. */
function blockRect(pb: PlacedBlock): { x: number; y: number; width: number; height: number } {
  if (pb.table) return { x: pb.table.x, y: pb.table.y, width: pb.table.width, height: pb.table.height };
  if (pb.image) return { x: pb.x, y: pb.y, width: pb.image.width, height: pb.image.height };
  let top = pb.y + (pb.lines[0]?.y ?? 0);
  let bot = top;
  let right = pb.x;
  for (const line of pb.lines) {
    const ly = pb.y + line.y;
    top = Math.min(top, ly);
    bot = Math.max(bot, ly + line.height);
    for (const f of line.fragments) right = Math.max(right, pb.x + f.x + f.width);
  }
  return { x: pb.x, y: top, width: Math.max(1, right - pb.x), height: Math.max(1, bot - top) };
}

function lineNode(pb: PlacedBlock, line: LineBox, pageIndex: number, i: number): TreeNode {
  const y = pb.y + line.y;
  let lx = pb.x + (line.fragments[0]?.x ?? 0);
  let lr = lx;
  for (const f of line.fragments) { lx = Math.min(lx, pb.x + f.x); lr = Math.max(lr, pb.x + f.x + f.width); }
  const width = Math.max(1, lr - lx);
  const frags: TreeNode[] = line.fragments.map((f, fi) => ({
    key: `${pb.blockId}#l${i}f${fi}`, kind: "run",
    label: "frag",
    preview: `${previewText(f.text) || "∅"} [${f.startOffset}–${f.endOffset}]`,
    badges: [`${r0(f.width)}px`],
    target: { kind: "rect", pageIndex, x: pb.x + f.x, y, width: Math.max(1, f.width), height: line.height, label: "frag" },
    data: { x: r0(pb.x + f.x), y: r0(y), width: r0(f.width), height: r0(line.height), startOffset: f.startOffset, endOffset: f.endOffset, text: f.text },
    children: [],
  }));
  return {
    key: `${pb.blockId}#l${i}`, kind: "layout",
    label: `line ${i}`,
    preview: `y=${r0(line.y)} h=${r0(line.height)} asc=${r0(line.ascent)}`,
    badges: [`${line.fragments.length} frag`],
    target: { kind: "rect", pageIndex, x: lx, y, width, height: line.height, label: `line ${i}` },
    data: { y: r0(line.y), height: r0(line.height), ascent: r0(line.ascent), fragments: line.fragments.length, emptyOffset: line.emptyOffset },
    children: frags,
  };
}

function placedBlockNode(pb: PlacedBlock, pageIndex: number): TreeNode {
  const rect = blockRect(pb);
  const target = { kind: "rect" as const, pageIndex, ...rect, label: shortId(pb.blockId) };
  const base = { blockId: pb.blockId, target, badges: [`@${r0(rect.x)},${r0(rect.y)}`] };
  if (pb.table) {
    const t = pb.table;
    return {
      key: `L:${pb.blockId}`, kind: "block", label: "▦ table", preview: `${t.rows.length} rows`, ...base,
      data: { blockId: pb.blockId, x: r0(t.x), y: r0(t.y), width: r0(t.width), height: r0(t.height), rows: t.rows.length },
      children: t.rows.map((row, ri) => ({
        key: `L:${pb.blockId}#r${ri}`, kind: "tag", label: `row ${ri}`, preview: `y=${r0(row.y)} h=${r0(row.height)}`, badges: [`${row.cells.length} cells`],
        data: { y: r0(row.y), height: r0(row.height) },
        children: row.cells.map((cell, ci) => ({
          key: `L:${pb.blockId}#r${ri}c${ci}`, kind: "tag", label: `cell ${ri},${ci}`,
          preview: `${r0(cell.width)}×${r0(cell.height)}`,
          badges: cell.originRow !== ri || cell.originCol !== ci ? [`origin ${cell.originRow},${cell.originCol}`] : [],
          target: { kind: "rect", pageIndex, x: cell.x, y: cell.y, width: cell.width, height: cell.height, label: `cell ${ri},${ci}` },
          data: { x: r0(cell.x), y: r0(cell.y), width: r0(cell.width), height: r0(cell.height), originRow: cell.originRow, originCol: cell.originCol },
          children: cell.blocks.map((cb) => placedBlockNode(cb, pageIndex)),
        })),
      })),
    };
  }
  if (pb.image) {
    return {
      key: `L:${pb.blockId}`, kind: "block", label: "🖼 image", preview: `${r0(pb.image.width)}×${r0(pb.image.height)}`, ...base,
      data: { blockId: pb.blockId, x: r0(pb.x), y: r0(pb.y), width: r0(pb.image.width), height: r0(pb.image.height) },
      children: [],
    };
  }
  return {
    key: `L:${pb.blockId}`, kind: "block", label: "¶ block", preview: previewText(pb.lines.flatMap((l) => l.fragments.map((f) => f.text)).join("")) || "(empty)", ...base,
    data: { blockId: pb.blockId, x: r0(pb.x), y: r0(pb.y), firstLineIndex: pb.firstLineIndex, lines: pb.lines.length },
    children: pb.lines.map((line, i) => lineNode(pb, line, pageIndex, i)),
  };
}

/** Build the layout (geometry) tree: one root per page. */
export function buildLayoutTree(tree: LayoutTree): TreeNode[] {
  return tree.pages.map((page: Page) => ({
    key: `L:page${page.index}`, kind: "tag",
    label: `Page ${page.index}`,
    preview: `#${page.number} · ${r0(page.widthPx)}×${r0(page.heightPx)} · ${page.blocks.length} blocks`,
    badges: [],
    data: { index: page.index, number: page.number, widthPx: r0(page.widthPx), heightPx: r0(page.heightPx), marginPx: page.marginPx },
    children: page.blocks.map((pb) => placedBlockNode(pb, page.index)),
  }));
}

/** Cheap signature: page/block counts + rounded block positions (changes on reflow). */
export function layoutSignature(tree: LayoutTree): string {
  let h = tree.pages.length;
  for (const p of tree.pages) {
    h = (h * 31 + p.blocks.length) | 0;
    for (const b of p.blocks) h = (((h * 31 + Math.round(b.x)) | 0) * 31 + Math.round(b.y)) | 0;
  }
  return String(h);
}
