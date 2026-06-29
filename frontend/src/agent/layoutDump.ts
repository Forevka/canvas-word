// Serialize the LayoutTree into a compact, JSON-friendly view for diagnostics —
// the data an agent reads to investigate "this document renders weirdly" reports
// (text placement, overlaps, wrong page breaks). Coordinates are page-local CSS px
// (same frame the paint layer draws in), rounded to keep payloads small.

import type { CharStyle } from "@cw/shared";
import type {
  InlineFragment,
  LayoutTree,
  LineBox,
  Page,
  PlacedBlock,
  PlacedTable,
} from "../layout/layoutTree";

export interface LayoutDumpOptions {
  /** Restrict to one page (0-based `Page.index`). */
  page?: number;
  /** Restrict to a single model block id (searched on every page, incl. table cells). */
  blockId?: string;
  /** Include each fragment's rendered text (default true). Set false for pure geometry. */
  includeText?: boolean;
  /** Cap on fragments emitted per line (guards against pathological lines). */
  maxFragmentsPerLine?: number;
}

const r2 = (n: number): number => Math.round(n * 100) / 100;

/** Only the style bits that affect placement/appearance, omitting defaults so the
 *  dump stays readable. */
function dumpStyle(s: CharStyle): Record<string, unknown> {
  const out: Record<string, unknown> = { font: s.fontFamily, sizePx: s.fontSizePx };
  if (s.bold) out.bold = true;
  if (s.italic) out.italic = true;
  if (s.underline) out.underline = true;
  if (s.strikethrough) out.strike = true;
  if (s.color && s.color !== "#000000") out.color = s.color;
  if (s.verticalAlign) out.verticalAlign = s.verticalAlign;
  if (s.highlightColor) out.highlight = s.highlightColor;
  if (s.link) out.link = s.link;
  if (s.hidden) out.hidden = true;
  return out;
}

function dumpFragment(f: InlineFragment, includeText: boolean): Record<string, unknown> {
  const out: Record<string, unknown> = {
    x: r2(f.x),
    width: r2(f.width),
    startOffset: f.startOffset,
    endOffset: f.endOffset,
    style: dumpStyle(f.style),
  };
  if (includeText) out.text = f.text;
  if (f.wordSpacingPx !== undefined) out.wordSpacingPx = r2(f.wordSpacingPx);
  // A present offsetMap means pretext collapsed whitespace — a frequent culprit in
  // "text overlaps / spacing looks wrong" reports, so flag it.
  if (f.offsetMap) out.whitespaceCollapsed = true;
  return out;
}

function dumpLine(l: LineBox, includeText: boolean, maxFrags: number): Record<string, unknown> {
  const frags = l.fragments.slice(0, maxFrags).map((f) => dumpFragment(f, includeText));
  const out: Record<string, unknown> = {
    y: r2(l.y),
    height: r2(l.height),
    ascent: r2(l.ascent),
    fragments: frags,
  };
  if (l.fragments.length > maxFrags) out.fragmentsTruncated = l.fragments.length - maxFrags;
  if (l.emptyOffset !== undefined) out.emptyOffset = l.emptyOffset;
  if (l.leaders && l.leaders.length > 0) out.leaderCount = l.leaders.length;
  return out;
}

function dumpTable(t: PlacedTable, includeText: boolean, maxFrags: number): Record<string, unknown> {
  return {
    x: r2(t.x),
    y: r2(t.y),
    width: r2(t.width),
    height: r2(t.height),
    colWidths: t.colWidths.map(r2),
    rows: t.rows.map((row) => ({
      y: r2(row.y),
      height: r2(row.height),
      cells: row.cells.map((c) => ({
        x: r2(c.x),
        y: r2(c.y),
        width: r2(c.width),
        height: r2(c.height),
        ...(c.shading ? { shading: c.shading } : {}),
        blocks: c.blocks.map((b) => dumpBlock(b, includeText, maxFrags)),
      })),
    })),
  };
}

function dumpBlock(b: PlacedBlock, includeText: boolean, maxFrags: number): Record<string, unknown> {
  const kind = b.image ? "image" : b.table ? "table" : "paragraph";
  const out: Record<string, unknown> = {
    blockId: b.blockId,
    kind,
    x: r2(b.x),
    y: r2(b.y),
    firstLineIndex: b.firstLineIndex,
  };
  if (b.marker) out.marker = { text: b.marker.text, x: r2(b.marker.x) };
  if (b.toc) out.toc = { numText: b.toc.numText, numX: r2(b.toc.numX), targetId: b.toc.targetId };
  if (b.image) {
    out.image = {
      width: r2(b.image.width),
      height: r2(b.image.height),
      // src can be a huge data: URL — report only its scheme/length, not the bytes.
      srcKind: b.image.src.slice(0, b.image.src.indexOf(":") + 1) || "unknown",
      srcLength: b.image.src.length,
      ...(b.image.clip ? { clipped: true } : {}),
    };
  } else if (b.table) {
    out.table = dumpTable(b.table, includeText, maxFrags);
  } else {
    out.lines = b.lines.map((l) => dumpLine(l, includeText, maxFrags));
  }
  return out;
}

function matchesBlock(b: PlacedBlock, blockId: string): boolean {
  if (b.blockId === blockId) return true;
  if (b.table) {
    for (const row of b.table.rows)
      for (const cell of row.cells) for (const cb of cell.blocks) if (matchesBlock(cb, blockId)) return true;
  }
  return false;
}

function dumpPage(p: Page, opts: { includeText: boolean; maxFragmentsPerLine: number; blockId?: string }): Record<string, unknown> {
  let blocks = p.blocks;
  if (opts.blockId) blocks = blocks.filter((b) => matchesBlock(b, opts.blockId!));
  return {
    index: p.index,
    number: p.number,
    widthPx: r2(p.widthPx),
    heightPx: r2(p.heightPx),
    marginPx: {
      top: r2(p.marginPx.top),
      right: r2(p.marginPx.right),
      bottom: r2(p.marginPx.bottom),
      left: r2(p.marginPx.left),
    },
    contentTopPx: r2(p.contentTopPx),
    contentBottomPx: r2(p.contentBottomPx),
    ...(p.footnoteRuleY !== undefined ? { footnoteRuleY: r2(p.footnoteRuleY) } : {}),
    ...(p.endnoteRuleY !== undefined ? { endnoteRuleY: r2(p.endnoteRuleY) } : {}),
    blockCount: p.blocks.length,
    headerBlocks: p.header?.length ?? 0,
    footerBlocks: p.footer?.length ?? 0,
    blocks: blocks.map((b) => dumpBlock(b, opts.includeText, opts.maxFragmentsPerLine)),
  };
}

/** A compact, JSON-serializable snapshot of the laid-out document geometry. */
export function dumpLayout(tree: LayoutTree, opts: LayoutDumpOptions = {}): Record<string, unknown> {
  const includeText = opts.includeText ?? true;
  const maxFrags = opts.maxFragmentsPerLine ?? 200;
  let pages = tree.pages;
  if (opts.page !== undefined) pages = pages.filter((p) => p.index === opts.page);
  const pageOpts: { includeText: boolean; maxFragmentsPerLine: number; blockId?: string } = {
    includeText,
    maxFragmentsPerLine: maxFrags,
    ...(opts.blockId !== undefined ? { blockId: opts.blockId } : {}),
  };
  return {
    pageCount: tree.pages.length,
    defaultPageWidthPx: r2(tree.pageWidthPx),
    defaultPageHeightPx: r2(tree.pageHeightPx),
    pages: pages.map((p) => dumpPage(p, pageOpts)),
  };
}
