// Headless TOC page-number recalculation — the editor's "recalculate TOC" (Word's
// F9) lifted out of the editor closure into a pure, Node-safe core. No DOM, no
// dispatch: it reads a laid-out document and returns the text edits that bring each
// TOC entry's trailing page number in line with the heading's CURRENT page.
//
// A TOC entry is a body paragraph "<label><tab><number>" whose run carries an
// in-document link (style.link = "#<bookmark>"); the bookmark resolves through
// doc.bookmarks to the target heading block, and the layout tree tells us which
// page that block landed on. The number is plain text, so baking it makes the new
// value survive .docx export (Word shows it statically until the field is refreshed).
//
// The editor (frontend/src/index.ts) and the backend recalc route both call this —
// the editor wraps the edits in ops for undo/collab, the backend applies them to
// the model directly.

import {
  applyOp,
  styleAtRuns,
  textOfRuns,
  type CharStyle,
  type Document,
  type Op,
} from "@cw/shared";
import { createLayoutEngine } from "../layout/engine";
import type { LayoutTree, PlacedBlock } from "../layout/layoutTree";

/** One text replacement inside a TOC paragraph: the page-number digits. */
export interface TocEdit {
  blockId: string;
  /** Offset of the first digit in the paragraph's flat text. */
  start: number;
  /** Offset just past the last digit. */
  end: number;
  /** New page-number string. */
  text: string;
  /** Resolved char style at `start` — carried so insertText keeps the entry's font. */
  style?: CharStyle | undefined;
}

/** blockId → displayed page number of the FIRST page it appears on (body only,
 *  recursing into table cells so headings inside tables resolve too). */
function pageOfBlockMap(tree: LayoutTree): Map<string, number> {
  const pageOfBlock = new Map<string, number>();
  const scan = (blocks: PlacedBlock[], pageNum: number): void => {
    for (const pb of blocks) {
      if (!pageOfBlock.has(pb.blockId)) pageOfBlock.set(pb.blockId, pageNum);
      if (pb.table) for (const row of pb.table.rows) for (const cell of row.cells) scan(cell.blocks, pageNum);
    }
  };
  for (const pg of tree.pages) scan(pg.blocks, pg.number);
  return pageOfBlock;
}

/** Compute the page-number edits for every TOC entry, given a layout tree. Pure:
 *  mutates nothing. Returns [] when every entry is already current. The page
 *  numbers come from `tree`, but the offsets are computed against `doc` — pass an
 *  un-baked doc with a baked tree to build edits that apply to the un-baked head. */
export function computeTocEdits(doc: Document, tree: LayoutTree): TocEdit[] {
  const pageOfBlock = pageOfBlockMap(tree);
  const edits: TocEdit[] = [];
  for (const b of doc.blocks) {
    if (b.kind !== "paragraph") continue;
    const anchor = b.runs.find((r) => r.style.link?.startsWith("#"))?.style.link?.slice(1);
    if (!anchor) continue;
    const targetId = doc.bookmarks?.[anchor]?.start.blockId;
    if (!targetId) continue;
    const num = pageOfBlock.get(targetId);
    if (num === undefined) continue; // dangling entry: heading was deleted
    const full = textOfRuns(b.runs);
    const m = full.match(/\t(\d+)(\s*)$/); // trailing tab + page number
    if (!m) continue;
    if (m[1] === String(num)) continue; // already current
    const numStart = full.length - m[2]!.length - m[1]!.length;
    edits.push({
      blockId: b.id,
      start: numStart,
      end: numStart + m[1]!.length,
      text: String(num),
      style: styleAtRuns(b.runs, numStart) ?? b.runs[0]?.style,
    });
  }
  return edits;
}

/** Convenience for headless callers: lay the doc out internally, then compute the
 *  edits. The caller MUST have awaited installMeasureHost() first — the layout
 *  engine measures text, which needs the fontkit measure context. */
export function computeTocEditsWithLayout(doc: Document): TocEdit[] {
  const tree = createLayoutEngine().layout(doc);
  return computeTocEdits(doc, tree);
}

/** TocEdit[] → Op[] (deleteRange + insertText), the same pair the editor emits. */
export function tocEditsToOps(edits: readonly TocEdit[]): Op[] {
  const ops: Op[] = [];
  for (const e of edits) {
    ops.push({ type: "deleteRange", blockId: e.blockId, start: e.start, end: e.end });
    ops.push({
      type: "insertText",
      at: { blockId: e.blockId, offset: e.start },
      text: e.text,
      ...(e.style ? { style: e.style } : {}),
    });
  }
  return ops;
}

/** Lay out, compute edits, and apply them to the model. Returns the updated
 *  document (applyOp is immutable) and the number of entries changed. For the
 *  stateless route. The caller MUST have awaited installMeasureHost() first. */
export function recalcToc(doc: Document): { doc: Document; changed: number } {
  const edits = computeTocEditsWithLayout(doc);
  if (edits.length === 0) return { doc, changed: 0 };
  let next = doc;
  // Each edit targets a distinct paragraph (one TOC entry), so the delete+insert
  // pairs never shift one another's offsets. The blocks all exist — the edits were
  // just computed against this same doc.
  for (const op of tocEditsToOps(edits)) {
    next = applyOp(next, op).doc;
  }
  return { doc: next, changed: edits.length };
}
