// Commands: pure (state) -> Transaction | null. The keymap, the IME proxy, and
// (later) toolbar buttons all dispatch through these.

import type { Block, CharStyle, ImageBlock, ParaStyle, Paragraph, Run, SdtProps, SdtType, TableBlock, TableCell, TableRow } from "../model/document";
import type { DocPosition, DocSelection } from "../model/position";
import { isCollapsed } from "../model/position";
import type { Op, SectionGeometry } from "../model/ops";
import { sliceRuns, applyStylePatchToRuns, containerOf, containerBlocks, locateImage } from "../model/ops";
import {
  blockById,
  blockIndexOf,
  isInCell,
  locateParagraph,
  paragraphsOf,
  prevGrapheme,
  nextGrapheme,
  styleAtRuns,
  textOfBlock,
  textOfRuns,
  words,
} from "../model/text";
import type { DocFragment } from "../input/clipboard";
import type { Command, EditorState, Transaction, TransactionOrigin } from "./state";

const caret = (blockId: string, offset: number): DocSelection => ({
  anchor: { blockId, offset },
  focus: { blockId, offset },
});

let blockIdCounter = 0;
const freshBlockId = (): string => `n${Date.now().toString(36)}_${blockIdCounter++}`;

/** Order a selection by document position (model order, no layout needed). */
function orderedRange(state: EditorState, sel: DocSelection): [DocPosition, DocPosition] {
  const ai = blockIndexOf(state.doc, sel.anchor.blockId);
  const fi = blockIndexOf(state.doc, sel.focus.blockId);
  if (ai < fi || (ai === fi && sel.anchor.offset <= sel.focus.offset)) {
    return [sel.anchor, sel.focus];
  }
  return [sel.focus, sel.anchor];
}

/** Ops that remove the selected range. Cross-block: trim first tail, trim last
 *  head, drop middles (ANY top-level block kind, images and tables included),
 *  merge first+last. Caret lands at the range start.
 *  Returns null for unsupported shapes: ranges that cross a table-cell
 *  boundary cannot be structurally merged. */
function deleteRangeOps(state: EditorState, from: DocPosition, to: DocPosition): Op[] | null {
  if (from.blockId === to.blockId) {
    if (from.offset === to.offset) return [];
    return [{ type: "deleteRange", blockId: from.blockId, start: from.offset, end: to.offset }];
  }
  // Cross-paragraph: both ends must be top-level blocks of the SAME container
  // (body, or one band story) — no merging across cell/story boundaries.
  const fromC = containerOf(state.doc, from.blockId);
  const toC = containerOf(state.doc, to.blockId);
  if (!fromC || !toC || fromC.where !== toC.where) return null;
  const docBlocks = containerBlocks(state.doc, fromC.where);
  const fi = fromC.index;
  const ti = toC.index;
  const ops: Op[] = [];
  const firstLen = textOfBlock(state.doc, from.blockId).length;
  if (from.offset < firstLen) {
    ops.push({ type: "deleteRange", blockId: from.blockId, start: from.offset, end: firstLen });
  }
  if (to.offset > 0) {
    ops.push({ type: "deleteRange", blockId: to.blockId, start: 0, end: to.offset });
  }
  for (let i = fi + 1; i < ti; i++) {
    ops.push({ type: "removeBlock", blockId: docBlocks[i]!.id });
  }
  ops.push({ type: "mergeParagraphs", firstBlockId: from.blockId });
  return ops;
}

function withSelectionDeleted(
  state: EditorState,
): { ops: Op[]; at: DocPosition } | null {
  const sel = state.selection;
  if (!sel) return null;
  if (isCollapsed(sel)) return { ops: [], at: sel.focus };
  const [from, to] = orderedRange(state, sel);
  const ops = deleteRangeOps(state, from, to);
  if (ops === null) return null;
  return { ops, at: from };
}

const tr = (ops: Op[], selectionAfter: DocSelection | null, origin: TransactionOrigin): Transaction => ({
  ops,
  selectionAfter,
  origin,
});

// ---------------------------------------------------------------------------

export function insertText(text: string, origin: TransactionOrigin = "typing"): Command {
  return (state) => {
    const base = withSelectionDeleted(state);
    if (!base || text.length === 0) return null;
    const op: Op = { type: "insertText", at: base.at, text };
    // A pending toggle at a collapsed caret styles the next typed text.
    if (state.pendingStyle) {
      const block = blockById(state.doc, base.at.blockId);
      const inherited = block ? styleAtRuns(block.runs, base.at.offset) : undefined;
      if (inherited) op.style = { ...inherited, ...state.pendingStyle };
    }
    base.ops.push(op);
    return tr(base.ops, caret(base.at.blockId, base.at.offset + text.length), origin);
  };
}

/** Insert a (possibly multi-paragraph) fragment at the selection. Multi-block
 *  insertion avoids index-based ops entirely: split at the caret, then grow a
 *  chain of empty paragraphs by splitting the tail at offset 0 and filling each. */
export function insertFragment(fragment: DocFragment): Command {
  return (state) => {
    const base = withSelectionDeleted(state);
    if (!base || fragment.blocks.length === 0) return null;
    const ops = base.ops;
    const at = base.at;
    const first = fragment.blocks[0]!;

    if (fragment.inline) {
      const len = textOfRuns(first.runs).length;
      if (len === 0) return null;
      ops.push({ type: "insertRuns", at, runs: first.runs });
      return tr(ops, caret(at.blockId, at.offset + len), "paste");
    }
    if (isInCell(state.doc, at.blockId)) return null; // multi-paragraph paste needs splits

    let tailPtr = freshBlockId();
    ops.push({ type: "splitParagraph", at, newBlockId: tailPtr });
    if (textOfRuns(first.runs).length > 0) {
      ops.push({ type: "insertRuns", at, runs: first.runs });
    }
    let caretAfter = caret(tailPtr, 0);
    for (let i = 1; i < fragment.blocks.length; i++) {
      const fb = fragment.blocks[i]!;
      const isLast = i === fragment.blocks.length - 1;
      if (!isLast) {
        // tailPtr is currently the paragraph holding the post-caret text; split
        // it at 0 so tailPtr becomes an EMPTY paragraph we can fill, and the
        // post-caret text moves into a fresh tail.
        const nextTail = freshBlockId();
        ops.push({ type: "splitParagraph", at: { blockId: tailPtr, offset: 0 }, newBlockId: nextTail });
        if (textOfRuns(fb.runs).length > 0) {
          ops.push({ type: "insertRuns", at: { blockId: tailPtr, offset: 0 }, runs: fb.runs });
        }
        ops.push({ type: "setParaStyle", blockId: tailPtr, patch: fb.style });
        tailPtr = nextTail;
      } else {
        const len = textOfRuns(fb.runs).length;
        if (len > 0) {
          ops.push({ type: "insertRuns", at: { blockId: tailPtr, offset: 0 }, runs: fb.runs });
        }
        caretAfter = caret(tailPtr, len);
      }
    }
    return tr(ops, caretAfter, "paste");
  };
}

export function deleteBackward(): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel) return null;
    if (!isCollapsed(sel)) {
      const base = withSelectionDeleted(state)!;
      return tr(base.ops, caret(base.at.blockId, base.at.offset), "typing");
    }
    const { blockId, offset } = sel.focus;
    if (offset > 0) {
      const start = prevGrapheme(textOfBlock(state.doc, blockId), offset);
      return tr(
        [{ type: "deleteRange", blockId, start, end: offset }],
        caret(blockId, start),
        "typing",
      );
    }
    // At paragraph start: merge with the previous block in the SAME container
    // (body or band story; cells merge only WITHIN their own cell).
    const found = containerOf(state.doc, blockId);
    if (!found) {
      const loc = locateParagraph(state.doc, blockId);
      if (loc?.kind === "cell" && loc.pi > 0) {
        const table = containerBlocks(state.doc, loc.where)[loc.bi] as TableBlock;
        const row = table.rows[loc.ri]!;
        const cell = row.cells[loc.ci]!;
        const prevBlock = cell.blocks[loc.pi - 1]!;
        if (prevBlock.kind === "image") {
          // Backspace before an image inside a cell deletes the image.
          const cells = row.cells.slice();
          cells[loc.ci] = { ...cell, blocks: cell.blocks.filter((b) => b.id !== prevBlock.id) };
          return tr(
            [{ type: "setTableRow", tableId: table.id, rowIndex: loc.ri, row: { cells } }],
            caret(blockId, 0),
            "command",
          );
        }
        if (prevBlock.kind !== "paragraph") return null;
        return tr(
          [{ type: "mergeParagraphs", firstBlockId: prevBlock.id }],
          caret(prevBlock.id, textOfRuns(prevBlock.runs).length),
          "command",
        );
      }
      return null; // first paragraph of a cell: the cell boundary stops backspace
    }
    // Word's ladder: list level/membership goes first...
    const ladder = listBackspaceLadder(state);
    if (ladder) return ladder;
    // ...then Backspace at the start of a page/column-break paragraph removes
    // the BREAK; the paragraphs merge only on the next press.
    const here = containerBlocks(state.doc, found.where)[found.index];
    if (here?.kind === "paragraph" && here.style.pageBreakBefore === true) {
      return tr(
        [{ type: "setParaStyle", blockId, patch: { pageBreakBefore: false } }],
        caret(blockId, 0),
        "command",
      );
    }
    if (here?.kind === "paragraph" && here.style.columnBreakBefore === true) {
      return tr(
        [{ type: "setParaStyle", blockId, patch: { columnBreakBefore: false } }],
        caret(blockId, 0),
        "command",
      );
    }
    const prev = containerBlocks(state.doc, found.where)[found.index - 1];
    if (!prev) return null;
    if (prev.kind === "image") {
      // Backspace before an image deletes the image.
      return tr([{ type: "removeBlock", blockId: prev.id }], caret(blockId, 0), "command");
    }
    if (prev.kind !== "paragraph") return null; // table boundary: stop
    const prevLen = textOfRuns(prev.runs).length;
    return tr(
      [{ type: "mergeParagraphs", firstBlockId: prev.id }],
      caret(prev.id, prevLen),
      "command", // structural deletes don't coalesce with typing
    );
  };
}

export function deleteForward(): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel) return null;
    if (!isCollapsed(sel)) {
      const base = withSelectionDeleted(state)!;
      return tr(base.ops, caret(base.at.blockId, base.at.offset), "typing");
    }
    const { blockId, offset } = sel.focus;
    const text = textOfBlock(state.doc, blockId);
    if (offset < text.length) {
      const end = nextGrapheme(text, offset);
      return tr([{ type: "deleteRange", blockId, start: offset, end }], caret(blockId, offset), "typing");
    }
    const found = containerOf(state.doc, blockId);
    if (!found) {
      const loc = locateParagraph(state.doc, blockId);
      if (loc?.kind === "cell") {
        const table = containerBlocks(state.doc, loc.where)[loc.bi] as TableBlock;
        const row = table.rows[loc.ri]!;
        const cell = row.cells[loc.ci]!;
        const nextBlock = cell.blocks[loc.pi + 1];
        if (nextBlock?.kind === "image") {
          const cells = row.cells.slice();
          cells[loc.ci] = { ...cell, blocks: cell.blocks.filter((b) => b.id !== nextBlock.id) };
          return tr(
            [{ type: "setTableRow", tableId: table.id, rowIndex: loc.ri, row: { cells } }],
            caret(blockId, offset),
            "command",
          );
        }
        if (nextBlock?.kind === "paragraph") {
          return tr([{ type: "mergeParagraphs", firstBlockId: blockId }], caret(blockId, offset), "command");
        }
      }
      return null;
    }
    const next = containerBlocks(state.doc, found.where)[found.index + 1];
    if (!next) return null;
    if (next.kind === "image") {
      return tr([{ type: "removeBlock", blockId: next.id }], caret(blockId, offset), "command");
    }
    if (next.kind !== "paragraph") return null;
    return tr([{ type: "mergeParagraphs", firstBlockId: blockId }], caret(blockId, offset), "command");
  };
}

export function splitParagraph(): Command {
  return (state) => {
    const base = withSelectionDeleted(state);
    if (!base) return null;
    // Enter on an EMPTY list paragraph exits the list instead of splitting (Word).
    if (base.ops.length === 0) {
      const block = blockById(state.doc, base.at.blockId);
      if (block?.style.list && textOfRuns(block.runs).length === 0) {
        return tr(
          [{ type: "setParaStyle", blockId: block.id, patch: { list: undefined } }],
          state.selection,
          "command",
        );
      }
    }
    // Works everywhere: body, band stories, AND table cells (the op splices
    // within the cell's paragraph list) — multi-paragraph cells, like Word.
    // List membership is inherited by the split (style is cloned).
    const newBlockId = freshBlockId();
    base.ops.push({ type: "splitParagraph", at: base.at, newBlockId });
    return tr(base.ops, caret(newBlockId, 0), "command");
  };
}

// ---------------------------------------------------------------------------
// Lists

import {
  defaultListDefinition,
  DEFAULT_BULLET_LIST_ID,
  DEFAULT_NUMBER_LIST_ID,
} from "../model/lists";

/** Top-level paragraphs covered by the selection (lists are body-only for now). */
function selectedTopLevelParagraphs(state: EditorState): Paragraph[] {
  const out: Paragraph[] = [];
  const seen = new Set<string>();
  for (const s of selectedSegments(state)) {
    if (seen.has(s.block.id)) continue;
    seen.add(s.block.id);
    if (containerOf(state.doc, s.block.id)?.where === "body") out.push(s.block);
  }
  return out;
}

export function toggleList(kind: "bullet" | "decimal"): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel) return null;
    const paragraphs = selectedTopLevelParagraphs(state);
    if (paragraphs.length === 0) return null;
    const listId = kind === "bullet" ? DEFAULT_BULLET_LIST_ID : DEFAULT_NUMBER_LIST_ID;
    const allIn = paragraphs.every((p) => p.style.list?.listId === listId);
    const ops: Op[] = [];
    if (!allIn && !state.doc.lists?.[listId]) {
      ops.push({ type: "setListDefinition", listId, def: defaultListDefinition(kind) });
    }
    for (const p of paragraphs) {
      ops.push({
        type: "setParaStyle",
        blockId: p.id,
        patch: allIn
          ? { list: undefined }
          : { list: { listId, level: p.style.list?.level ?? 0 }, indentFirstLinePx: 0 },
      });
    }
    return tr(ops, sel, "command");
  };
}

/** Tab / Shift+Tab at the start of a list paragraph: demote / promote. */
export function changeListLevel(delta: 1 | -1): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel) return null;
    const paragraphs = selectedTopLevelParagraphs(state).filter((p) => p.style.list);
    if (paragraphs.length === 0) return null;
    const ops: Op[] = paragraphs.map((p) => ({
      type: "setParaStyle",
      blockId: p.id,
      patch: {
        list: { listId: p.style.list!.listId, level: Math.max(0, Math.min(8, p.style.list!.level + delta)) },
      },
    }));
    return tr(ops, sel, "command");
  };
}

/** Backspace at the start of a list paragraph: demote, then leave the list —
 *  Word's ladder, ahead of page-break removal and paragraph merging. */
export function listBackspaceLadder(state: EditorState): Transaction | null {
  const sel = state.selection;
  if (!sel || !isCollapsed(sel) || sel.focus.offset !== 0) return null;
  const block = blockById(state.doc, sel.focus.blockId);
  if (!block?.style.list || containerOf(state.doc, block.id)?.where !== "body") return null;
  const patch: Partial<ParaStyle> =
    block.style.list.level > 0
      ? { list: { listId: block.style.list.listId, level: block.style.list.level - 1 } }
      : { list: undefined };
  return tr([{ type: "setParaStyle", blockId: block.id, patch }], sel, "command");
}

// ---------------------------------------------------------------------------
// Sections

/** Page-setup geometry of the section containing top-level block `fromIndex`:
 *  the first section-break paragraph at/after it terminates that section;
 *  none → the document-final section (`doc.section`). */
function sectionGeometryAt(doc: EditorState["doc"], fromIndex: number): SectionGeometry {
  for (let i = fromIndex; i < doc.blocks.length; i++) {
    const b = doc.blocks[i]!;
    if (b.kind === "paragraph" && b.style.sectionBreak) {
      const p = b.style.sectionBreak.props;
      const columns = p.columns === undefined ? (doc.section.columns ?? null) : p.columns;
      return {
        pageWidthPx: p.pageWidthPx ?? doc.section.pageWidthPx,
        pageHeightPx: p.pageHeightPx ?? doc.section.pageHeightPx,
        marginPx: { ...(p.marginPx ?? doc.section.marginPx) },
        columns: columns ? { ...columns } : null,
        pageNumberStart: p.pageNumberStart ?? null, // restart is per-section, never inherited
      };
    }
  }
  return {
    pageWidthPx: doc.section.pageWidthPx,
    pageHeightPx: doc.section.pageHeightPx,
    marginPx: { ...doc.section.marginPx },
    columns: doc.section.columns ? { ...doc.section.columns } : null,
    pageNumberStart: doc.section.pageNumberStart ?? null,
  };
}

/** Insert a next-page section break at the caret: split, then mark the FIRST
 *  half as the section terminator (OOXML sectPr placement). The break snapshots
 *  the current section's geometry so later page-setup changes to the following
 *  section don't retroactively reshape this one. Bands stay inherited. */
export function insertSectionBreak(): Command {
  return (state) => {
    const base = withSelectionDeleted(state);
    if (!base) return null;
    const c = containerOf(state.doc, base.at.blockId);
    if (!c || c.where !== "body") return null; // top-level body paragraphs only
    const block = blockById(state.doc, base.at.blockId);
    if (!block) return null;
    const geo = sectionGeometryAt(state.doc, c.index);
    const props: import("../model/document").SectionPatch = {
      pageWidthPx: geo.pageWidthPx,
      pageHeightPx: geo.pageHeightPx,
      marginPx: geo.marginPx,
      columns: geo.columns,
    };
    if (geo.pageNumberStart !== null) props.pageNumberStart = geo.pageNumberStart;
    const newBlockId = freshBlockId();
    // The tail must NOT clone an existing sectionBreak (splitting the break
    // paragraph itself would otherwise duplicate the section).
    const { sectionBreak: _sb, ...restStyle } = block.style;
    base.ops.push({ type: "splitParagraph", at: base.at, newBlockId, newStyle: { ...restStyle } });
    base.ops.push({
      type: "setParaStyle",
      blockId: base.at.blockId,
      patch: { sectionBreak: { type: "nextPage", props } },
    });
    return tr(base.ops, caret(newBlockId, 0), "command");
  };
}

/** Apply page-setup geometry to the CARET's section: patch its terminating
 *  break paragraph, or `doc.section` when the caret sits in the final section. */
export function applyPageSetup(geometry: SectionGeometry): Command {
  return (state) => {
    const sel = state.selection;
    let fromIndex = 0;
    if (sel) {
      const loc = locateParagraph(state.doc, sel.focus.blockId);
      const outsideBody =
        loc?.kind === "band" || loc?.kind === "footnote" || (loc?.kind === "cell" && loc.where !== "body");
      if (outsideBody) {
        // Band/footnote stories live on doc.section — target the final section.
        return tr([{ type: "setSectionProps", geometry }], sel, "command");
      }
      if (loc) fromIndex = loc.bi; // body cells → owning table's index
    }
    for (let i = fromIndex; i < state.doc.blocks.length; i++) {
      const b = state.doc.blocks[i]!;
      if (b.kind === "paragraph" && b.style.sectionBreak) {
        const props: typeof b.style.sectionBreak.props = {
          ...b.style.sectionBreak.props,
          pageWidthPx: geometry.pageWidthPx,
          pageHeightPx: geometry.pageHeightPx,
          marginPx: geometry.marginPx,
          columns: geometry.columns,
        };
        if (geometry.pageNumberStart !== null) props.pageNumberStart = geometry.pageNumberStart;
        else delete props.pageNumberStart;
        return tr(
          [{ type: "setParaStyle", blockId: b.id, patch: { sectionBreak: { type: "nextPage", props } } }],
          sel,
          "command",
        );
      }
    }
    return tr([{ type: "setSectionProps", geometry }], sel, "command");
  };
}

/** Current page-setup of the caret's section (dialog defaults). */
export function pageSetupAt(state: EditorState): SectionGeometry {
  const sel = state.selection;
  let fromIndex = 0;
  if (sel) {
    const loc = locateParagraph(state.doc, sel.focus.blockId);
    const outsideBody =
      loc?.kind === "band" || loc?.kind === "footnote" || (loc?.kind === "cell" && loc.where !== "body");
    if (outsideBody) fromIndex = state.doc.blocks.length; // → doc.section
    else if (loc) fromIndex = loc.bi;
  }
  return sectionGeometryAt(state.doc, fromIndex);
}

/** Word's "Different first page" / "Different odd & even pages" checkboxes:
 *  enabling seeds EMPTY first/even bands (Word starts them blank); disabling
 *  removes them so pages fall back to the default header/footer. Acts on
 *  `doc.section` (the band store story editing targets). */
export function setBandVariantEnabled(kind: "first" | "even", enabled: boolean): Command {
  return (state) => {
    const bands: ["headerFirst", "footerFirst"] | ["headerEven", "footerEven"] =
      kind === "first" ? ["headerFirst", "footerFirst"] : ["headerEven", "footerEven"];
    const has = bands.some((b) => state.doc.section[b] !== undefined);
    if (enabled === has) return null;
    const seedStyle = (base: "header" | "footer"): { char: CharStyle; para: ParaStyle } => {
      const src = state.doc.section[base]?.find((b): b is Paragraph => b.kind === "paragraph");
      return {
        char: src?.runs[0]?.style ?? { ...DEFAULT_CELL_CHAR, fontSizePx: 12, color: "#5f6368" },
        para: src ? { ...src.style } : { align: base === "footer" ? "center" : "left", lineHeight: 1.4, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 },
      };
    };
    const ops: Op[] = bands.map((band) => {
      if (!enabled) return { type: "setSectionBand", band, blocks: null };
      const { char, para } = seedStyle(band.startsWith("header") ? "header" : "footer");
      return {
        type: "setSectionBand",
        band,
        blocks: [{ kind: "paragraph", id: freshBlockId(), revision: 0, runs: [{ text: "", style: char }], style: para }],
      };
    });
    return tr(ops, state.selection, "command");
  };
}

// ---------------------------------------------------------------------------
// Table of contents — entries are REAL paragraphs (lay out, paint, hit-test
// for free) tagged with `tocEntry`; their page numbers are paint-only and
// resolved per relayout, so numbers never go stale. Entry TEXT goes stale when
// headings are renamed — `insertTocCmd` regenerates (Word's update-field).

const TOC_LEVELS = 3;

function headingLevel(p: Paragraph): number | null {
  const m = p.style.namedStyle?.match(/^heading(\d)$/i);
  if (!m) return null;
  const lvl = Number(m[1]);
  return lvl >= 1 && lvl <= TOC_LEVELS ? lvl : null;
}

function tocEntryBlocks(doc: EditorState["doc"]): Paragraph[] {
  const out: Paragraph[] = [];
  for (let i = 0; i < doc.blocks.length; i++) {
    const b = doc.blocks[i]!;
    if (b.kind === "paragraph" && b.style.tocEntry) {
      out.push(b);
      continue;
    }
    // The generated title travels with the entries (regeneration replaces it).
    if (b.kind === "paragraph" && b.style.namedStyle === "tocTitle") out.push(b);
  }
  return out;
}

function buildTocParagraphs(doc: EditorState["doc"]): Paragraph[] {
  const char = (size: number, bold = false): CharStyle => ({
    fontFamily: "Georgia, serif",
    fontSizePx: size,
    bold,
    italic: false,
    underline: false,
    strikethrough: false,
    color: "#202124",
  });
  const out: Paragraph[] = [
    {
      kind: "paragraph",
      id: freshBlockId(),
      revision: 0,
      runs: [{ text: "Table of Contents", style: char(20, true) }],
      style: {
        align: "left", lineHeight: 1.4, spaceBeforePx: 8, spaceAfterPx: 12,
        indentFirstLinePx: 0, indentLeftPx: 0, namedStyle: "tocTitle",
      },
    },
  ];
  for (const b of doc.blocks) {
    if (b.kind !== "paragraph") continue;
    const level = headingLevel(b);
    if (level === null) continue;
    const text = textOfRuns(b.runs).replace(/\v/g, " ").trim();
    if (text.length === 0) continue;
    out.push({
      kind: "paragraph",
      id: freshBlockId(),
      revision: 0,
      runs: [{ text, style: char(level === 1 ? 14 : 13, level === 1) }],
      style: {
        align: "left", lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 2,
        indentFirstLinePx: 0, indentLeftPx: (level - 1) * 20,
        tocEntry: { targetId: b.id, level },
      },
    });
  }
  return out;
}

/** Insert a TOC at the caret, or REGENERATE the existing one (Word's F9). */
export function insertTocCmd(): Command {
  return (state) => {
    const fresh = buildTocParagraphs(state.doc);
    if (fresh.length <= 1) return null; // no headings -> nothing to list
    const existing = tocEntryBlocks(state.doc);
    const ops: Op[] = [];
    let insertIndex: number;
    if (existing.length > 0) {
      insertIndex = state.doc.blocks.findIndex((b) => b.id === existing[0]!.id);
      for (const b of existing) ops.push({ type: "removeBlock", blockId: b.id });
    } else {
      const sel = state.selection;
      const loc = sel ? locateParagraph(state.doc, sel.focus.blockId) : null;
      insertIndex =
        loc && (loc.kind === "top" || (loc.kind === "cell" && loc.where === "body")) ? loc.bi : 0;
    }
    fresh.forEach((p, k) => ops.push({ type: "insertBlock", index: insertIndex + k, block: p }));
    const first = fresh[0]!;
    return tr(ops, caret(first.id, 0), "command");
  };
}

// ---------------------------------------------------------------------------
// Content controls (OOXML w:sdt) — contiguous runs sharing a CharStyle.sdtId
// form one inline control; properties live in Document.sdts. Placeholder text
// is gray and replaced WHOLE on first input (Word).

export interface SdtRange {
  blockId: string;
  start: number;
  end: number;
}

/** Every contiguous run span carrying this sdtId, document order. */
export function findSdtRanges(doc: EditorState["doc"], id: string): SdtRange[] {
  const out: SdtRange[] = [];
  for (const p of paragraphsOf(doc)) {
    let off = 0;
    let open: SdtRange | null = null;
    for (const r of p.runs) {
      if (r.style.sdtId === id) {
        if (open) open.end = off + r.text.length;
        else open = { blockId: p.id, start: off, end: off + r.text.length };
      } else if (open) {
        out.push(open);
        open = null;
      }
      off += r.text.length;
    }
    if (open) out.push(open);
  }
  return out;
}

/** The control containing a position (the run at the caret, either side). */
export function sdtAtPosition(doc: EditorState["doc"], pos: DocPosition): string | null {
  const block = blockById(doc, pos.blockId);
  if (!block) return null;
  let off = 0;
  let before: string | null = null;
  let after: string | null = null;
  for (const r of block.runs) {
    const end = off + r.text.length;
    if (pos.offset > off && pos.offset <= end) before = r.style.sdtId ?? null;
    if (pos.offset >= off && pos.offset < end) after = r.style.sdtId ?? null;
    off = end;
  }
  return after ?? before;
}

const PLACEHOLDER_TEXT: Record<SdtType, string> = {
  richText: "Click or tap here to enter text.",
  plainText: "Click or tap here to enter text.",
  checkbox: "☐",
  dropDown: "Choose an item.",
  comboBox: "Choose an item.",
  date: "Click or tap to enter a date.",
};

const SDT_PLACEHOLDER_COLOR = "#767676";

let sdtCounter = 0;
const freshSdtId = (): string => `sdt_${Date.now().toString(36)}_${sdtCounter++}`;

/** Insert a content control at the caret (placeholder content) or wrap the
 *  selected text (becomes the control's content). */
export function insertContentControl(type: SdtType, props: Partial<SdtProps> = {}): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel) return null;
    const id = freshSdtId();
    const full: SdtProps = { type, ...props };
    if (type === "checkbox") full.checked = props.checked ?? false;
    const ops: Op[] = [];
    if (!isCollapsed(sel) && type !== "checkbox") {
      // Wrap the selection: its text becomes real (non-placeholder) content.
      const [from, to] = orderedRange(state, sel);
      if (from.blockId !== to.blockId) return null; // inline controls only
      const block = blockById(state.doc, from.blockId);
      if (!block) return null;
      ops.push({ type: "setSdtProps", id, props: full });
      ops.push({
        type: "setRuns",
        blockId: from.blockId,
        runs: applyStylePatchToRuns(block.runs, from.offset, to.offset, { sdtId: id }),
      });
      return tr(ops, sel, "command");
    }
    // Collapsed caret: insert gray placeholder content (checkbox: the glyph).
    const at = isCollapsed(sel) ? sel.focus : orderedRange(state, sel)[0];
    const block = blockById(state.doc, at.blockId);
    if (!block) return null;
    const inherited = styleAtRuns(block.runs, at.offset) ?? DEFAULT_CELL_CHAR;
    const isCheckbox = type === "checkbox";
    if (!isCheckbox) full.placeholder = true;
    ops.push({ type: "setSdtProps", id, props: full });
    ops.push({
      type: "insertText",
      at,
      text: PLACEHOLDER_TEXT[type],
      style: {
        ...inherited,
        color: isCheckbox ? inherited.color : SDT_PLACEHOLDER_COLOR,
        sdtId: id,
        link: undefined,
        footnoteRef: undefined,
      },
    });
    const len = PLACEHOLDER_TEXT[type].length;
    return tr(
      ops,
      {
        anchor: { blockId: at.blockId, offset: at.offset },
        focus: { blockId: at.blockId, offset: at.offset + len },
      },
      "command",
    );
  };
}

/** Replace a control's content with `text` (dropdown pick, date pick, first
 *  typed character into a placeholder). Clears the placeholder flag. */
export function setSdtContent(id: string, text: string, patch: Partial<SdtProps> = {}): Command {
  return (state) => {
    const props = state.doc.sdts?.[id];
    const range = findSdtRanges(state.doc, id)[0];
    if (!props || !range) return null;
    const block = blockById(state.doc, range.blockId);
    if (!block) return null;
    const base = styleAtRuns(block.runs, range.start + 1) ?? DEFAULT_CELL_CHAR;
    const ops: Op[] = [
      { type: "deleteRange", blockId: range.blockId, start: range.start, end: range.end },
      {
        type: "insertText",
        at: { blockId: range.blockId, offset: range.start },
        text,
        style: { ...base, color: props.placeholder ? "#202124" : base.color, sdtId: id },
      },
      { type: "setSdtProps", id, props: { ...props, ...patch, placeholder: false } },
    ];
    return tr(ops, caret(range.blockId, range.start + text.length), "command");
  };
}

export function toggleSdtCheckbox(id: string): Command {
  return (state) => {
    const props = state.doc.sdts?.[id];
    if (!props || props.type !== "checkbox" || props.lockContent) return null;
    const checked = !(props.checked ?? false);
    return setSdtContent(id, checked ? "☒" : "☐", { checked })(state);
  };
}

/** Replace a control's whole content with an edited fragment (the inspector's
 *  Save). Reuses insertFragment over the control's span so single- and
 *  multi-paragraph (body) controls both work; every inserted run is re-tagged
 *  with the sdtId so the control's membership survives. Returns null when the
 *  span can't be structurally replaced (e.g. multi-paragraph inside a cell). */
export function replaceSdtContent(id: string, fragment: DocFragment): Command {
  return (state) => {
    const props = state.doc.sdts?.[id];
    const ranges = findSdtRanges(state.doc, id);
    if (!props || ranges.length === 0) return null;
    const first = ranges[0]!;
    const last = ranges[ranges.length - 1]!;
    const tagged: DocFragment = {
      inline: fragment.inline,
      blocks: fragment.blocks.map((b) => ({
        style: b.style,
        runs: b.runs.map((r) => ({ text: r.text, style: { ...r.style, sdtId: id } })),
      })),
    };
    const spanState: EditorState = {
      ...state,
      selection: {
        anchor: { blockId: first.blockId, offset: first.start },
        focus: { blockId: last.blockId, offset: last.end },
      },
    };
    const inner = insertFragment(tagged)(spanState);
    if (!inner) return null;
    // Editing real content clears the gray placeholder flag.
    const ops = props.placeholder
      ? [...inner.ops, { type: "setSdtProps" as const, id, props: { ...props, placeholder: false } }]
      : inner.ops;
    return { ops, selectionAfter: inner.selectionAfter, origin: "command" };
  };
}

/** Replace a BLOCK-LEVEL control's whole block span (paragraphs, images,
 *  tables, blank lines) with a new block list — the inspector's Save for
 *  controls whose content holds objects that insertFragment can't round-trip.
 *  Paragraph runs are re-tagged with the sdtId (so the control survives) and
 *  given fresh ids; image/table blocks are inserted verbatim. The span is
 *  forced to begin and end with a tagged paragraph so findSdtRanges recovers
 *  the full extent next time (objects at the very edges would otherwise be
 *  clipped out of the control). */
export function replaceSdtBlockSpan(id: string, blocks: Block[]): Command {
  return (state) => {
    const props = state.doc.sdts?.[id];
    const ranges = findSdtRanges(state.doc, id);
    if (!props || ranges.length === 0) return null;
    const first = containerOf(state.doc, ranges[0]!.blockId);
    const last = containerOf(state.doc, ranges[ranges.length - 1]!.blockId);
    if (!first || !last || first.where !== last.where) return null;
    const where = first.where;
    const span = containerBlocks(state.doc, where);
    const lo = first.index;
    const hi = last.index;
    if (lo > hi || lo < 0 || hi >= span.length) return null;

    const baseStyle: CharStyle =
      styleAtRuns((span[lo] as Paragraph).runs ?? [], ranges[0]!.start + 1) ?? DEFAULT_CELL_CHAR;
    const taggedPara = (p: Paragraph): Paragraph => ({
      kind: "paragraph",
      id: freshBlockId(),
      revision: 0,
      style: p.style,
      runs: (p.runs.length > 0 ? p.runs : [{ text: "", style: baseStyle }]).map((r) => ({
        text: r.text,
        style: { ...r.style, sdtId: id },
      })),
    });
    const emptyTagged = (): Paragraph =>
      taggedPara({ kind: "paragraph", id: "", revision: 0, runs: [], style: (span[lo] as Paragraph).style });

    const fresh: Block[] = blocks.map((b) =>
      b.kind === "paragraph" ? taggedPara(b) : b,
    );
    // Guarantee tagged-paragraph bookends.
    if (fresh.length === 0 || fresh[0]!.kind !== "paragraph") fresh.unshift(emptyTagged());
    if (fresh[fresh.length - 1]!.kind !== "paragraph") fresh.push(emptyTagged());

    const ops: Op[] = [];
    for (let i = hi; i >= lo; i--) ops.push({ type: "removeBlock", blockId: span[i]!.id });
    fresh.forEach((b, k) => ops.push({ type: "insertBlock", index: lo + k, block: b, where }));
    if (props.placeholder) ops.push({ type: "setSdtProps", id, props: { ...props, placeholder: false } });

    const firstPara = fresh.find((b) => b.kind === "paragraph")!;
    return { ops, selectionAfter: caret(firstPara.id, 0), origin: "command" };
  };
}

/** Replace a CELL-HOSTED control's content in place, one tagged range at a time.
 *  Used when the control's ranges span multiple table cells (insertFragment can't
 *  replace a cross-cell selection). `runsPerRange[i]` is the new content for the
 *  i-th range in findSdtRanges order; a missing entry leaves that range as-is.
 *  Runs are re-tagged with the sdtId; the table grid (cells, rows) is untouched. */
export function replaceSdtCellContent(id: string, runsPerRange: Run[][]): Command {
  return (state) => {
    const props = state.doc.sdts?.[id];
    const ranges = findSdtRanges(state.doc, id);
    if (!props || ranges.length === 0) return null;

    // Group by block; within a block rewrite high→low so earlier offsets stay valid.
    const byBlock = new Map<string, { idx: number; r: SdtRange }[]>();
    ranges.forEach((r, idx) => {
      const list = byBlock.get(r.blockId) ?? [];
      list.push({ idx, r });
      byBlock.set(r.blockId, list);
    });
    const ops: Op[] = [];
    for (const [blockId, list] of byBlock) {
      const block = blockById(state.doc, blockId);
      if (!block) continue;
      list.sort((a, b) => b.r.start - a.r.start);
      for (const { idx, r } of list) {
        const provided = runsPerRange[idx];
        if (provided === undefined) continue; // untouched range
        const base = styleAtRuns(block.runs, r.start + 1) ?? styleAtRuns(block.runs, r.start) ?? DEFAULT_CELL_CHAR;
        // The inspector has no font UI, so typography (family/size/color) stays
        // the cell's own; only the inline toggles the user can flip (B/I/U via
        // contentEditable, links) ride over from the edited runs.
        let runs: Run[] = provided.map((rn) => ({
          text: rn.text,
          style: {
            ...base,
            bold: rn.style.bold,
            italic: rn.style.italic,
            underline: rn.style.underline,
            strikethrough: rn.style.strikethrough,
            verticalAlign: rn.style.verticalAlign,
            highlightColor: rn.style.highlightColor,
            link: rn.style.link,
            sdtId: id,
          },
        }));
        if (runs.length === 0 || runs.every((rn) => rn.text.length === 0)) {
          runs = [{ text: "", style: { ...base, sdtId: id } }];
        }
        ops.push({ type: "deleteRange", blockId, start: r.start, end: r.end });
        ops.push({ type: "insertRuns", at: { blockId, offset: r.start }, runs });
      }
    }
    if (ops.length === 0) return null;
    if (props.placeholder) ops.push({ type: "setSdtProps", id, props: { ...props, placeholder: false } });
    const f = ranges[0]!;
    return { ops, selectionAfter: caret(f.blockId, f.start), origin: "command" };
  };
}

/** Remove the control: strip the run markers; optionally delete its content. */
export function removeContentControl(id: string, deleteContents: boolean): Command {
  return (state) => {
    const props = state.doc.sdts?.[id];
    if (!props || props.lockControl) return null;
    const ranges = findSdtRanges(state.doc, id);
    const ops: Op[] = [];
    // Back-to-front so earlier offsets stay valid when deleting.
    for (const r of [...ranges].reverse()) {
      if (deleteContents) {
        ops.push({ type: "deleteRange", blockId: r.blockId, start: r.start, end: r.end });
      } else {
        const block = blockById(state.doc, r.blockId)!;
        ops.push({
          type: "setRuns",
          blockId: r.blockId,
          runs: applyStylePatchToRuns(block.runs, r.start, r.end, { sdtId: undefined }),
        });
      }
    }
    ops.push({ type: "setSdtProps", id, props: null });
    const first = ranges[0];
    return tr(
      ops,
      first ? caret(first.blockId, first.start) : state.selection,
      "command",
    );
  };
}

// ---------------------------------------------------------------------------
// Footnotes — the ref run's TEXT is its number; inserting renumbers every
// later ref in the same transaction, so model text and the page-bottom note
// markers (which echo the ref text) can never disagree.

export function insertFootnoteCmd(): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel || !isCollapsed(sel)) return null;
    const at = sel.focus;
    const loc = locateParagraph(state.doc, at.blockId);
    if (loc?.kind !== "top") return null; // body top-level paragraphs only (v1)

    // Document-ordered refs with their host paragraph + run index.
    interface RefAt {
      para: Paragraph;
      runIdx: number;
      offset: number;
    }
    const refs: RefAt[] = [];
    for (const b of state.doc.blocks) {
      if (b.kind !== "paragraph") continue;
      let off = 0;
      b.runs.forEach((r, runIdx) => {
        if (r.style.footnoteRef) refs.push({ para: b, runIdx, offset: off });
        off += r.text.length;
      });
    }
    const orderOf = (blockId: string): number => state.doc.blocks.findIndex((b) => b.id === blockId);
    const caretOrder = orderOf(at.blockId);
    let idx = 0;
    while (idx < refs.length) {
      const r = refs[idx]!;
      const ro = orderOf(r.para.id);
      if (ro < caretOrder || (ro === caretOrder && r.offset < at.offset)) idx++;
      else break;
    }
    const newNumber = idx + 1;

    const ops: Op[] = [];
    // Renumber subsequent refs (one setRuns per affected paragraph).
    const renumber = new Map<string, Map<number, string>>();
    for (let k = idx; k < refs.length; k++) {
      const r = refs[k]!;
      let m = renumber.get(r.para.id);
      if (!m) {
        m = new Map();
        renumber.set(r.para.id, m);
      }
      m.set(r.runIdx, String(k + 2));
    }
    for (const [paraId, m] of renumber) {
      const p = state.doc.blocks.find((b) => b.id === paraId) as Paragraph;
      const runs = p.runs.map((r, i) => (m.has(i) ? { ...r, text: m.get(i)! } : r));
      ops.push({ type: "setRuns", blockId: paraId, runs });
    }

    const noteId = `fn_${freshBlockId()}`;
    const block = blockById(state.doc, at.blockId)!;
    const inherited = styleAtRuns(block.runs, at.offset) ?? DEFAULT_CELL_CHAR;
    ops.push({
      type: "insertText",
      at,
      text: String(newNumber),
      style: { ...inherited, verticalAlign: "super", footnoteRef: noteId, link: undefined },
    });
    const notePara: Paragraph = {
      kind: "paragraph",
      id: freshBlockId(),
      revision: 0,
      runs: [{ text: "", style: { ...inherited, fontSizePx: 12, verticalAlign: undefined, footnoteRef: undefined, link: undefined } }],
      style: { align: "left", lineHeight: 1.4, spaceBeforePx: 0, spaceAfterPx: 2, indentFirstLinePx: 0, indentLeftPx: 0 },
    };
    ops.push({ type: "setFootnote", noteId, paras: [notePara] });
    // Word drops the caret into the new note for immediate typing.
    return tr(ops, caret(notePara.id, 0), "command");
  };
}

/** Ctrl+Enter: split at the caret; the new paragraph starts a fresh page. */
export function insertPageBreak(): Command {
  return (state) => {
    const base = withSelectionDeleted(state);
    if (!base) return null;
    if (isInCell(state.doc, base.at.blockId)) return null; // no page breaks inside cells
    const newBlockId = freshBlockId();
    base.ops.push({ type: "splitParagraph", at: base.at, newBlockId });
    base.ops.push({ type: "setParaStyle", blockId: newBlockId, patch: { pageBreakBefore: true } });
    return tr(base.ops, caret(newBlockId, 0), "command");
  };
}

/** Ctrl+Shift+Enter: split; the new paragraph starts the next newspaper column
 *  (a layout no-op in single-column sections, exactly like Word). */
export function insertColumnBreak(): Command {
  return (state) => {
    const base = withSelectionDeleted(state);
    if (!base) return null;
    if (isInCell(state.doc, base.at.blockId)) return null;
    const newBlockId = freshBlockId();
    base.ops.push({ type: "splitParagraph", at: base.at, newBlockId });
    base.ops.push({ type: "setParaStyle", blockId: newBlockId, patch: { columnBreakBefore: true } });
    return tr(base.ops, caret(newBlockId, 0), "command");
  };
}

export function toggleCharStyle(
  key: "bold" | "italic" | "underline" | "strikethrough",
): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel || isCollapsed(sel)) return null; // pending typing-style: milestone 5
    const [from, to] = orderedRange(state, sel);
    const blocks = paragraphsOf(state.doc);
    const fi = blockIndexOf(state.doc, from.blockId);
    const ti = blockIndexOf(state.doc, to.blockId);

    // Toggle semantics: if every covered character already has the style, remove
    // it; otherwise apply it (Word behavior).
    const segments: { block: Paragraph; start: number; end: number }[] = [];
    for (let i = fi; i <= ti; i++) {
      const block = blocks[i]!;
      const len = textOfRuns(block.runs).length;
      const start = i === fi ? from.offset : 0;
      const end = i === ti ? to.offset : len;
      if (end > start) segments.push({ block, start, end });
    }
    if (segments.length === 0) return null;
    const allSet = segments.every((s) =>
      sliceRuns(s.block.runs, s.start, s.end).every((r) => r.style[key]),
    );
    const patch: Partial<CharStyle> = { [key]: !allSet };

    const ops: Op[] = segments.map((s) => ({
      type: "setRuns",
      blockId: s.block.id,
      runs: applyStylePatchToRuns(s.block.runs, s.start, s.end, patch),
    }));
    return tr(ops, sel, "command"); // selection is preserved (lengths unchanged)
  };
}

/** Insert an atomic block (image/table) at a collapsed top-level caret: split
 *  the paragraph there and slot the block between head and tail. */
function insertBlockAtCaret(state: EditorState, makeBlock: () => Block): Transaction | null {
  const sel = state.selection;
  if (!sel || !isCollapsed(sel)) return null;
  const at = sel.focus;
  const bi = state.doc.blocks.findIndex((b) => b.id === at.blockId);
  if (bi < 0) return null; // inside a cell
  const tailId = freshBlockId();
  const block = makeBlock();
  const ops: Op[] = [
    { type: "splitParagraph", at, newBlockId: tailId },
    { type: "insertBlock", index: bi + 1, block },
  ];
  return tr(ops, caret(tailId, 0), "command");
}

export function insertImage(src: string, widthPx: number, heightPx: number): Command {
  return (state) =>
    insertBlockAtCaret(state, () => ({
      kind: "image",
      id: freshBlockId(),
      revision: 0,
      src,
      widthPx,
      heightPx,
      align: "center",
    }));
}

export function insertTable(rows: number, cols: number): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel) return null;
    const block = blockById(state.doc, sel.focus.blockId);
    const style = (block ? styleAtRuns(block.runs, sel.focus.offset) : undefined) ?? {
      fontFamily: "Georgia, serif",
      fontSizePx: 16,
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      color: "#202124",
    };
    const cellPara = (): Paragraph => ({
      kind: "paragraph",
      id: freshBlockId(),
      revision: 0,
      runs: [{ text: "", style }],
      style: {
        align: "left",
        lineHeight: 1.4,
        spaceBeforePx: 0,
        spaceAfterPx: 0,
        indentFirstLinePx: 0,
        indentLeftPx: 0,
      },
    });
    return insertBlockAtCaret(state, () => ({
      kind: "table",
      id: freshBlockId(),
      revision: 0,
      rows: Array.from({ length: rows }, () => ({
        cells: Array.from({ length: cols }, () => ({ id: freshBlockId(), blocks: [cellPara()] })),
      })),
    }));
  };
}

// ---------------------------------------------------------------------------
// Table structure commands — all act on the table cell containing the caret.

interface CellContext {
  table: TableBlock;
  bi: number;
  ri: number;
  ci: number;
}

function cellContext(state: EditorState): CellContext | null {
  const sel = state.selection;
  if (!sel) return null;
  const loc = locateParagraph(state.doc, sel.focus.blockId);
  if (loc?.kind !== "cell") return null;
  return {
    table: containerBlocks(state.doc, loc.where)[loc.bi] as TableBlock,
    bi: loc.bi,
    ri: loc.ri,
    ci: loc.ci,
  };
}

/** First caret-capable paragraph of a cell (cells may lead with images now). */
const firstCellPara = (cell: TableCell | undefined): Paragraph | undefined =>
  cell?.blocks.find((b): b is Paragraph => b.kind === "paragraph");

const DEFAULT_CELL_CHAR: CharStyle = {
  fontFamily: "Georgia, serif",
  fontSizePx: 14,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#202124",
};

function emptyCellPara(proto: Paragraph | undefined): Paragraph {
  return {
    kind: "paragraph",
    id: freshBlockId(),
    revision: 0,
    runs: [{ text: "", style: proto?.runs[0]?.style ?? DEFAULT_CELL_CHAR }],
    style: proto
      ? { ...proto.style }
      : { align: "left", lineHeight: 1.35, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 },
  };
}

export function insertTableRowCmd(side: "above" | "below"): Command {
  return (state) => {
    const ctx = cellContext(state);
    if (!ctx) return null;
    const protoRow = ctx.table.rows[ctx.ri]!;
    const row: TableRow = {
      cells: protoRow.cells.map((c) => ({ id: freshBlockId(), blocks: [emptyCellPara(firstCellPara(c))] })),
    };
    const rowIndex = side === "above" ? ctx.ri : ctx.ri + 1;
    const caretCell = row.cells[Math.min(ctx.ci, row.cells.length - 1)]!;
    return tr(
      [{ type: "insertTableRow", tableId: ctx.table.id, rowIndex, row }],
      caret(caretCell.blocks[0]!.id, 0),
      "command",
    );
  };
}

export function insertTableColumnCmd(side: "left" | "right"): Command {
  return (state) => {
    const ctx = cellContext(state);
    if (!ctx) return null;
    const colIndex = side === "left" ? ctx.ci : ctx.ci + 1;
    const cells = ctx.table.rows.map((row) => ({
      id: freshBlockId(),
      blocks: [emptyCellPara(firstCellPara(row.cells[ctx.ci]))],
    }));
    const caretCell = cells[ctx.ri]!;
    return tr(
      [{ type: "insertTableColumn", tableId: ctx.table.id, colIndex, cells }],
      caret(caretCell.blocks[0]!.id, 0),
      "command",
    );
  };
}

/** Caret target for after a whole table (or its last row/col) disappears. */
function caretAfterTable(state: EditorState, bi: number): DocSelection | null {
  const neighbor = state.doc.blocks[bi + 1] ?? state.doc.blocks[bi - 1];
  if (neighbor && neighbor.kind === "paragraph") return caret(neighbor.id, 0);
  return null;
}

export function deleteTableRowCmd(): Command {
  return (state) => {
    const ctx = cellContext(state);
    if (!ctx) return null;
    if (ctx.table.rows.length <= 1) return deleteTableCmd()(state);
    const targetRow = ctx.table.rows[ctx.ri === 0 ? 1 : ctx.ri - 1]!;
    const caretPara =
      firstCellPara(targetRow.cells[Math.min(ctx.ci, targetRow.cells.length - 1)]) ??
      firstCellPara(targetRow.cells.find((c) => firstCellPara(c)));
    return tr(
      [{ type: "removeTableRow", tableId: ctx.table.id, rowIndex: ctx.ri }],
      caretPara ? caret(caretPara.id, 0) : state.selection,
      "command",
    );
  };
}

export function deleteTableColumnCmd(): Command {
  return (state) => {
    const ctx = cellContext(state);
    if (!ctx) return null;
    const colCount = Math.max(...ctx.table.rows.map((r) => r.cells.length));
    if (colCount <= 1) return deleteTableCmd()(state);
    const row = ctx.table.rows[ctx.ri]!;
    const caretPara = firstCellPara(row.cells[ctx.ci === 0 ? 1 : ctx.ci - 1]);
    return tr(
      [{ type: "removeTableColumn", tableId: ctx.table.id, colIndex: ctx.ci }],
      caretPara ? caret(caretPara.id, 0) : state.selection,
      "command",
    );
  };
}

/** Merge horizontally: selection anchor and focus must sit in different cells
 *  of the SAME row. Content concatenates; the merged cell spans the columns.
 *  Invertible in one step via the whole-row setTableRow op. */
export function mergeCellsCmd(): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel) return null;
    const a = locateParagraph(state.doc, sel.anchor.blockId);
    const f = locateParagraph(state.doc, sel.focus.blockId);
    if (a?.kind !== "cell" || f?.kind !== "cell") return null;
    if (a.bi !== f.bi || a.ri !== f.ri || a.ci === f.ci) return null;
    const table = state.doc.blocks[a.bi] as TableBlock;
    const row = table.rows[a.ri]!;
    const ciStart = Math.min(a.ci, f.ci);
    const ciEnd = Math.max(a.ci, f.ci);
    const covered = row.cells.slice(ciStart, ciEnd + 1);
    const spanSum = covered.reduce((s, c) => s + (c.colSpan ?? 1), 0);
    const mergedCell: TableCell = {
      id: covered[0]!.id,
      blocks: covered.flatMap((c) => c.blocks),
      colSpan: spanSum,
    };
    const cells = [...row.cells.slice(0, ciStart), mergedCell, ...row.cells.slice(ciEnd + 1)];
    const firstPara = firstCellPara(mergedCell);
    return tr(
      [{ type: "setTableRow", tableId: table.id, rowIndex: a.ri, row: { cells } }],
      firstPara ? caret(firstPara.id, 0) : null,
      "command",
    );
  };
}

/** Split a merged cell back into single-column cells (content stays in the first). */
export function unmergeCellCmd(): Command {
  return (state) => {
    const ctx = cellContext(state);
    if (!ctx) return null;
    const row = ctx.table.rows[ctx.ri]!;
    const cell = row.cells[ctx.ci]!;
    const span = cell.colSpan ?? 1;
    if (span <= 1) return null;
    const restored: TableCell[] = [{ id: cell.id, blocks: cell.blocks }];
    for (let k = 1; k < span; k++) {
      restored.push({ id: freshBlockId(), blocks: [emptyCellPara(firstCellPara(cell))] });
    }
    const cells = [...row.cells.slice(0, ctx.ci), ...restored, ...row.cells.slice(ctx.ci + 1)];
    const caretPara = firstCellPara(cell);
    return tr(
      [{ type: "setTableRow", tableId: ctx.table.id, rowIndex: ctx.ri, row: { cells } }],
      caretPara ? caret(caretPara.id, 0) : null,
      "command",
    );
  };
}

export function deleteTableCmd(): Command {
  return (state) => {
    const ctx = cellContext(state);
    if (!ctx) return null;
    return tr(
      [{ type: "removeBlock", blockId: ctx.table.id }],
      caretAfterTable(state, ctx.bi),
      "command",
    );
  };
}

// ---------------------------------------------------------------------------
// Image commands (target = the SELECTED object, passed in by the wiring layer)

export function setImageProps(
  blockId: string,
  patch: Partial<Pick<ImageBlock, "widthPx" | "heightPx" | "align" | "wrap">>,
  origin: TransactionOrigin = "command",
): Command {
  return (state) => {
    const exists = state.doc.blocks.some((b) => b.id === blockId && b.kind === "image");
    if (!exists) return null;
    return tr([{ type: "setImageProps", blockId, patch }], state.selection, origin);
  };
}

export function deleteImage(blockId: string): Command {
  return (state) => {
    const loc = locateImage(state.doc, blockId);
    if (!loc) return null;
    if (loc.kind === "cell") {
      const table = state.doc.blocks[loc.bi] as TableBlock;
      const row = table.rows[loc.ri]!;
      const cell = row.cells[loc.ci]!;
      const cells = row.cells.slice();
      cells[loc.ci] = { ...cell, blocks: cell.blocks.filter((b) => b.id !== blockId) };
      const caretPara = firstCellPara(cells[loc.ci]);
      return tr(
        [{ type: "setTableRow", tableId: table.id, rowIndex: loc.ri, row: { cells } }],
        caretPara ? caret(caretPara.id, 0) : state.selection,
        "command",
      );
    }
    const bi = state.doc.blocks.findIndex((b) => b.id === blockId);
    return tr([{ type: "removeBlock", blockId }], caretAfterTable(state, bi), "command");
  };
}

/** Insert an image INSIDE the caret's table cell (after the caret paragraph). */
export function insertImageInCell(src: string, widthPx: number, heightPx: number): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel || !isCollapsed(sel)) return null;
    const loc = locateParagraph(state.doc, sel.focus.blockId);
    if (loc?.kind !== "cell") return null;
    const table = containerBlocks(state.doc, loc.where)[loc.bi] as TableBlock;
    const row = table.rows[loc.ri]!;
    const cell = row.cells[loc.ci]!;
    const img: Block = {
      kind: "image",
      id: freshBlockId(),
      revision: 0,
      src,
      widthPx,
      heightPx,
      align: "center",
    };
    const blocks = cell.blocks.slice();
    blocks.splice(loc.pi + 1, 0, img);
    const cells = row.cells.slice();
    cells[loc.ci] = { ...cell, blocks };
    return tr(
      [{ type: "setTableRow", tableId: table.id, rowIndex: loc.ri, row: { cells } }],
      sel,
      "command",
    );
  };
}

export function setTableColFractionsCmd(
  tableId: string,
  fractions: number[],
  origin: TransactionOrigin = "command",
): Command {
  return (state) =>
    tr([{ type: "setTableColFractions", blockId: tableId, fractions }], state.selection, origin);
}

// ---------------------------------------------------------------------------
// Direct formatting (absolute patches — power the font/size/spacing controls)

/** Paragraphs covered by the selection, with the run-range inside each. */
function selectedSegments(state: EditorState): { block: Paragraph; start: number; end: number }[] {
  const sel = state.selection;
  if (!sel) return [];
  const [from, to] = orderedRange(state, sel);
  const blocks = paragraphsOf(state.doc);
  const fi = blockIndexOf(state.doc, from.blockId);
  const ti = blockIndexOf(state.doc, to.blockId);
  if (fi < 0 || ti < 0) return [];
  const out: { block: Paragraph; start: number; end: number }[] = [];
  for (let i = fi; i <= ti; i++) {
    const block = blocks[i]!;
    const len = textOfRuns(block.runs).length;
    out.push({
      block,
      start: i === fi ? from.offset : 0,
      end: i === ti ? to.offset : len,
    });
  }
  return out;
}

/** Absolute character-style patch over the selection (font, size, color...).
 *  Collapsed carets are handled by the wiring layer via pendingStyle. */
export function setCharStyle(patch: Partial<CharStyle>): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel || isCollapsed(sel)) return null;
    const segments = selectedSegments(state).filter((s) => s.end > s.start);
    if (segments.length === 0) return null;
    const ops: Op[] = segments.map((s) => ({
      type: "setRuns",
      blockId: s.block.id,
      runs: applyStylePatchToRuns(s.block.runs, s.start, s.end, patch),
    }));
    return tr(ops, sel, "command");
  };
}

/** Toggle one char-style VALUE over the selection: if every covered character
 *  already has it, clear it; otherwise set it (highlight, sub/super, ...). */
function toggleCharValue<K extends keyof CharStyle>(key: K, value: CharStyle[K]): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel || isCollapsed(sel)) return null;
    const segments = selectedSegments(state).filter((s) => s.end > s.start);
    if (segments.length === 0) return null;
    const allSet = segments.every((s) =>
      sliceRuns(s.block.runs, s.start, s.end).every((r) => r.style[key] === value),
    );
    const patch = { [key]: allSet ? undefined : value } as Partial<CharStyle>;
    const ops: Op[] = segments.map((s) => ({
      type: "setRuns",
      blockId: s.block.id,
      runs: applyStylePatchToRuns(s.block.runs, s.start, s.end, patch),
    }));
    return tr(ops, sel, "command");
  };
}

export const toggleHighlight = (color = "#ffeb3b"): Command => toggleCharValue("highlightColor", color);
export const toggleVerticalAlign = (which: "sub" | "super"): Command =>
  toggleCharValue("verticalAlign", which);

/** Set (url) or remove (null) a hyperlink. A collapsed caret expands to the
 *  word under it. */
export function setLinkCmd(url: string | null): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel) return null;
    const patch: Partial<CharStyle> = { link: url ?? undefined };
    if (!isCollapsed(sel)) return setCharStyle(patch)(state);
    const block = blockById(state.doc, sel.focus.blockId);
    if (!block) return null;
    const text = textOfRuns(block.runs);
    let start = 0;
    let end = 0;
    for (const s of words.segment(text)) {
      const sEnd = s.index + s.segment.length;
      if (s.isWordLike && sel.focus.offset >= s.index && sel.focus.offset <= sEnd) {
        start = s.index;
        end = sEnd;
        break;
      }
    }
    if (end <= start) return null;
    return tr(
      [{ type: "setRuns", blockId: block.id, runs: applyStylePatchToRuns(block.runs, start, end, patch) }],
      sel,
      "command",
    );
  };
}

/** AutoCorrect helper: replace N chars before the collapsed caret + insert. */
export function replaceBackAndInsert(deleteCount: number, text: string): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel || !isCollapsed(sel) || sel.focus.offset < deleteCount) return null;
    const { blockId, offset } = sel.focus;
    const at = offset - deleteCount;
    return tr(
      [
        { type: "deleteRange", blockId, start: at, end: offset },
        { type: "insertText", at: { blockId, offset: at }, text },
      ],
      caret(blockId, at + text.length),
      "typing", // coalesces with the keystroke that triggered it
    );
  };
}

/** Paragraph-property patch (line spacing, indents...) over covered paragraphs. */
export function setParaProps(patch: Partial<ParaStyle>): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel) return null;
    const ops: Op[] = selectedSegments(state).map((s) => ({
      type: "setParaStyle",
      blockId: s.block.id,
      patch,
    }));
    return ops.length ? tr(ops, sel, "command") : null;
  };
}

// ---------------------------------------------------------------------------
// Named styles (Word's style gallery) — apply / update-to-match / create

import {
  defaultStylesheet,
  paragraphsWithStyle,
  resolveStyle,
  styleById,
  type NamedStyle,
  type Stylesheet,
} from "../model/stylesheet";

const sheetOf = (state: EditorState): Stylesheet => state.doc.stylesheet ?? defaultStylesheet();

/** Restyle one paragraph to a style's RESOLVED templates: para props the style
 *  defines are set (plus the namedStyle ref); char fields the style defines are
 *  patched onto every run — direct formatting on other fields survives (Word). */
function restyleOps(block: Paragraph, char: Partial<CharStyle>, para: Partial<ParaStyle>, styleId?: string): Op[] {
  const ops: Op[] = [];
  const paraPatch: Partial<ParaStyle> = { ...para };
  if (styleId !== undefined) paraPatch.namedStyle = styleId;
  ops.push({ type: "setParaStyle", blockId: block.id, patch: paraPatch });
  if (Object.keys(char).length > 0) {
    const len = textOfRuns(block.runs).length;
    const runs =
      len === 0
        ? [{ text: "", style: { ...(block.runs[0]?.style ?? DEFAULT_CELL_CHAR), ...char } }]
        : applyStylePatchToRuns(block.runs, 0, len, char);
    ops.push({ type: "setRuns", blockId: block.id, runs });
  }
  return ops;
}

export function applyNamedStyle(styleId: string): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel) return null;
    const sheet = sheetOf(state);
    if (!styleById(sheet, styleId)) return null;
    const { char, para } = resolveStyle(sheet, styleId);
    const ops: Op[] = [];
    const seen = new Set<string>();
    for (const s of selectedSegments(state)) {
      if (seen.has(s.block.id)) continue;
      seen.add(s.block.id);
      ops.push(...restyleOps(s.block, char, para, styleId));
    }
    return ops.length ? tr(ops, sel, "command") : null;
  };
}

/** Word's "Update <style> to Match Selection": redefine the style from the
 *  caret formatting, then re-patch every paragraph referencing it. */
export function updateStyleToSelection(styleId: string): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel) return null;
    const sheet = sheetOf(state);
    const existing = styleById(sheet, styleId);
    if (!existing) return null;
    const block = blockById(state.doc, sel.focus.blockId);
    if (!block) return null;
    const char = styleAtRuns(block.runs, sel.focus.offset);
    if (!char) return null;
    const { namedStyle: _omit, ...para } = block.style;

    const updated: NamedStyle = { ...existing, char: { ...char }, para: { ...para } };
    // The style is now self-contained — drop basedOn so resolution is exact.
    delete updated.basedOn;
    const stylesheet: Stylesheet = {
      ...sheet,
      styles: sheet.styles.map((s) => (s.id === styleId ? updated : s)),
    };

    const ops: Op[] = [{ type: "setStylesheet", stylesheet }];
    for (const p of paragraphsWithStyle(state.doc, styleId)) {
      ops.push(...restyleOps(p, updated.char, updated.para));
    }
    return tr(ops, sel, "command");
  };
}

/** Create a new named style from the caret's formatting and apply it. */
export function createStyleFromSelection(name: string): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel || name.trim().length === 0) return null;
    const sheet = sheetOf(state);
    const id = name.trim().replace(/\s+/g, "");
    if (styleById(sheet, id)) return null; // no duplicates
    const block = blockById(state.doc, sel.focus.blockId);
    if (!block) return null;
    const char = styleAtRuns(block.runs, sel.focus.offset);
    if (!char) return null;
    const { namedStyle: _omit, ...para } = block.style;
    const style: NamedStyle = { id, name: name.trim(), char: { ...char }, para: { ...para } };
    const stylesheet: Stylesheet = { ...sheet, styles: [...sheet.styles, style] };
    const ops: Op[] = [
      { type: "setStylesheet", stylesheet },
      ...restyleOps(block, style.char, style.para, id),
    ];
    return tr(ops, sel, "command");
  };
}

export function setAlignment(align: ParaStyle["align"]): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel) return null;
    const [from, to] = orderedRange(state, sel);
    const fi = blockIndexOf(state.doc, from.blockId);
    const ti = blockIndexOf(state.doc, to.blockId);
    const blocks = paragraphsOf(state.doc);
    const ops: Op[] = [];
    for (let i = fi; i <= ti; i++) {
      ops.push({ type: "setParaStyle", blockId: blocks[i]!.id, patch: { align } });
    }
    return tr(ops, sel, "command");
  };
}
