// Commands: pure (state) -> Transaction | null. The keymap, the IME proxy, and
// (later) toolbar buttons all dispatch through these.

import type { Block, CellBorder, CellBorders, CharStyle, EquationBlock, FieldDef, FieldSpec, GridSlot, ImageBlock, MathEquation, ParaStyle, Paragraph, Run, RowProps, SdtProps, SdtType, ShapeBlock, ShapePreset, TableBlock, TableCell, TableGrid, TableRow, TableStyle, TocOptions, TocSwitches } from "@cw/shared";
import { bakeTableStyleRows, DEFAULT_TBL_LOOK } from "@cw/shared";
import { buildTocParagraphs, buildTocInstruction, buildInstruction, evaluateField } from "@cw/shared";
import type { BookmarkRange, DocPosition, DocSelection, GridRect } from "@cw/shared";
import { isCollapsed, BAND_CONTAINERS } from "@cw/shared";
import type { ImagePropsPatch, Op, SectionGeometry, ShapePropsPatch, TablePropsPatch } from "@cw/shared";
import { sliceRuns, applyStylePatchToRuns, mapTextInRuns, splitRunsAt, normalizeRuns, containerOf, containerBlocks, containerListOf, locateImage, locateEquation, locateShape, locateBlock, freshId } from "@cw/shared";
import { inSdt, innermostSdtId, pushSdt, removeSdt, ancestryThrough } from "@cw/shared";
import { buildTableGrid, cellsInRect, gridOriginOfCell, mergeRows, normalizeRect, rebuildRows, unmergeRows } from "@cw/shared";
import type { CellSelection } from "./state";
import {
  blockById,
  blockIndexOf,
  bodyParagraphs,
  isHiddenParagraph,
  isInCell,
  locateParagraph,
  paragraphsOf,
  prevGrapheme,
  nextGrapheme,
  styleAtRuns,
  styleOfCharAt,
  textOfBlock,
  textOfRuns,
  words,
} from "@cw/shared";
import type { DocFragment } from "../input/clipboard";
import { deletePageGroupOps, pinPageBreakBeforeOps, reorderPageGroupsOps } from "../layout/pageGroups";
import type { Command, EditorState, Transaction, TransactionOrigin } from "./state";

const caret = (blockId: string, offset: number): DocSelection => ({
  anchor: { blockId, offset },
  focus: { blockId, offset },
});

// Block/cell ids come from the shared, siteId-namespaced generator so two
// collaborating clients never mint the same id (see shared/ids). The frontend
// configures the siteId at startup (configureIds in main.ts).
const freshBlockId = (): string => freshId();

// ---------------------------------------------------------------------------
// Bookmarks (name → range; managed from the Bookmarks panel)

/** Create or move a bookmark to cover the current selection (or caret = a point
 *  bookmark). Names are unique; re-adding an existing name moves it. */
export function addBookmarkCmd(name: string): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel || name.trim().length === 0) return null;
    const [from, to] = orderedRange(state, sel);
    const range: BookmarkRange = { start: from, end: to };
    return tr([{ type: "setBookmark", name, range }], sel, "command");
  };
}

export function removeBookmarkCmd(name: string): Command {
  return (state) => {
    if (!state.doc.bookmarks?.[name]) return null;
    return tr([{ type: "setBookmark", name, range: null }], state.selection, "command");
  };
}

/** Rename a bookmark, preserving its range (delete old + set new, one undo step). */
export function renameBookmarkCmd(oldName: string, newName: string): Command {
  return (state) => {
    const range = state.doc.bookmarks?.[oldName];
    if (!range || newName.trim().length === 0 || newName === oldName || state.doc.bookmarks?.[newName]) return null;
    return tr(
      [
        { type: "setBookmark", name: oldName, range: null },
        { type: "setBookmark", name: newName, range },
      ],
      state.selection,
      "command",
    );
  };
}

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
  // Hidden (w:vanish) anchor paragraphs in the deleted span are PRESERVED — they
  // survive even Select-All → Delete. When any sit between the ends we also skip
  // the final merge so the two cleared paragraphs stay separate around them
  // (merging would absorb the hidden runs / orphan their block id).
  let hiddenBetween = false;
  for (let i = fi + 1; i < ti; i++) {
    const b = docBlocks[i]!;
    if (b.kind === "paragraph" && isHiddenParagraph(b)) {
      hiddenBetween = true;
      continue;
    }
    ops.push({ type: "removeBlock", blockId: b.id });
  }
  if (!hiddenBetween) ops.push({ type: "mergeParagraphs", firstBlockId: from.blockId });
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
    // List membership is inherited by the split (style is cloned). A block-level
    // content control wrapping this paragraph is inherited too, so Enter inside a
    // block-level SDT keeps the new paragraph in the control (carry its sdtPath).
    const splitBlock = blockById(state.doc, base.at.blockId);
    const newBlockId = freshBlockId();
    base.ops.push({
      type: "splitParagraph",
      at: base.at,
      newBlockId,
      ...(splitBlock?.sdtPath?.length ? { newSdtPath: splitBlock.sdtPath } : {}),
    });
    return tr(base.ops, caret(newBlockId, 0), "command");
  };
}

// ---------------------------------------------------------------------------
// Lists

import {
  defaultListDefinition,
  multilevelListDefinition,
  DEFAULT_BULLET_LIST_ID,
  DEFAULT_NUMBER_LIST_ID,
  DEFAULT_MULTILEVEL_LIST_ID,
  type ListDefinition,
} from "@cw/shared";

/** Paragraphs covered by the selection that can take a list: top-level body
 *  paragraphs and body table-cell paragraphs (Word numbers a list straight
 *  through cells in reading order). Bands and footnotes are excluded. */
function selectedTopLevelParagraphs(state: EditorState): Paragraph[] {
  const out: Paragraph[] = [];
  const seen = new Set<string>();
  for (const s of selectedSegments(state)) {
    if (seen.has(s.block.id)) continue;
    seen.add(s.block.id);
    const loc = locateParagraph(state.doc, s.block.id);
    if (loc?.kind === "top" || (loc?.kind === "cell" && loc.where === "body")) out.push(s.block);
  }
  return out;
}

/** Toggle the covered paragraphs into (or out of) the list `listId`, materializing
 *  its definition on first use. Shared by the bullet / numbered / multilevel buttons. */
function toggleListById(listId: string, makeDef: () => ListDefinition): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel) return null;
    const paragraphs = selectedTopLevelParagraphs(state);
    if (paragraphs.length === 0) return null;
    const allIn = paragraphs.every((p) => p.style.list?.listId === listId);
    const ops: Op[] = [];
    if (!allIn && !state.doc.lists?.[listId]) {
      ops.push({ type: "setListDefinition", listId, def: makeDef() });
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

export function toggleList(kind: "bullet" | "decimal"): Command {
  const listId = kind === "bullet" ? DEFAULT_BULLET_LIST_ID : DEFAULT_NUMBER_LIST_ID;
  return toggleListById(listId, () => defaultListDefinition(kind));
}

/** Apply a SPECIFIC list style (bullet glyph / number format) from the list
 *  pickers. Unlike toggleList this always sets the style (re)materializing its
 *  definition — picking a style never toggles the list off. */
export function applyListStyleCmd(def: ListDefinition): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel) return null;
    const paragraphs = selectedTopLevelParagraphs(state);
    if (paragraphs.length === 0) return null;
    const ops: Op[] = [{ type: "setListDefinition", listId: def.id, def }];
    for (const p of paragraphs) {
      ops.push({
        type: "setParaStyle",
        blockId: p.id,
        patch: { list: { listId: def.id, level: p.style.list?.level ?? 0 }, indentFirstLinePx: 0 },
      });
    }
    return tr(ops, sel, "command");
  };
}

/** Create or replace a list definition (the List style constructor). No content
 *  re-patch is needed: markers are read live from Document.lists at layout time.
 *  Undoable via setListDefinition's inverse. */
export function upsertListDefinition(def: ListDefinition): Command {
  return (state) => tr([{ type: "setListDefinition", listId: def.id, def }], state.selection, "command");
}

/** Delete a list definition. Paragraphs still referencing it fall back to the
 *  layout engine's default marker rendering. */
export function deleteListDefinition(listId: string): Command {
  return (state) => tr([{ type: "setListDefinition", listId, def: null }], state.selection, "command");
}

/** Apply (or remove) a legal-style multilevel numbered list (1 / 1.1 / 1.1.1). */
export function toggleMultilevelList(): Command {
  return toggleListById(DEFAULT_MULTILEVEL_LIST_ID, multilevelListDefinition);
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
        columns: cloneColumns(columns),
        pageNumberStart: p.pageNumberStart ?? null, // restart is per-section, never inherited
        headerDistancePx: p.headerDistancePx ?? doc.section.headerDistancePx ?? null,
        footerDistancePx: p.footerDistancePx ?? doc.section.footerDistancePx ?? null,
        pageColorHex: p.pageColorHex ?? doc.section.pageColorHex ?? null,
        pageBorders: p.pageBorders ?? doc.section.pageBorders ?? null,
        // breakType lives directly on the break paragraph (never inherited); the
        // default "nextPage" seeds the dialog's "Next page" option.
        breakType: b.style.sectionBreak.type,
        lineNumbering: p.lineNumbering ? { ...p.lineNumbering } : null,
      };
    }
  }
  return {
    pageWidthPx: doc.section.pageWidthPx,
    pageHeightPx: doc.section.pageHeightPx,
    marginPx: { ...doc.section.marginPx },
    columns: cloneColumns(doc.section.columns ?? null),
    pageNumberStart: doc.section.pageNumberStart ?? null,
    headerDistancePx: doc.section.headerDistancePx ?? null,
    footerDistancePx: doc.section.footerDistancePx ?? null,
    pageColorHex: doc.section.pageColorHex ?? null,
    pageBorders: doc.section.pageBorders ?? null,
    breakType: doc.section.breakType ?? null,
    lineNumbering: doc.section.lineNumbering ? { ...doc.section.lineNumbering } : null,
  };
}

/** Deep-clone the columns object so per-column `cols` entries are not shared. */
function cloneColumns(columns: SectionGeometry["columns"]): SectionGeometry["columns"] {
  if (!columns) return null;
  return { ...columns, ...(columns.cols ? { cols: columns.cols.map((c) => ({ ...c })) } : {}) };
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
    const props: import("@cw/shared").SectionPatch = {
      pageWidthPx: geo.pageWidthPx,
      pageHeightPx: geo.pageHeightPx,
      marginPx: geo.marginPx,
      columns: geo.columns,
    };
    if (geo.pageNumberStart !== null) props.pageNumberStart = geo.pageNumberStart;
    if (geo.headerDistancePx !== null) props.headerDistancePx = geo.headerDistancePx;
    if (geo.footerDistancePx !== null) props.footerDistancePx = geo.footerDistancePx;
    if (geo.pageColorHex !== null) props.pageColorHex = geo.pageColorHex;
    if (geo.pageBorders !== null) props.pageBorders = geo.pageBorders;
    if (geo.lineNumbering !== null) props.lineNumbering = { ...geo.lineNumbering };
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
        loc?.kind === "band" || loc?.kind === "footnote" || loc?.kind === "endnote" || (loc?.kind === "cell" && loc.where !== "body");
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
        if (geometry.headerDistancePx !== null) props.headerDistancePx = geometry.headerDistancePx;
        else delete props.headerDistancePx;
        if (geometry.footerDistancePx !== null) props.footerDistancePx = geometry.footerDistancePx;
        else delete props.footerDistancePx;
        if (geometry.pageColorHex !== null) props.pageColorHex = geometry.pageColorHex;
        else delete props.pageColorHex;
        if (geometry.pageBorders !== null) props.pageBorders = geometry.pageBorders;
        else delete props.pageBorders;
        if (geometry.lineNumbering !== null) props.lineNumbering = geometry.lineNumbering;
        else delete props.lineNumbering;
        return tr(
          [
            {
              type: "setParaStyle",
              blockId: b.id,
              patch: { sectionBreak: { type: geometry.breakType ?? "nextPage", props } },
            },
          ],
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
      loc?.kind === "band" || loc?.kind === "footnote" || loc?.kind === "endnote" || (loc?.kind === "cell" && loc.where !== "body");
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

/** Insert a TOC at the caret, or REGENERATE the existing one (Word's F9). When
 *  regenerating WITHOUT an explicit look, the existing TOC's title presence and
 *  per-level styling are preserved (same fix as the "Update Field (TOC)" context
 *  action) — so the ribbon's insert/update button no longer clobbers an imported
 *  TOC with the editor defaults. An explicit `opts` (the headless backend route)
 *  still wins. */
export function insertTocCmd(opts?: TocOptions): Command {
  return (state) => {
    const existing = tocEntryBlocks(state.doc);
    const effectiveOpts = opts ?? (existing.length > 0 ? tocOptionsFromExisting(state.doc) : undefined);
    const fresh = buildTocParagraphs(state.doc, effectiveOpts);
    if (fresh.length <= 1) return null; // no headings -> nothing to list
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

/** Derive TocOptions that REPRODUCE the document's current TOC look — its title
 *  (or none) and each level's char/para styling sampled from existing entries —
 *  so regenerating an imported TOC doesn't clobber it with the editor defaults. */
function tocOptionsFromExisting(doc: EditorState["doc"]): TocOptions {
  const blocks = tocEntryBlocks(doc);
  const opts: TocOptions = {};

  // Title: reuse the existing generated title, or suppress it entirely when the
  // TOC has none (the common imported case — Word's title is a separate heading
  // we don't touch). Without this, the default "Table of Contents" gets injected.
  const titlePara = blocks.find((b) => b.style.namedStyle === "tocTitle" && !b.style.tocEntry);
  if (titlePara) {
    opts.title = {
      text: textOfRuns(titlePara.runs),
      ...(titlePara.runs[0] ? { char: titlePara.runs[0].style } : {}),
      para: titlePara.style,
      ...(titlePara.style.namedStyle ? { namedStyle: titlePara.style.namedStyle } : {}),
    };
  } else {
    opts.title = null;
  }

  // Per-level styling: the first existing entry of each level supplies the look.
  // Strip the run's stale in-document `link` (each entry links to its OWN heading
  // via tocEntry.targetId; a copied "#anchor" would point every entry at one).
  const levels: Record<number, { char?: Partial<CharStyle>; para?: Partial<ParaStyle> }> = {};
  for (const b of blocks) {
    const te = b.style.tocEntry;
    if (!te || levels[te.level]) continue;
    const sampleChar = b.runs[0]?.style;
    const char = sampleChar ? { ...sampleChar } : undefined;
    if (char) delete char.link;
    levels[te.level] = { ...(char ? { char } : {}), para: b.style };
  }
  if (Object.keys(levels).length > 0) opts.levels = levels;
  return opts;
}

/** Refresh an existing TOC in place (Word's F9): re-list headings + recompute
 *  page numbers while PRESERVING the current TOC's title presence and styling.
 *  Distinct from insertTocCmd, which (re)builds with the caller's/default look. */
export function updateTocFieldCmd(): Command {
  return (state) => insertTocCmd(tocOptionsFromExisting(state.doc))(state);
}

/** Apply edited TOC field switches (`\o` level range, `\h`, `\z`, `\u`, `\p`, …):
 *  persist them on `doc.tocInstruction` AND regenerate the entries so the change
 *  takes visible effect (a wider `\o` range lists more levels), preserving the
 *  current TOC's look. The instruction is the export source of truth, so this also
 *  round-trips to `.docx`. */
export function setTocSwitchesCmd(sw: TocSwitches): Command {
  return (state) => {
    const instruction = buildTocInstruction(sw);
    const setInstr: Op = { type: "setTocInstruction", instruction };
    // buildTocParagraphs reads tocInstruction for the level range, so regenerate
    // against a copy that already has the new instruction; the resulting block ops
    // are independent of it and ride in the same transaction after setInstr.
    const withInstr: EditorState = { ...state, doc: { ...state.doc, tocInstruction: instruction } };
    const regen = insertTocCmd(tocOptionsFromExisting(withInstr.doc))(withInstr);
    if (!regen) return tr([setInstr], state.selection, "command");
    return tr([setInstr, ...regen.ops], regen.selectionAfter, "command");
  };
}

// ---------------------------------------------------------------------------
// Generic custom fields (Document.fields). A contiguous run of body blocks
// sharing a `fieldId` is the field's result region; "Update Field" replaces it
// with freshly-resolved content. TOC keeps its own tocEntry path (insertTocCmd).

/** Classify the field (if any) the block at `blockId` belongs to — drives the
 *  "Update Field (…)" context-menu entry. */
export function fieldAtBlock(
  doc: EditorState["doc"],
  blockId: string,
): { kind: "toc" } | { kind: "custom"; def: import("@cw/shared").FieldDef } | null {
  const b = blockById(doc, blockId);
  if (!b) return null;
  if (b.kind === "paragraph" && (b.style.tocEntry || b.style.namedStyle === "tocTitle")) return { kind: "toc" };
  if (b.fieldId && doc.fields?.[b.fieldId]) return { kind: "custom", def: doc.fields[b.fieldId]! };
  return null;
}

/** Replace a custom field's result region with `blocks` (re-stamped with the
 *  fieldId, fresh ids so they never collide with the originals). The field's
 *  Document.fields entry is untouched, so the instruction round-trips. */
export function replaceFieldResultCmd(fieldId: string, blocks: Block[]): Command {
  return (state) => {
    const all = state.doc.blocks;
    const lo = all.findIndex((b) => b.fieldId === fieldId);
    if (lo < 0) return null;
    let hi = lo;
    while (hi + 1 < all.length && all[hi + 1]!.fieldId === fieldId) hi++;
    const fresh: Block[] = blocks.map((b) => ({ ...b, id: freshBlockId(), fieldId }));
    if (fresh.length === 0) return null;
    const ops: Op[] = [];
    for (let i = hi; i >= lo; i--) ops.push({ type: "removeBlock", blockId: all[i]!.id });
    fresh.forEach((b, k) => ops.push({ type: "insertBlock", index: lo + k, block: b }));
    return tr(ops, caret(fresh[0]!.id, 0), "command");
  };
}

// ---------------------------------------------------------------------------
// Content controls (OOXML w:sdt) — runs sharing a CharStyle.sdtPath form one
// inline control (nested controls share a path PREFIX); block-level controls live
// on Block.sdtPath. Properties live in Document.sdts. Placeholder text is gray and
// replaced WHOLE on first input (Word).

export interface SdtRange {
  blockId: string;
  start: number;
  end: number;
}

/** Re-style runs in [start,end) via `fn` (others pass through), splitting and
 *  renormalizing. Used to add/remove a single sdt id on a sub-range when a
 *  uniform patch won't do (each run's path is transformed independently). */
function mapRunStylesInRange(runs: Run[], start: number, end: number, fn: (s: CharStyle) => CharStyle): Run[] {
  const [head, fromStart] = splitRunsAt(runs, start);
  const [middle, tail] = splitRunsAt(fromStart, end - start);
  const styled = middle.map((r) => ({ text: r.text, style: fn(r.style) }));
  const fallback = styled[0]?.style ?? head[0]?.style ?? tail[0]?.style ?? runs[0]?.style ?? DEFAULT_CELL_CHAR;
  return normalizeRuns([...head, ...styled, ...tail], fallback);
}

/** Every contiguous span belonging to control `id`, document order. Membership is
 *  path-aware: a run whose inline `sdtPath` contains `id` (at any nesting level),
 *  or a run in a block whose block-level `sdtPath` contains `id` (whole-paragraph
 *  block control). */
export function findSdtRanges(doc: EditorState["doc"], id: string): SdtRange[] {
  const out: SdtRange[] = [];
  for (const p of paragraphsOf(doc)) {
    const blockMember = p.sdtPath?.includes(id) ?? false;
    let off = 0;
    let open: SdtRange | null = null;
    for (const r of p.runs) {
      if (blockMember || inSdt(r.style, id)) {
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

/** The control containing a position (the run at the caret, either side). An
 *  untagged spot inside a table a block-level control wraps — an empty cell, or a
 *  blank line in a cell, neither of which carries the sdtId on a run — still
 *  belongs to that control, so fall back to the table's wrapping control. */
export function sdtAtPosition(doc: EditorState["doc"], pos: DocPosition): string | null {
  const block = blockById(doc, pos.blockId);
  if (!block) return null;
  let off = 0;
  let before: string | null = null;
  let after: string | null = null;
  for (const r of block.runs) {
    const end = off + r.text.length;
    if (pos.offset > off && pos.offset <= end) before = innermostSdtId(r.style) ?? null;
    if (pos.offset >= off && pos.offset < end) after = innermostSdtId(r.style) ?? null;
    off = end;
  }
  const direct = after ?? before;
  if (direct) return direct;
  // No inline control on the runs — a block-level control wrapping this paragraph?
  if (block.sdtPath?.length) return block.sdtPath[block.sdtPath.length - 1]!;
  const loc = locateParagraph(doc, pos.blockId);
  if (loc?.kind !== "cell") return null;
  const table = containerBlocks(doc, loc.where)[loc.bi];
  if (table?.kind !== "table") return null;
  // A block-level control wrapping the whole table, else the dominant cell control.
  if (table.sdtPath?.length) return table.sdtPath[table.sdtPath.length - 1]!;
  return dominantCellSdt(table);
}

/** The hyperlink URL on the run at `pos` (favouring the run the caret is entering,
 *  else the one it's leaving), or null. Position-based analog of the point-based
 *  `linkAt` (geometry.ts): links live on run styles (`run.style.link`), so this
 *  scans the caret paragraph's runs — used to drive the hyperlink context toolbar. */
export function linkAtPosition(doc: EditorState["doc"], pos: DocPosition): string | null {
  const block = blockById(doc, pos.blockId);
  if (!block) return null;
  let off = 0;
  let before: string | null = null;
  let after: string | null = null;
  for (const r of block.runs) {
    const end = off + r.text.length;
    if (pos.offset > off && pos.offset <= end) before = r.style.link ?? null;
    if (pos.offset >= off && pos.offset < end) after = r.style.link ?? null;
    off = end;
  }
  return after ?? before;
}

/** The footnote/endnote whose superscript marker run sits at `pos` (favouring the
 *  run the caret is entering, else the one it's leaving), or null. Note markers are
 *  single superscript runs carrying `footnoteRef`/`endnoteRef`; this scans the caret
 *  paragraph's runs the same way `linkAtPosition` scans for links, and drives the
 *  footnote/endnote context toolbar. */
export function noteAtPosition(
  doc: EditorState["doc"],
  pos: DocPosition,
): { kind: "footnote" | "endnote"; id: string } | null {
  const block = blockById(doc, pos.blockId);
  if (!block || block.kind !== "paragraph") return null;
  const noteOf = (r: Run): { kind: "footnote" | "endnote"; id: string } | null =>
    r.style.footnoteRef
      ? { kind: "footnote", id: r.style.footnoteRef }
      : r.style.endnoteRef
        ? { kind: "endnote", id: r.style.endnoteRef }
        : null;
  let off = 0;
  let before: { kind: "footnote" | "endnote"; id: string } | null = null;
  let after: { kind: "footnote" | "endnote"; id: string } | null = null;
  for (const r of block.runs) {
    const end = off + r.text.length;
    if (pos.offset > off && pos.offset <= end) before = noteOf(r);
    if (pos.offset >= off && pos.offset < end) after = noteOf(r);
    off = end;
  }
  return after ?? before;
}

/** The INLINE content control directly on the run at `pos`, ignoring any
 *  block-level control wrapping the whole paragraph/table. Splitting a paragraph
 *  is only unsafe INSIDE an inline control (it can't span a paragraph break); a
 *  block-level control can legitimately hold many paragraphs, so Enter must work
 *  there. Used by the Enter handler to decide whether to allow the split. */
export function inlineSdtAtPosition(doc: EditorState["doc"], pos: DocPosition): string | null {
  const block = blockById(doc, pos.blockId);
  if (!block) return null;
  let off = 0;
  let before: string | null = null;
  let after: string | null = null;
  for (const r of block.runs) {
    const end = off + r.text.length;
    if (pos.offset > off && pos.offset <= end) before = innermostSdtId(r.style) ?? null;
    if (pos.offset >= off && pos.offset < end) after = innermostSdtId(r.style) ?? null;
    off = end;
  }
  return after ?? before;
}

/** The full outer→inner content-control chain at a position: a containing table's
 *  block-level controls (when the position is in a cell), then the paragraph's
 *  block-level controls, then the run's inline controls. Empty when none. Drives
 *  the nested-control adornment (frames + breadcrumb). */
export function sdtStackAtPosition(doc: EditorState["doc"], pos: DocPosition): string[] {
  const block = blockById(doc, pos.blockId);
  if (!block) return [];
  let before: string[] | undefined;
  let after: string[] | undefined;
  let off = 0;
  for (const r of block.runs) {
    const end = off + r.text.length;
    if (pos.offset > off && pos.offset <= end) before = r.style.sdtPath;
    if (pos.offset >= off && pos.offset < end) after = r.style.sdtPath;
    off = end;
  }
  const inline = after ?? before ?? [];
  const blockPath = block.sdtPath ?? [];
  let tablePath: string[] = [];
  const loc = locateParagraph(doc, pos.blockId);
  if (loc?.kind === "cell") {
    const table = containerBlocks(doc, loc.where)[loc.bi];
    if (table?.kind === "table") tablePath = table.sdtPath ?? [];
  }
  return [...tablePath, ...blockPath, ...inline];
}

// Inline fields (OOXML complex field) — contiguous runs sharing a CharStyle.fieldId
// form one inline field's RESULT; the definition lives in Document.fields. Mirrors
// the SDT helpers above (block-region fields use Block.fieldId + fieldAtBlock).

/** Every contiguous run span carrying this inline fieldId, document order. */
export function findFieldRanges(doc: EditorState["doc"], id: string): SdtRange[] {
  const out: SdtRange[] = [];
  for (const p of paragraphsOf(doc)) {
    let off = 0;
    let open: SdtRange | null = null;
    for (const r of p.runs) {
      if (r.style.fieldId === id) {
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

/** The inline field id at a position (the run on either side of the caret), or null. */
export function fieldAtPosition(doc: EditorState["doc"], pos: DocPosition): string | null {
  const block = blockById(doc, pos.blockId);
  if (!block) return null;
  let off = 0;
  let before: string | null = null;
  let after: string | null = null;
  for (const r of block.runs) {
    const end = off + r.text.length;
    if (pos.offset > off && pos.offset <= end) before = r.style.fieldId ?? null;
    if (pos.offset >= off && pos.offset < end) after = r.style.fieldId ?? null;
    off = end;
  }
  return after ?? before;
}

const runsTextLen = (runs: Run[]): number => runs.reduce((n, r) => n + r.text.length, 0);

/** Evaluate a field spec to its result runs, all tagged with the field id (a
 *  `{token}` run for PAGE/NUMPAGES, materialized text for DATE/TIME/IF). */
function fieldResultRuns(spec: FieldSpec, baseStyle: CharStyle, id: string): Run[] {
  const res = evaluateField(spec, baseStyle, { now: new Date() });
  const runs = res.kind === "token" ? [{ text: res.token, style: baseStyle }] : res.runs;
  return runs.map((r) => ({ text: r.text, style: { ...r.style, fieldId: id } }));
}

const fieldBaseStyle = (doc: EditorState["doc"], blockId: string, offset: number): CharStyle | undefined => {
  const block = blockById(doc, blockId);
  return block ? styleAtRuns(block.runs, offset) ?? block.runs[0]?.style : undefined;
};

/** Insert a built-in field at the caret: register its FieldDef and insert the
 *  fieldId-tagged result runs. */
export function insertFieldCmd(spec: FieldSpec): Command {
  return (state) => {
    const base = withSelectionDeleted(state);
    if (!base) return null;
    const baseStyle = fieldBaseStyle(state.doc, base.at.blockId, base.at.offset);
    if (!baseStyle) return null;
    const id = freshBlockId();
    const def: FieldDef = { id, instruction: buildInstruction(spec), name: spec.type, kind: "builtin", spec };
    const runs = fieldResultRuns(spec, baseStyle, id);
    if (runs.length === 0) return null;
    base.ops.push({ type: "setField", id, def }, { type: "insertRuns", at: base.at, runs });
    return tr(base.ops, caret(base.at.blockId, base.at.offset + runsTextLen(runs)), "command");
  };
}

/** Re-define an existing inline field and replace its result in place. */
export function editFieldCmd(fieldId: string, spec: FieldSpec): Command {
  return (state) => {
    const ranges = findFieldRanges(state.doc, fieldId);
    if (ranges.length === 0) return null;
    const blockId = ranges[0]!.blockId;
    const same = ranges.filter((r) => r.blockId === blockId);
    const start = Math.min(...same.map((r) => r.start));
    const end = Math.max(...same.map((r) => r.end));
    const baseStyle = fieldBaseStyle(state.doc, blockId, start);
    if (!baseStyle) return null;
    const def: FieldDef = { id: fieldId, instruction: buildInstruction(spec), name: spec.type, kind: "builtin", spec };
    const runs = fieldResultRuns(spec, baseStyle, fieldId);
    const ops: Op[] = [
      { type: "setField", id: fieldId, def },
      { type: "deleteRange", blockId, start, end },
      { type: "insertRuns", at: { blockId, offset: start }, runs },
    ];
    return tr(ops, caret(blockId, start + runsTextLen(runs)), "command");
  };
}

/** Recompute an existing field's result (Word's F9): DATE/TIME refresh to now, IF
 *  re-evaluates. PAGE/NUMPAGES are layout-resolved per page, so they're a no-op. */
export function updateFieldCmd(fieldId: string): Command {
  return (state) => {
    const def = state.doc.fields?.[fieldId];
    if (!def?.spec || def.spec.type === "PAGE" || def.spec.type === "NUMPAGES") return null;
    return editFieldCmd(fieldId, def.spec)(state);
  };
}

/** The control wrapping the most cells of a table — the block-level container,
 *  distinguished from a stray inline field that tags a single cell. */
function dominantCellSdt(table: TableBlock): string | null {
  const cellsWith = new Map<string, number>();
  for (const row of table.rows) {
    for (const cell of row.cells) {
      const ids = new Set<string>();
      for (const b of cell.blocks) {
        for (const bid of b.sdtPath ?? []) ids.add(bid);
        if (b.kind === "paragraph") for (const r of b.runs) for (const sid of r.style.sdtPath ?? []) ids.add(sid);
      }
      for (const id of ids) cellsWith.set(id, (cellsWith.get(id) ?? 0) + 1);
    }
  }
  let best: string | null = null;
  let bestN = 0;
  for (const [id, n] of cellsWith) {
    if (n > bestN) {
      best = id;
      bestN = n;
    }
  }
  return best;
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
        // Append to any existing inline ancestry so wrapping inside another control nests.
        runs: mapRunStylesInRange(block.runs, from.offset, to.offset, (s) => ({ ...s, sdtPath: pushSdt(s.sdtPath, id) })),
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
        sdtPath: pushSdt(inherited.sdtPath, id),
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

/** Wrap a SELECTED IMAGE OBJECT in a block-level content control. The text-based
 *  insertContentControl can't act on an object selection (its `state.selection` is
 *  null while an image is selected), so this is the object-side entry: tag the image
 *  block with a fresh sdtId — appended to any existing block-level ancestry, so it
 *  nests — and register the control's props. The image is the control's real content
 *  (no placeholder). Works for a top-level image and one inside a table cell. */
export function wrapImageInContentControl(blockId: string, type: SdtType = "richText", props: Partial<SdtProps> = {}): Command {
  return (state) => {
    const loc = locateImage(state.doc, blockId);
    if (!loc) return null;
    const id = freshSdtId();
    const full: SdtProps = { type, ...props };
    const updated: ImageBlock = {
      ...loc.image,
      sdtPath: pushSdt(loc.image.sdtPath, id),
      revision: loc.image.revision + 1,
    };
    const ops: Op[] = [{ type: "setSdtProps", id, props: full }];
    if (loc.kind === "top") {
      // No "set block sdtPath" op exists; remove+insert the same id is the
      // established block-rewrite pattern (cf. replaceSdtBlockSpan).
      ops.push({ type: "removeBlock", blockId });
      ops.push({ type: "insertBlock", index: loc.index, block: updated, where: loc.where });
    } else {
      const table = state.doc.blocks[loc.bi] as TableBlock;
      const row = table.rows[loc.ri]!;
      const cells = row.cells.map((c, ci) =>
        ci === loc.ci ? { ...c, blocks: c.blocks.map((b, ii) => (ii === loc.ii ? updated : b)) } : c,
      );
      ops.push({ type: "setTableRow", tableId: table.id, rowIndex: loc.ri, row: { ...row, cells } });
    }
    // Keep the object selection (state.selection stays null) so the new control
    // chrome appears around the still-selected image.
    return tr(ops, state.selection, "command");
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
    // Preserve the control's full ancestry up to and including `id`.
    const path = ancestryThrough(base.sdtPath, id) ?? [id];
    const ops: Op[] = [
      { type: "deleteRange", blockId: range.blockId, start: range.start, end: range.end },
      {
        type: "insertText",
        at: { blockId: range.blockId, offset: range.start },
        text,
        style: { ...base, color: props.placeholder ? "#202124" : base.color, sdtPath: path },
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
    // The ancestry the replaced content must carry (through `id`), read from the
    // control's existing first run.
    const origBlock = blockById(state.doc, first.blockId);
    const baseStyle = origBlock ? styleAtRuns(origBlock.runs, first.start + 1) : undefined;
    const path = ancestryThrough(baseStyle?.sdtPath, id) ?? [id];
    const tagged: DocFragment = {
      inline: fragment.inline,
      blocks: fragment.blocks.map((b) => ({
        style: b.style,
        runs: b.runs.map((r) => ({ text: r.text, style: { ...r.style, sdtPath: path } })),
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
    // Block-level control membership lives on Block.sdtPath; preserve ancestry through `id`.
    const blockPath = ancestryThrough((span[lo] as Block).sdtPath, id) ?? [id];
    const taggedPara = (p: Paragraph): Paragraph => ({
      kind: "paragraph",
      id: freshBlockId(),
      revision: 0,
      style: p.style,
      // Runs keep their own inline sdt ancestry; the block carries the block-level path.
      runs: p.runs.length > 0 ? p.runs.map((r) => ({ text: r.text, style: r.style })) : [{ text: "", style: baseStyle }],
      sdtPath: blockPath,
    });
    const emptyTagged = (): Paragraph =>
      taggedPara({ kind: "paragraph", id: "", revision: 0, runs: [], style: (span[lo] as Paragraph).style });

    const fresh: Block[] = blocks.map((b) =>
      b.kind === "paragraph" ? taggedPara(b) : { ...b, sdtPath: blockPath },
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
        // Preserve the inline control's full ancestry up to and including `id`.
        const path = ancestryThrough(base.sdtPath, id) ?? [id];
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
            sdtPath: path,
          },
        }));
        if (runs.length === 0 || runs.every((rn) => rn.text.length === 0)) {
          runs = [{ text: "", style: { ...base, sdtPath: path } }];
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
        // Remove just THIS id from each run's path, preserving any outer/inner controls.
        // (Removing a block-level control's Block.sdtPath interactively is a v1
        // follow-up — block controls are import/inspector-driven.)
        const runs = mapRunStylesInRange(block.runs, r.start, r.end, (s) => ({ ...s, sdtPath: removeSdt(s.sdtPath, id) }));
        ops.push({ type: "setRuns", blockId: r.blockId, runs });
      }
    }
    // Block-level IMAGE controls (created by wrapImageInContentControl) carry the id
    // on Block.sdtPath, which findSdtRanges (paragraph-runs only) can't see — strip
    // it from the image block itself, else it would point at a now-deleted control.
    if (!deleteContents) {
      const stripImage = (img: ImageBlock): ImageBlock => {
        const path = removeSdt(img.sdtPath, id); // undefined once the path is empty
        const next: ImageBlock = { ...img, revision: img.revision + 1 };
        if (path) next.sdtPath = path;
        else delete next.sdtPath;
        return next;
      };
      state.doc.blocks.forEach((b, bi) => {
        if (b.kind === "image" && b.sdtPath?.includes(id)) {
          ops.push({ type: "removeBlock", blockId: b.id });
          ops.push({ type: "insertBlock", index: bi, block: stripImage(b), where: "body" });
        } else if (b.kind === "table") {
          b.rows.forEach((row, ri) => {
            if (!row.cells.some((c) => c.blocks.some((cb) => cb.kind === "image" && cb.sdtPath?.includes(id)))) return;
            const cells = row.cells.map((c) => ({
              ...c,
              blocks: c.blocks.map((cb) => (cb.kind === "image" && cb.sdtPath?.includes(id) ? stripImage(cb) : cb)),
            }));
            ops.push({ type: "setTableRow", tableId: b.id, rowIndex: ri, row: { ...row, cells } });
          });
        }
      });
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

/** A note reference marker's location: its host paragraph, run index, start
 *  offset within the paragraph, and the note id it points at. */
interface NoteRefAt {
  para: Paragraph;
  runIdx: number;
  offset: number;
  id: string;
}

/** Every footnote/endnote reference of one kind in document order, walking body
 *  AND table-cell paragraphs so numbering stays correct wherever a ref lives.
 *  Shared by the insert commands (renumber-on-insert) and `deleteNoteCmd`
 *  (renumber-on-delete). Callers that need document ORDER of arbitrary blocks
 *  keep `body` around separately for that lookup. */
function noteRefsIn(body: Paragraph[], refKey: "footnoteRef" | "endnoteRef"): NoteRefAt[] {
  const refs: NoteRefAt[] = [];
  body.forEach((para) => {
    let off = 0;
    para.runs.forEach((r, runIdx) => {
      const id = r.style[refKey];
      if (id) refs.push({ para, runIdx, offset: off, id });
      off += r.text.length;
    });
  });
  return refs;
}

export function insertFootnoteCmd(): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel || !isCollapsed(sel)) return null;
    const at = sel.focus;
    const loc = locateParagraph(state.doc, at.blockId);
    // Body paragraphs OR body table cells — not header/footer bands or notes.
    if (loc?.kind !== "top" && !(loc?.kind === "cell" && loc.where === "body")) return null;

    const body = bodyParagraphs(state.doc);
    const refs = noteRefsIn(body, "footnoteRef");
    const orderOf = (blockId: string): number => body.findIndex((p) => p.id === blockId);
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
      const p = blockById(state.doc, paraId)!; // cell-aware (refs may live in cells)
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

// ---------------------------------------------------------------------------
// Endnotes — identical machinery to footnotes (the ref run's TEXT is its
// number; inserting renumbers every later ref in the same transaction), but the
// note body lives in Document.endnotes and lays out at the document end.

export function insertEndnoteCmd(): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel || !isCollapsed(sel)) return null;
    const at = sel.focus;
    const loc = locateParagraph(state.doc, at.blockId);
    // Body paragraphs OR body table cells — not header/footer bands or notes.
    if (loc?.kind !== "top" && !(loc?.kind === "cell" && loc.where === "body")) return null;

    const body = bodyParagraphs(state.doc);
    const refs = noteRefsIn(body, "endnoteRef");
    const orderOf = (blockId: string): number => body.findIndex((p) => p.id === blockId);
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
      const p = blockById(state.doc, paraId)!; // cell-aware (refs may live in cells)
      const runs = p.runs.map((r, i) => (m.has(i) ? { ...r, text: m.get(i)! } : r));
      ops.push({ type: "setRuns", blockId: paraId, runs });
    }

    const noteId = `en_${freshBlockId()}`;
    const block = blockById(state.doc, at.blockId)!;
    const inherited = styleAtRuns(block.runs, at.offset) ?? DEFAULT_CELL_CHAR;
    // styleAtRuns can echo a preceding note marker's style, so drop BOTH ref
    // flags before composing the new endnote runs — neither the superscript
    // marker nor the seeded body may carry a stray footnoteRef/endnoteRef.
    const baseNoteStyle = { ...inherited, footnoteRef: undefined, endnoteRef: undefined, verticalAlign: undefined, link: undefined };
    ops.push({
      type: "insertText",
      at,
      text: String(newNumber),
      style: { ...baseNoteStyle, verticalAlign: "super", endnoteRef: noteId },
    });
    const notePara: Paragraph = {
      kind: "paragraph",
      id: freshBlockId(),
      revision: 0,
      runs: [{ text: "", style: { ...baseNoteStyle, fontSizePx: 12 } }],
      style: { align: "left", lineHeight: 1.4, spaceBeforePx: 0, spaceAfterPx: 2, indentFirstLinePx: 0, indentLeftPx: 0 },
    };
    ops.push({ type: "setEndnote", noteId, paras: [notePara] });
    // Word drops the caret into the new note for immediate typing.
    return tr(ops, caret(notePara.id, 0), "command");
  };
}

/** Delete a footnote/endnote by id: drop its superscript marker run from the host
 *  paragraph, delete the note body (`setFootnote`/`setEndnote` with no paras), and
 *  renumber every remaining ref of the same kind in one transaction — mirroring the
 *  insert commands' renumber pass in reverse. Drives the note context toolbar's
 *  Delete action. No-op if the note doesn't exist. */
export function deleteNoteCmd(kind: "footnote" | "endnote", noteId: string): Command {
  return (state) => {
    const doc = state.doc;
    const registry = kind === "footnote" ? doc.footnotes : doc.endnotes;
    if (!registry?.[noteId]) return null;
    const refKey = kind === "footnote" ? "footnoteRef" : "endnoteRef";
    const noteOp = (id: string, paras: Paragraph[] | null): Op =>
      kind === "footnote" ? { type: "setFootnote", noteId: id, paras } : { type: "setEndnote", noteId: id, paras };

    // Every ref of this kind, document order (body + cell paragraphs).
    const refs = noteRefsIn(bodyParagraphs(doc), refKey);
    const target = refs.find((r) => r.id === noteId);
    // Orphan body (marker already gone) — just drop the dangling body.
    if (!target) return tr([noteOp(noteId, null)], state.selection ?? null, "command");

    // Per-paragraph run edits: drop the target marker run; renumber the rest by
    // their post-deletion document order (1-based).
    const edits = new Map<string, Map<number, string | null>>(); // null → drop the run
    const editFor = (paraId: string): Map<number, string | null> => {
      let m = edits.get(paraId);
      if (!m) {
        m = new Map();
        edits.set(paraId, m);
      }
      return m;
    };
    editFor(target.para.id).set(target.runIdx, null);
    refs
      .filter((r) => r.id !== noteId)
      .forEach((r, i) => editFor(r.para.id).set(r.runIdx, String(i + 1)));

    const ops: Op[] = [];
    for (const [paraId, m] of edits) {
      const p = blockById(doc, paraId)!; // cell-aware (refs may live in cells)
      const runs: Run[] = [];
      p.runs.forEach((r, i) => {
        if (!m.has(i)) {
          runs.push(r);
          return;
        }
        const v = m.get(i);
        if (v !== null) runs.push({ ...r, text: v! });
      });
      // A paragraph must keep at least one run; if dropping the marker emptied it,
      // seed an empty run carrying a cleaned style (no stray ref/superscript flags).
      if (runs.length === 0) {
        const base = p.runs[target.runIdx] ?? p.runs[0]!;
        runs.push({ text: "", style: { ...base.style, footnoteRef: undefined, endnoteRef: undefined, verticalAlign: undefined } });
      }
      ops.push({ type: "setRuns", blockId: paraId, runs });
    }
    ops.push(noteOp(noteId, null));
    // Collapse the caret to where the marker was.
    return tr(ops, caret(target.para.id, target.offset), "command");
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
  key: "bold" | "italic" | "underline" | "strikethrough" | "caps" | "smallCaps" | "doubleStrikethrough",
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

export function insertImage(src: string, widthPx: number, heightPx: number, mediaId?: string): Command {
  return (state) =>
    insertBlockAtCaret(state, () => {
      const img: ImageBlock = {
        kind: "image",
        id: freshBlockId(),
        revision: 0,
        src,
        widthPx,
        heightPx,
        align: "center",
      };
      // Content address (sha256) so the bytes survive serialize/export and resolve
      // from the shared media store after a commit-back from a child document.
      if (mediaId) img.mediaId = mediaId;
      return img;
    });
}

/** Default size for a freshly-inserted shape (px). A modest box the user resizes. */
export const SHAPE_DEFAULT_WIDTH_PX = 180;
export const SHAPE_DEFAULT_HEIGHT_PX = 120;

/** Insert a drawing shape (preset geometry) as its own block at the caret. Seeds a
 *  solid fill + solid outline so the shape is visible and immediately editable. A
 *  "line" preset gets no fill — OOXML prst="line" is the box diagonal, a stroke-only
 *  geometry — but keeps a full box so it stays selectable/resizable. */
export function insertShape(preset: ShapePreset): Command {
  return (state) =>
    insertBlockAtCaret(state, (): ShapeBlock => {
      const shape: ShapeBlock = {
        kind: "shape",
        id: freshBlockId(),
        revision: 0,
        geometry: { preset },
        widthPx: SHAPE_DEFAULT_WIDTH_PX,
        heightPx: SHAPE_DEFAULT_HEIGHT_PX,
        align: "left",
        stroke: { color: "#1f77b4", widthPt: 1 },
      };
      if (preset !== "line") shape.fill = { color: "#bcd6ef" };
      return shape;
    });
}

/** Insert a display equation as its own block at the caret. */
export function insertEquation(equation: MathEquation): Command {
  return (state) =>
    insertBlockAtCaret(state, (): EquationBlock => ({
      kind: "equation",
      id: freshBlockId(),
      revision: 0,
      equation,
      align: "center",
    }));
}

/** Replace an existing display equation's MathML (the editor's Apply). */
export function editEquationCmd(blockId: string, equation: MathEquation): Command {
  return (state) => tr([{ type: "setEquation", blockId, equation }], state.selection, "command");
}

/** Set a display equation's horizontal alignment (left / center / right).
 *  Container-aware: resolves equations in the body, header/footer bands, or
 *  table cells (cf. the cell-aware reducer op). */
export function setEquationAlignCmd(blockId: string, align: "left" | "center" | "right"): Command {
  return (state) => {
    if (!locateEquation(state.doc, blockId)) return null;
    return tr([{ type: "setEquationAlign", blockId, align }], state.selection, "command");
  };
}

/** Min/max uniform scale for a display equation (drag-to-resize). Keeps an
 *  equation readable at the small end and within a page at the large end. */
export const EQUATION_SCALE_MIN = 0.25;
export const EQUATION_SCALE_MAX = 4;

/** Set a display equation's uniform size multiplier (drag-to-resize, like an
 *  image). Clamps to [EQUATION_SCALE_MIN, EQUATION_SCALE_MAX]; a scale of ~1
 *  clears the field back to the default. Container-aware via locateEquation. */
export function setEquationScaleCmd(blockId: string, scale: number): Command {
  return (state) => {
    const loc = locateEquation(state.doc, blockId);
    if (!loc) return null;
    if (!Number.isFinite(scale)) return null;
    const clamped = Math.min(EQUATION_SCALE_MAX, Math.max(EQUATION_SCALE_MIN, scale));
    if (clamped === (loc.block.scale ?? 1)) return null; // no-op — don't dirty history
    return tr([{ type: "setEquationScale", blockId, scale: clamped }], state.selection, "command");
  };
}

/** Insert an inline equation (a single U+FFFC run carrying the MathML) at the
 *  caret, in the surrounding text style. Mirrors insertFieldCmd. */
export function insertInlineEquation(equation: MathEquation): Command {
  return (state) => {
    const base = withSelectionDeleted(state);
    if (!base) return null;
    const baseStyle = fieldBaseStyle(state.doc, base.at.blockId, base.at.offset);
    if (!baseStyle) return null;
    const runs: Run[] = [{ text: "￼", style: { ...baseStyle, equation: { ...equation, display: false } } }];
    base.ops.push({ type: "insertRuns", at: base.at, runs });
    return tr(base.ops, caret(base.at.blockId, base.at.offset + 1), "command");
  };
}

/** Replace an existing inline equation (the run at `blockId`/`offset`) with new
 *  MathML, preserving the run's surrounding style. */
export function editInlineEquationCmd(blockId: string, offset: number, equation: MathEquation): Command {
  return (state) => {
    const b = blockById(state.doc, blockId);
    if (!b || b.kind !== "paragraph") return null;
    // `offset` is the equation char's INDEX (the run AT it), not a caret position.
    const style = styleOfCharAt(b.runs, offset);
    if (!style?.equation) return null;
    const ops: Op[] = [
      { type: "deleteRange", blockId, start: offset, end: offset + 1 },
      { type: "insertRuns", at: { blockId, offset }, runs: [{ text: "￼", style: { ...style, equation: { ...equation, display: false } } }] },
    ];
    return tr(ops, caret(blockId, offset + 1), "command");
  };
}

/** Decode a symbol-font hex code point (OOXML w:sym/@w:char, usually a Private-Use
 *  value like "F0E0") to its glyph, so the inserted run paints in the symbol face.
 *  Falls back to the replacement char for an unparseable code so insertion never
 *  throws — mirrors the importer's symbolGlyph. */
function decodeSymbolGlyph(charHex: string): string {
  const cp = parseInt(charHex, 16);
  return Number.isFinite(cp) && cp > 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : "�";
}

/** Insert a symbol-font glyph at the caret (OOXML w:sym): a run whose `text` is the
 *  decoded glyph, `fontFamily` is the symbol font (so it paints correctly without
 *  consulting the marker), and `symbol` carries the font + upper-case hex code point
 *  for a faithful docx round-trip. Mirrors insertInlineEquation / insertFieldCmd. */
export function insertSymbolCmd(font: string, charHex: string): Command {
  return (state) => {
    const base = withSelectionDeleted(state);
    if (!base) return null;
    const baseStyle = fieldBaseStyle(state.doc, base.at.blockId, base.at.offset);
    if (!baseStyle) return null;
    const char = charHex.toUpperCase();
    const glyph = decodeSymbolGlyph(char);
    const runs: Run[] = [{ text: glyph, style: { ...baseStyle, fontFamily: `${font}, sans-serif`, symbol: { font, char } } }];
    base.ops.push({ type: "insertRuns", at: base.at, runs });
    return tr(base.ops, caret(base.at.blockId, base.at.offset + glyph.length), "command");
  };
}

/** Delete an inline equation (the single U+FFFC run at char index `offset`). */
export function removeInlineEquationCmd(blockId: string, offset: number): Command {
  return (state) => {
    const b = blockById(state.doc, blockId);
    if (!b || b.kind !== "paragraph") return null;
    if (!styleOfCharAt(b.runs, offset)?.equation) return null;
    return tr([{ type: "deleteRange", blockId, start: offset, end: offset + 1 }], caret(blockId, offset), "command");
  };
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

/** Carry a proto cell's VISUAL format (not its content, merges, or width) onto a
 *  freshly inserted cell, so a new row/column inherits the look of the row/column
 *  it grew from. Table-level defaults (borders/shading/cell margin) are baked onto
 *  every cell at import/build time, so without this a new cell falls back to the
 *  engine's bare defaults (light grid, no fill, default padding) instead of the
 *  table's. colSpan/rowSpan/preferredWidth are intentionally NOT copied. */
function cloneCellFormat(proto: TableCell | undefined): Partial<TableCell> {
  if (!proto) return {};
  const out: Partial<TableCell> = {};
  if (proto.shading !== undefined) out.shading = proto.shading;
  if (proto.borders !== undefined) out.borders = structuredClone(proto.borders);
  if (proto.margin !== undefined) out.margin = { ...proto.margin };
  if (proto.vAlign !== undefined) out.vAlign = proto.vAlign;
  if (proto.textDirection !== undefined) out.textDirection = proto.textDirection;
  if (proto.noWrap !== undefined) out.noWrap = proto.noWrap;
  if (proto.fitText !== undefined) out.fitText = proto.fitText;
  if (proto.hideMark !== undefined) out.hideMark = proto.hideMark;
  return out;
}

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
      cells: protoRow.cells.map((c) => ({
        id: freshBlockId(),
        blocks: [emptyCellPara(firstCellPara(c))],
        ...cloneCellFormat(c),
      })),
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
      ...cloneCellFormat(row.cells[ctx.ci]),
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

/** Find a body- or band-level table by id (cell selections are body-only today,
 *  but a table can also live in a header/footer story). */
export function findTableById(doc: EditorState["doc"], tableId: string): { table: TableBlock; where: "body" | (typeof BAND_CONTAINERS)[number] } | null {
  for (const where of ["body", ...BAND_CONTAINERS] as const) {
    const t = containerListOf(doc, where).find((b): b is TableBlock => b.kind === "table" && b.id === tableId);
    if (t) return { table: t, where };
  }
  return null;
}

/** Resolve a rectangular cell range to act on: an explicit `override` (the modal
 *  passes the range it captured) or the live cell selection wins; otherwise a
 *  text selection straddling two cells of the same table is the rectangle
 *  between them. Returns grid-space coordinates. */
function cellRangeFromState(state: EditorState, override?: CellSelection | null): { table: TableBlock; rect: GridRect } | null {
  const cs = override ?? state.cellSelection;
  if (cs) {
    const found = findTableById(state.doc, cs.tableId);
    if (!found) return null;
    return { table: found.table, rect: { r0: cs.anchor.row, c0: cs.anchor.col, r1: cs.focus.row, c1: cs.focus.col } };
  }
  const sel = state.selection;
  if (!sel) return null;
  const a = locateParagraph(state.doc, sel.anchor.blockId);
  const f = locateParagraph(state.doc, sel.focus.blockId);
  if (a?.kind !== "cell" || f?.kind !== "cell" || a.where !== f.where || a.bi !== f.bi) return null;
  const table = containerBlocks(state.doc, a.where)[a.bi] as TableBlock;
  const grid = buildTableGrid(table);
  const oa = gridOriginOfCell(grid, a.ri, a.ci);
  const of = gridOriginOfCell(grid, f.ri, f.ci);
  if (!oa || !of) return null;
  return { table, rect: { r0: oa.row, c0: oa.col, r1: of.row, c1: of.col } };
}

/** A block is "blank" if it's an empty paragraph. Used to avoid stacking a pile
 *  of empty paragraphs when merging mostly-empty cells. */
const isBlankPara = (b: Block): boolean => b.kind === "paragraph" && b.runs.every((r) => r.text.length === 0);

/** Concatenate covered cells' content for a merge, dropping wholly-blank cells
 *  unless they're all blank (then keep the first cell's single paragraph). */
function mergedBlocks(cells: TableCell[]): Block[] {
  const nonBlank = cells.filter((c) => !c.blocks.every(isBlankPara));
  if (nonBlank.length === 0) return cells[0]!.blocks.slice(0, 1);
  return nonBlank.flatMap((c) => c.blocks);
}

/** Merge any rectangular block of cells (horizontal, vertical, or 2-D) into one
 *  spanning cell. The rectangle is normalized so it can't bisect an existing
 *  merged cell. Content concatenates into the top-left; styling follows it.
 *  Emitted as one setTableStructure op (multiple rows change), invertible. */
export function mergeCellsCmd(): Command {
  return (state) => {
    const range = cellRangeFromState(state);
    if (!range) return null;
    const grid = buildTableGrid(range.table);
    const rect = normalizeRect(grid, range.rect);
    if (rect.r0 === rect.r1 && rect.c0 === rect.c1) return null; // single cell — nothing to merge
    const topLeft = grid.slots[rect.r0]![rect.c0]!.cell;
    const colSpan = rect.c1 - rect.c0 + 1;
    const rowSpan = rect.r1 - rect.r0 + 1;
    const merged: TableCell = {
      id: topLeft.id,
      blocks: mergedBlocks(cellsInRect(grid, rect).map((s) => s.cell)),
      ...(colSpan > 1 ? { colSpan } : {}),
      ...(rowSpan > 1 ? { rowSpan } : {}),
      ...(topLeft.shading !== undefined ? { shading: topLeft.shading } : {}),
      ...(topLeft.borders !== undefined ? { borders: topLeft.borders } : {}),
      ...(topLeft.margin !== undefined ? { margin: topLeft.margin } : {}),
    };
    const rows = mergeRows(grid, rect, merged);
    const firstPara = firstCellPara(merged);
    return tr(
      [structureOp(range.table, rows)],
      firstPara ? caret(firstPara.id, 0) : null,
      "command",
    );
  };
}

/** The cell to unmerge: the top-left of an active cell selection, else the cell
 *  containing the caret. */
function unmergeTarget(state: EditorState): { table: TableBlock; row: number; col: number; cell: TableCell } | null {
  const cs = state.cellSelection;
  if (cs) {
    const found = findTableById(state.doc, cs.tableId);
    if (!found) return null;
    const grid = buildTableGrid(found.table);
    const rect = normalizeRect(grid, { r0: cs.anchor.row, c0: cs.anchor.col, r1: cs.focus.row, c1: cs.focus.col });
    const slot = grid.slots[rect.r0]?.[rect.c0];
    if (!slot) return null;
    return { table: found.table, row: slot.originRow, col: slot.originCol, cell: slot.cell };
  }
  const ctx = cellContext(state);
  if (!ctx) return null;
  const grid = buildTableGrid(ctx.table);
  const origin = gridOriginOfCell(grid, ctx.ri, ctx.ci);
  if (!origin) return null;
  return { table: ctx.table, row: origin.row, col: origin.col, cell: ctx.table.rows[ctx.ri]!.cells[ctx.ci]! };
}

/** Split a merged cell back into a 1x1 grid of cells (content stays in the
 *  top-left; the freed slots become empty cells). Works for colSpan, rowSpan,
 *  or both. */
export function unmergeCellCmd(): Command {
  return (state) => {
    const target = unmergeTarget(state);
    if (!target) return null;
    const { table, row, col, cell } = target;
    if ((cell.colSpan ?? 1) <= 1 && (cell.rowSpan ?? 1) <= 1) return null;
    const grid = buildTableGrid(table);
    const topLeft: TableCell = {
      id: cell.id,
      blocks: cell.blocks,
      ...(cell.shading !== undefined ? { shading: cell.shading } : {}),
      ...(cell.borders !== undefined ? { borders: cell.borders } : {}),
      ...(cell.margin !== undefined ? { margin: cell.margin } : {}),
    };
    const makeEmpty = (): TableCell => ({
      id: freshBlockId(),
      blocks: [emptyCellPara(firstCellPara(cell))],
      ...(cell.margin !== undefined ? { margin: cell.margin } : {}),
    });
    const rows = unmergeRows(grid, row, col, topLeft, makeEmpty);
    const caretPara = firstCellPara(cell);
    return tr(
      [structureOp(table, rows)],
      caretPara ? caret(caretPara.id, 0) : null,
      "command",
    );
  };
}

/** A setTableStructure op preserving the table's column fractions (cell merges
 *  and border edits never change the grid-column count). */
function structureOp(table: TableBlock, rows: TableRow[]): Op {
  return table.colFractions
    ? { type: "setTableStructure", tableId: table.id, rows, colFractions: table.colFractions }
    : { type: "setTableStructure", tableId: table.id, rows };
}

// ---------------------------------------------------------------------------
// Borders & shading (Word's "Borders and Shading")

/** Which roles of a cell-rectangle to target. Outer roles (top/right/bottom/
 *  left) are the selection's perimeter; insideH/insideV are the shared edges
 *  between selected cells. The modal's presets map onto these:
 *  All = every flag, Outside = the four sides, Inside = insideH+insideV, plus
 *  individual toggles. */
export interface BorderEdgeFlags {
  top?: boolean;
  right?: boolean;
  bottom?: boolean;
  left?: boolean;
  insideH?: boolean;
  insideV?: boolean;
}

/** Apply (spec) or clear (spec=null) borders on the targeted edges of every cell
 *  in the range. Each cell's physical edge is an OUTER edge when it lies on the
 *  selection perimeter and an INTERIOR edge otherwise, so the same flag set drives
 *  any cell shape (incl. merged cells spanning several grid tracks). Touched
 *  cells keep a present borders object (never reverting to the default grid). */
export function setCellsBordersCmd(spec: CellBorder | null, edges: BorderEdgeFlags, range?: CellSelection | null): Command {
  return (state) => {
    const found = cellRangeFromState(state, range);
    if (!found) return null;
    const grid = buildTableGrid(found.table);
    const rect = normalizeRect(grid, found.rect);
    let changed = false;
    const rows = rebuildRows(grid, (cell, s) => {
      const top = s.originRow;
      const bottom = s.originRow + s.rowSpan - 1;
      const left = s.originCol;
      const right = s.originCol + s.colSpan - 1;
      if (top < rect.r0 || bottom > rect.r1 || left < rect.c0 || right > rect.c1) return cell; // outside the rectangle
      const targetTop = top === rect.r0 ? edges.top : edges.insideH;
      const targetBottom = bottom === rect.r1 ? edges.bottom : edges.insideH;
      const targetLeft = left === rect.c0 ? edges.left : edges.insideV;
      const targetRight = right === rect.c1 ? edges.right : edges.insideV;
      if (!targetTop && !targetBottom && !targetLeft && !targetRight) return cell;
      const borders: CellBorders = { ...(cell.borders ?? {}) };
      const apply = (key: keyof CellBorders, on: boolean | undefined): void => {
        if (!on) return;
        if (spec) borders[key] = { ...spec };
        else delete borders[key];
      };
      apply("top", targetTop);
      apply("bottom", targetBottom);
      apply("left", targetLeft);
      apply("right", targetRight);
      changed = true;
      return { ...cell, borders };
    });
    if (!changed) return null;
    return tr([structureOp(found.table, rows)], state.selection, "command");
  };
}

/** Set (fill) or clear (fill=null) the background shading of every cell in the
 *  range. */
export function setCellsShadingCmd(fill: string | null, range?: CellSelection | null): Command {
  return (state) => {
    const found = cellRangeFromState(state, range);
    if (!found) return null;
    const grid = buildTableGrid(found.table);
    const rect = normalizeRect(grid, found.rect);
    let changed = false;
    const rows = rebuildRows(grid, (cell, s) => {
      if (s.originRow < rect.r0 || s.originRow > rect.r1 || s.originCol < rect.c0 || s.originCol > rect.c1) return cell;
      changed = true;
      if (fill === null) {
        if (cell.shading === undefined) return cell;
        const { shading: _drop, ...rest } = cell;
        return rest;
      }
      return { ...cell, shading: fill };
    });
    if (!changed) return null;
    return tr([structureOp(found.table, rows)], state.selection, "command");
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
// Table styles — round-trippable TableStyle entities whose effective per-cell
// formatting is BAKED onto the cells (the engine reads concrete cell props).

/** Every TableBlock in the document body (recursing into cells). */
function allTables(blocks: Block[], out: TableBlock[] = []): TableBlock[] {
  for (const b of blocks) {
    if (b.kind === "table") {
      out.push(b);
      for (const row of b.rows) for (const cell of row.cells) allTables(cell.blocks, out);
    }
  }
  return out;
}

/** Apply a table style to the table at the caret: set the reference + default
 *  tblLook and bake the effective cell formatting. */
export function applyTableStyle(styleId: string): Command {
  return (state) => {
    const ctx = cellContext(state);
    if (!ctx) return null;
    const styles = state.doc.tableStyles ?? {};
    const style = styles[styleId];
    if (!style) return null;
    const overrides = ctx.table.condOverrides ?? DEFAULT_TBL_LOOK;
    const rows = bakeTableStyleRows(ctx.table, style, styles, overrides);
    const ops: Op[] = [
      { type: "setTableStyleRef", tableId: ctx.table.id, styleId, condOverrides: overrides },
      structureOp(ctx.table, rows),
    ];
    return tr(ops, state.selection, "command");
  };
}

/** Whether table style `id`'s basedOn ancestry includes `targetId` (cycle-guarded,
 *  like `resolveTableStyle`). */
function tableStyleChainIncludes(styles: Record<string, TableStyle>, id: string, targetId: string): boolean {
  const seen = new Set<string>();
  for (let cur = styles[id]; cur && !seen.has(cur.id); cur = cur.basedOn ? styles[cur.basedOn] : undefined) {
    if (cur.id === targetId) return true;
    seen.add(cur.id);
  }
  return false;
}

/** Create or replace a table style, then re-bake every table whose resolved style
 *  chain includes it — including tables referencing a style DERIVED from it via
 *  `basedOn`, whose baked cells would otherwise stay stale until reapplied.
 *  One undoable transaction (setTableStyleSheet + per-table structure ops). */
export function upsertTableStyle(style: TableStyle): Command {
  return (state) => {
    const styles = { ...(state.doc.tableStyles ?? {}), [style.id]: style };
    const ops: Op[] = [{ type: "setTableStyleSheet", tableStyles: styles }];
    for (const table of allTables(state.doc.blocks)) {
      if (!table.styleId) continue;
      const tableStyle = styles[table.styleId];
      if (!tableStyle || !tableStyleChainIncludes(styles, table.styleId, style.id)) continue;
      const overrides = table.condOverrides ?? DEFAULT_TBL_LOOK;
      // Bake with the table's OWN style (so its basedOn chain resolves correctly),
      // not the edited style — they differ for basedOn-derived references.
      ops.push(structureOp(table, bakeTableStyleRows(table, tableStyle, styles, overrides)));
    }
    return tr(ops, state.selection, "command");
  };
}

/** Delete a table style; referencing tables keep their baked cells as direct
 *  formatting but drop the reference. */
export function deleteTableStyle(styleId: string): Command {
  return (state) => {
    const styles = { ...(state.doc.tableStyles ?? {}) };
    if (!styles[styleId]) return null;
    delete styles[styleId];
    const ops: Op[] = [{ type: "setTableStyleSheet", tableStyles: styles }];
    for (const table of allTables(state.doc.blocks)) {
      if (table.styleId === styleId) ops.push({ type: "setTableStyleRef", tableId: table.id, styleId: null, condOverrides: null });
    }
    return tr(ops, state.selection, "command");
  };
}

// ---------------------------------------------------------------------------
// Image commands (target = the SELECTED object, passed in by the wiring layer)

/** Replace a registered custom block's `data` (undoable). Bumps the block's
 *  revision so only it re-lays-out/repaints. Powers async data (fetch → update). */
export function setCustomBlockData(
  blockId: string,
  data: unknown,
  origin: TransactionOrigin = "command",
): Command {
  return (state) => {
    if (!locateBlock(state.doc, blockId, "custom")) return null;
    return tr([{ type: "setCustomBlockData", blockId, data }], state.selection, origin);
  };
}

export function setImageProps(
  blockId: string,
  patch: ImagePropsPatch,
  origin: TransactionOrigin = "command",
): Command {
  return (state) => {
    // Cell-aware: images inside table cells (common in content-control report
    // templates) must resize/align too. The reducer's op already handles cells —
    // a top-level-only `doc.blocks` guard here was what made them inert.
    if (!locateImage(state.doc, blockId)) return null;
    return tr([{ type: "setImageProps", blockId, patch }], state.selection, origin);
  };
}

/** Set a drawing shape's props (size / align / fill / stroke). Cell-aware, like
 *  setImageProps: shapes in table cells resize + align through the reducer op. */
export function setShapeProps(
  blockId: string,
  patch: ShapePropsPatch,
  origin: TransactionOrigin = "command",
): Command {
  return (state) => {
    if (!locateShape(state.doc, blockId)) return null;
    return tr([{ type: "setShapeProps", blockId, patch }], state.selection, origin);
  };
}

/** Set (or clear, with `null`) the SELECTED image's crop insets (a:srcRect 0..1
 *  fractions). The interactive crop tool previews the new window purely in the DOM
 *  overlay and commits ONE op here on exit — same single-undo-step protocol as
 *  drag-to-resize. Reuses the `setImageProps` op (its inverse restores the prior
 *  crop), so no new op is needed. */
export function setImageCropCmd(
  blockId: string,
  crop: NonNullable<ImageBlock["crop"]> | null,
  origin: TransactionOrigin = "command",
): Command {
  return (state) => {
    if (!locateImage(state.doc, blockId)) return null;
    return tr([{ type: "setImageProps", blockId, patch: { crop } }], state.selection, origin);
  };
}

type ImageAnchor = NonNullable<ImageBlock["anchor"]>;

/** Anchor for an image being lifted out of the flow into the behind/front layer:
 *  keep its existing anchor (just flip the layer) or seed one at the margin. */
function anchorFor(img: ImageBlock, behind: boolean): ImageAnchor {
  return img.anchor
    ? { ...img.anchor, behind }
    : { behind, offsetXPx: 0, offsetYPx: 0, relFromH: "margin", relFromV: "paragraph" };
}

/** z extremes across every anchored image in the document (body + cells). */
function anchorZRange(doc: EditorState["doc"]): { min: number; max: number } {
  let min = 0;
  let max = 0;
  const walk = (blocks: Block[]): void => {
    for (const b of blocks) {
      if (b.kind === "image" && b.anchor) {
        const z = b.anchor.z ?? 0;
        if (z < min) min = z;
        if (z > max) max = z;
      } else if (b.kind === "table") {
        for (const row of b.rows) for (const cell of row.cells) walk(cell.blocks);
      }
    }
  };
  walk(doc.blocks);
  return { min, max };
}

/** Toggle an image between the behind-text and in-front-of-text layers (lifting
 *  an in-line image into an anchored one). Clears `wrap` — the two are exclusive. */
export function setImageLayer(blockId: string, behind: boolean): Command {
  return (state) => {
    const loc = locateImage(state.doc, blockId);
    if (!loc) return null;
    const anchor = anchorFor(loc.image, behind);
    return tr([{ type: "setImageProps", blockId, patch: { anchor, wrap: null } }], state.selection, "command");
  };
}

/** Raise an image above every other object AND the text (in-front layer, top z). */
export function bringImageToFront(blockId: string): Command {
  return (state) => {
    const loc = locateImage(state.doc, blockId);
    if (!loc) return null;
    const anchor: ImageAnchor = { ...anchorFor(loc.image, false), z: anchorZRange(state.doc).max + 1 };
    return tr([{ type: "setImageProps", blockId, patch: { anchor, wrap: null } }], state.selection, "command");
  };
}

/** Drop an image below every other object AND the text (behind layer, bottom z). */
export function sendImageToBack(blockId: string): Command {
  return (state) => {
    const loc = locateImage(state.doc, blockId);
    if (!loc) return null;
    const anchor: ImageAnchor = { ...anchorFor(loc.image, true), z: anchorZRange(state.doc).min - 1 };
    return tr([{ type: "setImageProps", blockId, patch: { anchor, wrap: null } }], state.selection, "command");
  };
}

/** Reposition an anchored image to absolute anchor offsets (drag-move). Pass
 *  origin "transient" for the live drag preview and "command" for the committed
 *  drop — the same revert-then-commit dance as resize keeps a drag = one undo step. */
export function moveAnchoredImage(
  blockId: string,
  offsetXPx: number,
  offsetYPx: number,
  origin: TransactionOrigin = "command",
): Command {
  return (state) => {
    const loc = locateImage(state.doc, blockId);
    if (!loc?.image.anchor) return null; // only anchored (out-of-flow) images move
    const anchor: ImageAnchor = { ...loc.image.anchor, offsetXPx, offsetYPx };
    return tr([{ type: "setImageProps", blockId, patch: { anchor } }], state.selection, origin);
  };
}

/** Remove a display-equation block object — wherever it lives (body, header/footer
 *  band, or table cell) — and park the caret on a neighbour. For images use
 *  deleteImage (it also handles in-cell images). */
export function removeBlockObject(blockId: string): Command {
  return (state) => {
    // Object-selectable non-image blocks: display equations, drawing shapes, and
    // registered custom blocks. All remove identically (top-level removeBlock, or a
    // row rewrite in a cell) — only the location matters, not the block type.
    const loc =
      locateEquation(state.doc, blockId) ??
      locateShape(state.doc, blockId) ??
      locateBlock(state.doc, blockId, "custom");
    if (!loc) return null;
    if (loc.kind === "cell") {
      // In-cell equations can't go through removeBlock (it scans top-level only);
      // rewrite the row with the equation filtered out, like deleteImage.
      const table = containerBlocks(state.doc, loc.where)[loc.bi] as TableBlock;
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
    const blocks = containerBlocks(state.doc, loc.where);
    const neighbour = blocks[loc.index + 1] ?? blocks[loc.index - 1];
    return tr([{ type: "removeBlock", blockId }], neighbour ? caret(neighbour.id, 0) : state.selection, "command");
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
export function insertImageInCell(src: string, widthPx: number, heightPx: number, mediaId?: string): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel || !isCollapsed(sel)) return null;
    const loc = locateParagraph(state.doc, sel.focus.blockId);
    if (loc?.kind !== "cell") return null;
    const table = containerBlocks(state.doc, loc.where)[loc.bi] as TableBlock;
    const row = table.rows[loc.ri]!;
    const cell = row.cells[loc.ci]!;
    const img: ImageBlock = {
      kind: "image",
      id: freshBlockId(),
      revision: 0,
      src,
      widthPx,
      heightPx,
      align: "center",
    };
    // Content address so the bytes survive serialize/export (see insertImage).
    if (mediaId) img.mediaId = mediaId;
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

/** Set (or clear, with `null`) a single row's fixed/min height (w:trHeight). Drives
 *  the interactive row-boundary drag — `"transient"` for live preview frames, then a
 *  final committed op — mirroring {@link setTableColFractionsCmd}. */
export function setRowHeightCmd(
  tableId: string,
  rowIndex: number,
  height: NonNullable<NonNullable<TableRow["props"]>["height"]> | null,
  origin: TransactionOrigin = "command",
): Command {
  return (state) =>
    tr([{ type: "setRowHeight", tableId, rowIndex, height }], state.selection, origin);
}

/** Set a table's column-sizing mode. When switching to "fixed" (or cancelling
 *  autofit via a column drag), pass `freezeFractions` to pin the currently-
 *  rendered widths in the SAME transaction, so one undo reverts both. */
export function setTableWidthModeCmd(
  tableId: string,
  mode: TableBlock["widthMode"],
  freezeFractions?: number[],
  origin: TransactionOrigin = "command",
): Command {
  return (state) => {
    const ops: Op[] = [];
    if (freezeFractions) ops.push({ type: "setTableColFractions", blockId: tableId, fractions: freezeFractions });
    ops.push({ type: "setTableWidthMode", blockId: tableId, mode });
    return tr(ops, state.selection, origin);
  };
}

/** Set the column-sizing mode of the table containing the caret (ribbon/menu). */
export function setTableWidthModeAtSelectionCmd(mode: TableBlock["widthMode"]): Command {
  return (state) => {
    const ctx = cellContext(state);
    if (!ctx) return null;
    return tr([{ type: "setTableWidthMode", blockId: ctx.table.id, mode }], state.selection, "command");
  };
}

/** Set the preferred total width of the table containing the caret. A preferred
 *  width only applies in fixed layout, so this also switches the table to "fixed"
 *  (cancelling any AutoFit) in the SAME transaction — one undo reverts both. Pass
 *  null to clear the preference and span the full content width again. */
export function setTablePreferredWidthAtSelectionCmd(width: TableBlock["preferredWidth"] | null): Command {
  return (state) => {
    const ctx = cellContext(state);
    if (!ctx) return null;
    const ops: Op[] = [];
    if (width && ctx.table.widthMode && ctx.table.widthMode !== "fixed") {
      ops.push({ type: "setTableWidthMode", blockId: ctx.table.id, mode: "fixed" });
    }
    ops.push({ type: "setTablePreferredWidth", blockId: ctx.table.id, width });
    return tr(ops, state.selection, "command");
  };
}

/** Set the horizontal alignment of the table containing the caret (left/center/right). */
export function setTableAlignAtSelectionCmd(align: TableBlock["align"] | null): Command {
  return (state) => {
    const ctx = cellContext(state);
    if (!ctx) return null;
    return tr([{ type: "setTableAlign", blockId: ctx.table.id, align }], state.selection, "command");
  };
}

/** The table the caret/cell-selection sits in, or null. Lets the UI reflect the
 *  current table's state (e.g. tick the active AutoFit mode). */
export function tableAtSelection(state: EditorState): TableBlock | null {
  return cellContext(state)?.table ?? null;
}

// ---------------------------------------------------------------------------
// Table Properties — cell vAlign / text direction, row props, table-level props
// (issue #86). All reuse the existing setTableStructure / setRowHeight / a single
// new setTableProps op, applied over the captured selection with free undo.

/** rebuildRows that carries each source row's `props` across (the shared helper
 *  emits `{ cells }` only). Cell-property edits must NOT drop a row's height /
 *  cant-split / repeat-header. Grid rows map 1:1 to table rows. */
function rebuildRowsKeepProps(
  table: TableBlock,
  grid: TableGrid,
  mapCell: (cell: TableCell, slot: GridSlot) => TableCell,
): TableRow[] {
  return rebuildRows(grid, mapCell).map((row, ri) => {
    const props = table.rows[ri]?.props;
    return props ? { ...row, props } : row;
  });
}

/** The grid rectangle to act on: an explicit `override` / live cell-selection /
 *  straddling text selection (via cellRangeFromState), else the single cell at
 *  the caret. */
function cellRectAtSelection(
  state: EditorState,
  override?: CellSelection | null,
): { table: TableBlock; rect: GridRect } | null {
  const range = cellRangeFromState(state, override);
  if (range) return range;
  const ctx = cellContext(state);
  if (!ctx) return null;
  const grid = buildTableGrid(ctx.table);
  const origin = gridOriginOfCell(grid, ctx.ri, ctx.ci);
  if (!origin) return null;
  return { table: ctx.table, rect: { r0: origin.row, c0: origin.col, r1: origin.row, c1: origin.col } };
}

/** The inclusive table-row range covered by the selection (grid rows == table
 *  rows), or the caret's single row. */
function selectedRowRange(
  state: EditorState,
  override?: CellSelection | null,
): { table: TableBlock; r0: number; r1: number } | null {
  const found = cellRectAtSelection(state, override);
  if (!found) return null;
  const grid = buildTableGrid(found.table);
  const rect = normalizeRect(grid, found.rect);
  return { table: found.table, r0: rect.r0, r1: rect.r1 };
}

/** Set (or clear, with `null`/"top") the vertical alignment of every cell in the
 *  range — w:tcPr/w:vAlign. "top" is Word's default, so it clears the field. */
export function setCellVAlignCmd(
  vAlign: NonNullable<TableCell["vAlign"]> | null,
  range?: CellSelection | null,
): Command {
  return (state) => {
    const found = cellRectAtSelection(state, range);
    if (!found) return null;
    const grid = buildTableGrid(found.table);
    const rect = normalizeRect(grid, found.rect);
    let changed = false;
    const clear = vAlign === null || vAlign === "top";
    const rows = rebuildRowsKeepProps(found.table, grid, (cell, s) => {
      if (s.originRow < rect.r0 || s.originRow > rect.r1 || s.originCol < rect.c0 || s.originCol > rect.c1) return cell;
      if (clear) {
        if (cell.vAlign === undefined) return cell;
        const { vAlign: _drop, ...rest } = cell;
        changed = true;
        return rest;
      }
      if (cell.vAlign === vAlign) return cell;
      changed = true;
      return { ...cell, vAlign };
    });
    if (!changed) return null;
    return tr([structureOp(found.table, rows)], state.selection, "command");
  };
}

/** Set (or clear, with `null`/"lrTb") the text-flow direction of every cell in
 *  the range — w:tcPr/w:textDirection. "lrTb" is the default, so it clears. */
export function setCellTextDirectionCmd(
  dir: NonNullable<TableCell["textDirection"]> | null,
  range?: CellSelection | null,
): Command {
  return (state) => {
    const found = cellRectAtSelection(state, range);
    if (!found) return null;
    const grid = buildTableGrid(found.table);
    const rect = normalizeRect(grid, found.rect);
    let changed = false;
    const clear = dir === null || dir === "lrTb";
    const rows = rebuildRowsKeepProps(found.table, grid, (cell, s) => {
      if (s.originRow < rect.r0 || s.originRow > rect.r1 || s.originCol < rect.c0 || s.originCol > rect.c1) return cell;
      if (clear) {
        if (cell.textDirection === undefined) return cell;
        const { textDirection: _drop, ...rest } = cell;
        changed = true;
        return rest;
      }
      if (cell.textDirection === dir) return cell;
      changed = true;
      return { ...cell, textDirection: dir };
    });
    if (!changed) return null;
    return tr([structureOp(found.table, rows)], state.selection, "command");
  };
}

/** Toggle the boolean row props (w:cantSplit / w:tblHeader) on every row the
 *  selection covers. A falsy patch value clears the flag; an empty resulting
 *  `props` object is dropped. Reuses setTableStructure (cells are untouched, so
 *  carets survive). */
export function setRowPropsCmd(
  patch: { cantSplit?: boolean; repeatHeader?: boolean },
  range?: CellSelection | null,
): Command {
  return (state) => {
    const rr = selectedRowRange(state, range);
    if (!rr) return null;
    let changed = false;
    const rows = rr.table.rows.map((row, ri) => {
      if (ri < rr.r0 || ri > rr.r1) return row;
      const props: RowProps = { ...(row.props ?? {}) };
      let rowChanged = false;
      for (const key of Object.keys(patch) as (keyof typeof patch)[]) {
        const want = patch[key] ? true : undefined;
        if (want === undefined) {
          if (props[key] !== undefined) { delete props[key]; rowChanged = true; }
        } else if (props[key] !== true) { props[key] = true; rowChanged = true; }
      }
      if (!rowChanged) return row;
      changed = true;
      if (Object.keys(props).length === 0) {
        const { props: _drop, ...rest } = row;
        return rest;
      }
      return { ...row, props };
    });
    if (!changed) return null;
    return tr([structureOp(rr.table, rows)], state.selection, "command");
  };
}

/** Set (or clear, with `null`) the fixed/min height of every row the selection
 *  covers — one setRowHeight op per row in a single undoable transaction. */
export function setRowHeightAtSelectionCmd(
  height: NonNullable<RowProps["height"]> | null,
  range?: CellSelection | null,
): Command {
  return (state) => {
    const rr = selectedRowRange(state, range);
    if (!rr) return null;
    const sameHeight = (
      a: NonNullable<RowProps["height"]> | null,
      b: NonNullable<RowProps["height"]> | null,
    ): boolean => (a === b ? true : !!a && !!b && a.value === b.value && a.rule === b.rule);
    const ops: Op[] = [];
    for (let ri = rr.r0; ri <= rr.r1; ri++) {
      const cur = rr.table.rows[ri]?.props?.height ?? null;
      if (sameHeight(cur, height)) continue;
      ops.push({ type: "setRowHeight", tableId: rr.table.id, rowIndex: ri, height });
    }
    if (ops.length === 0) return null;
    return tr(ops, state.selection, "command");
  };
}

/** Resolve the table for a table-level prop edit: the caret's cell, else the
 *  live cell selection's table. */
function tableForPropsEdit(state: EditorState): TableBlock | null {
  const ctx = cellContext(state);
  if (ctx) return ctx.table;
  const cs = state.cellSelection;
  return cs ? findTableById(state.doc, cs.tableId)?.table ?? null : null;
}

/** Patch table-LEVEL props (indent + cascade defaults) of the caret/selection's
 *  table — w:tblPr/w:tblInd, w:tblBorders, w:shd, w:tblCellMar. One setTableProps
 *  op; `null` in the patch clears a field. */
export function setTablePropsAtSelectionCmd(patch: TablePropsPatch): Command {
  return (state) => {
    if (Object.keys(patch).length === 0) return null;
    const table = tableForPropsEdit(state);
    if (!table) return null;
    return tr([{ type: "setTableProps", blockId: table.id, patch }], state.selection, "command");
  };
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

/** Reset character formatting over the selection to plain text: drop the
 *  toggleable runs of style (bold/italic/…), colour and highlight, baseline
 *  shift, letter spacing and link. Font family/size are left alone — Word's
 *  "Clear All Formatting" also reverts the paragraph to Normal, which the caller
 *  pairs with applyNamedStyle("Normal"). No-op on a collapsed caret. */
export function clearCharFormatting(): Command {
  return setCharStyle({
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    color: "#202124",
    highlightColor: undefined,
    verticalAlign: undefined,
    letterSpacingPx: 0,
    link: undefined,
  });
}

export type CaseMode = "upper" | "lower" | "sentence" | "title" | "toggle";

const applyCase = (s: string, mode: CaseMode): string => {
  switch (mode) {
    case "upper":
      return s.toUpperCase();
    case "lower":
      return s.toLowerCase();
    case "toggle":
      return [...s]
        .map((c) => (c === c.toLowerCase() ? c.toUpperCase() : c.toLowerCase()))
        .join("");
    case "title":
      return s.replace(/\p{L}[\p{L}'’]*/gu, (w) => w[0]!.toUpperCase() + w.slice(1).toLowerCase());
    case "sentence": {
      const lower = s.toLowerCase();
      // Capitalise the first letter, and the first letter after . ! ? …
      return lower.replace(/(^|[.!?…]\s+)(\p{L})/gu, (_, lead: string, ch: string) => lead + ch.toUpperCase());
    }
  }
};

/** Transform the case of the selected text, preserving every run's style. */
export function changeCaseCmd(mode: CaseMode): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel || isCollapsed(sel)) return null;
    const segments = selectedSegments(state).filter((s) => s.end > s.start);
    if (segments.length === 0) return null;
    const ops: Op[] = segments.map((s) => ({
      type: "setRuns",
      blockId: s.block.id,
      runs: mapTextInRuns(s.block.runs, s.start, s.end, (t) => applyCase(t, mode)),
    }));
    return tr(ops, sel, "command");
  };
}

/** Add (or remove, with a negative delta) left indent on every covered
 *  paragraph, clamped at zero. Mirrors Word's Increase/Decrease Indent. */
export function adjustIndentCmd(deltaPx: number): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel) return null;
    const ops: Op[] = selectedSegments(state).map((s) => ({
      type: "setParaStyle",
      blockId: s.block.id,
      patch: { indentLeftPx: Math.max(0, (s.block.style.indentLeftPx ?? 0) + deltaPx) },
    }));
    return ops.length ? tr(ops, sel, "command") : null;
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
  mergeDuplicateNamedStyles,
  paragraphsWithStyle,
  resolveCharStyle,
  resolveStyle,
  styleById,
  styleType,
  wouldCycle,
  type NamedStyle,
  type NamedStyleType,
  type Stylesheet,
} from "@cw/shared";

/** A style definition the UI hands to upsertNamedStyle. `id` present ⇒ update in
 *  place; absent ⇒ create (id derived from the name). */
export interface NamedStyleSpec {
  id?: string;
  name: string;
  type: NamedStyleType;
  basedOn?: string;
  char?: Partial<CharStyle>;
  para?: Partial<ParaStyle>;
}

/** Every paragraph holding a run that references character style `id`. */
function paragraphsWithCharStyle(doc: EditorState["doc"], styleId: string): Paragraph[] {
  return paragraphsOf(doc).filter((p) => p.runs.some((r) => r.style.charStyleId === styleId));
}

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

/** Apply a CHARACTER style to the selected run range only (not the paragraph):
 *  bake the style's resolved char props over the runs and tag them with its id
 *  (w:rStyle). No-op on a collapsed caret or a non-character style. */
export function applyCharStyle(styleId: string): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel || isCollapsed(sel)) return null;
    const sheet = sheetOf(state);
    const st = styleById(sheet, styleId);
    if (!st || styleType(st) !== "character") return null;
    const patch: Partial<CharStyle> = { ...resolveCharStyle(sheet, styleId), charStyleId: styleId };
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

/** Re-patch only the runs of `block` that reference character style `styleId`,
 *  merging the resolved `char` over them (leaving the paragraph and other runs
 *  untouched). The charStyleId marker is preserved. */
function restyleCharRuns(block: Paragraph, char: Partial<CharStyle>, styleId: string): Op {
  const runs = block.runs.map((r) =>
    r.style.charStyleId === styleId ? { ...r, style: { ...r.style, ...char, charStyleId: styleId } } : r,
  );
  return { type: "setRuns", blockId: block.id, runs };
}

/** Create-or-update a named style from a spec, then RE-BAKE every piece of content
 *  referencing it so the visible formatting tracks the (re)definition — paragraphs
 *  for paragraph styles, runs for character styles. The whole thing is one undoable
 *  transaction (setStylesheet + the re-patch ops). The single generalized creator;
 *  createStyleFromSelection / updateStyleToSelection are thin wrappers over it. */
export function upsertNamedStyle(spec: NamedStyleSpec): Command {
  return (state) => {
    const sel = state.selection;
    const name = spec.name.trim();
    if (name.length === 0) return null;
    const sheet = sheetOf(state);
    const isCharacter = spec.type === "character";
    const existing = spec.id ? styleById(sheet, spec.id) : undefined;
    const id = spec.id ?? name.replace(/\s+/g, "");
    if (!existing && styleById(sheet, id)) return null; // creating but id taken
    if (wouldCycle(sheet, id, spec.basedOn)) return null;

    const style: NamedStyle = {
      id,
      name,
      type: spec.type,
      char: { ...(spec.char ?? {}) },
      para: isCharacter ? {} : { ...(spec.para ?? {}) },
    };
    if (spec.basedOn) style.basedOn = spec.basedOn;

    const styles = existing
      ? sheet.styles.map((s) => (s.id === id ? style : s))
      : [...sheet.styles, style];
    const stylesheet: Stylesheet = { ...sheet, styles };

    const ops: Op[] = [{ type: "setStylesheet", stylesheet }];
    if (isCharacter) {
      const char = resolveCharStyle(stylesheet, id);
      for (const p of paragraphsWithCharStyle(state.doc, id)) ops.push(restyleCharRuns(p, char, id));
    } else {
      const { char, para } = resolveStyle(stylesheet, id);
      const seen = new Set<string>();
      for (const p of paragraphsWithStyle(state.doc, id)) {
        seen.add(p.id);
        ops.push(...restyleOps(p, char, para));
      }
      // A brand-new paragraph style also applies to the current selection (like
      // Word's "New style from selection").
      if (!existing && sel) {
        for (const s of selectedSegments(state)) {
          if (seen.has(s.block.id)) continue;
          seen.add(s.block.id);
          ops.push(...restyleOps(s.block, char, para, id));
        }
      }
    }
    return tr(ops, sel, "command");
  };
}

/** Word's "Update <style> to Match Selection": redefine the style from the caret
 *  formatting (self-contained — drops basedOn), then re-bake referencing content. */
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
    if (styleType(existing) === "character") {
      return upsertNamedStyle({ id: styleId, name: existing.name, type: "character", char })(state);
    }
    const { namedStyle: _omit, ...para } = block.style;
    return upsertNamedStyle({ id: styleId, name: existing.name, type: "paragraph", char, para })(state);
  };
}

/** Create a new paragraph style from the caret's formatting and apply it. */
export function createStyleFromSelection(name: string): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel || name.trim().length === 0) return null;
    const block = blockById(state.doc, sel.focus.blockId);
    if (!block) return null;
    const char = styleAtRuns(block.runs, sel.focus.offset);
    if (!char) return null;
    const { namedStyle: _omit, ...para } = block.style;
    return upsertNamedStyle({ name, type: "paragraph", char, para })(state);
  };
}

/** Rename a style's display name only. The id is the docx key referenced by
 *  content, so it never changes — no re-patch needed. */
export function renameNamedStyle(styleId: string, newName: string): Command {
  return (state) => {
    const sel = state.selection;
    const sheet = sheetOf(state);
    const existing = styleById(sheet, styleId);
    if (!existing || newName.trim().length === 0) return null;
    const stylesheet: Stylesheet = {
      ...sheet,
      styles: sheet.styles.map((s) => (s.id === styleId ? { ...s, name: newName.trim() } : s)),
    };
    return tr([{ type: "setStylesheet", stylesheet }], sel, "command");
  };
}

/** Make `styleId` the document's default (Word's w:default="1"). */
export function setDefaultStyle(styleId: string): Command {
  return (state) => {
    const sel = state.selection;
    const sheet = sheetOf(state);
    if (!styleById(sheet, styleId)) return null;
    if (sheet.defaultStyleId === styleId) return null;
    return tr([{ type: "setStylesheet", stylesheet: { ...sheet, defaultStyleId: styleId } }], sel, "command");
  };
}

/** Delete a named style. Content referencing it falls back to the style's
 *  basedOn (if a still-present style of the same type) else the default style,
 *  and is RE-PATCHED so its visible formatting matches the new reference.
 *  basedOn chains pointing at the deleted style are rewritten to its parent.
 *  The default style cannot be deleted. */
export function deleteNamedStyle(styleId: string): Command {
  return (state) => {
    const sel = state.selection;
    const sheet = sheetOf(state);
    const victim = styleById(sheet, styleId);
    if (!victim || sheet.defaultStyleId === styleId) return null;

    const parent = victim.basedOn ? styleById(sheet, victim.basedOn) : undefined;
    // Same-type fallback keeps semantics (a paragraph keeps a paragraph style).
    const fallbackId =
      parent && styleType(parent) === styleType(victim) ? parent.id : sheet.defaultStyleId;

    const styles = sheet.styles
      .filter((s) => s.id !== styleId)
      // Reparent any child that based on the victim onto the victim's parent.
      .map((s) => {
        if (s.basedOn !== styleId) return s;
        const { basedOn: _drop, ...rest } = s;
        return victim.basedOn !== undefined ? { ...rest, basedOn: victim.basedOn } : rest;
      });
    const stylesheet: Stylesheet = { ...sheet, styles };

    const ops: Op[] = [{ type: "setStylesheet", stylesheet }];
    if (styleType(victim) === "character") {
      // Strip the reference from runs that carried it (formatting stays baked).
      for (const block of paragraphsOf(state.doc)) {
        if (!block.runs.some((r) => r.style.charStyleId === styleId)) continue;
        const runs = block.runs.map((r) =>
          r.style.charStyleId === styleId ? { ...r, style: { ...r.style, charStyleId: undefined } } : r,
        );
        ops.push({ type: "setRuns", blockId: block.id, runs });
      }
    } else {
      const { char, para } = resolveStyle(stylesheet, fallbackId);
      for (const p of paragraphsWithStyle(state.doc, styleId)) {
        ops.push(...restyleOps(p, char, para, fallbackId));
      }
    }
    return tr(ops, sel, "command");
  };
}

/** Collapse named styles that are identical except for their name (e.g. importer
 *  duplicates), remapping content references to the surviving style. Since merged
 *  styles share identical formatting, only the references move — no re-bake needed.
 *  No-op (returns null) when there are no duplicates. One undoable transaction. */
export function mergeDuplicateStyles(): Command {
  return (state) => {
    const sel = state.selection;
    const sheet = sheetOf(state);
    const { sheet: merged, remap } = mergeDuplicateNamedStyles(sheet);
    if (remap.size === 0) return null;
    const ops: Op[] = [{ type: "setStylesheet", stylesheet: merged }];
    for (const p of paragraphsOf(state.doc)) {
      const ns = p.style.namedStyle;
      if (ns !== undefined && remap.has(ns)) {
        ops.push({ type: "setParaStyle", blockId: p.id, patch: { namedStyle: remap.get(ns)! } });
      }
      if (p.runs.some((r) => r.style.charStyleId !== undefined && remap.has(r.style.charStyleId))) {
        const runs = p.runs.map((r) =>
          r.style.charStyleId !== undefined && remap.has(r.style.charStyleId)
            ? { ...r, style: { ...r.style, charStyleId: remap.get(r.style.charStyleId)! } }
            : r,
        );
        ops.push({ type: "setRuns", blockId: p.id, runs });
      }
    }
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

/** Set the base writing direction (LTR/RTL, OOXML w:bidi) on the selected
 *  paragraphs. RTL right-aligns by default and lays text out right-to-left. */
export function setDirection(direction: "ltr" | "rtl"): Command {
  return (state) => {
    const sel = state.selection;
    if (!sel) return null;
    const [from, to] = orderedRange(state, sel);
    const fi = blockIndexOf(state.doc, from.blockId);
    const ti = blockIndexOf(state.doc, to.blockId);
    const blocks = paragraphsOf(state.doc);
    const ops: Op[] = [];
    for (let i = fi; i <= ti; i++) {
      ops.push({ type: "setParaStyle", blockId: blocks[i]!.id, patch: { direction } });
    }
    return tr(ops, sel, "command");
  };
}

// ---------------------------------------------------------------------------
// Page groups (visual "Organize Pages" reorder). A page group is a contiguous
// range of body blocks delimited by a hard break (page break / section break);
// see ../layout/pageGroups. These wrap the pure ops into undoable transactions.

/** Move the movable unit [firstId…lastId] so it lands immediately before
 *  `anchorId` (or the document end when null). No-op transaction for an in-place
 *  or invalid move. */
export function reorderPageGroupCmd(firstId: string, lastId: string, anchorId: string | null): Command {
  return (state) => {
    const ops = reorderPageGroupsOps(state.doc, firstId, lastId, anchorId);
    return ops.length ? tr(ops, state.selection, "command") : null;
  };
}

/** Delete the movable unit [firstId…lastId] (its whole block range). No-op when
 *  that would empty the document. Clears the selection since it may have pointed
 *  into the removed blocks. */
export function deletePageGroupCmd(firstId: string, lastId: string): Command {
  return (state) => {
    const ops = deletePageGroupOps(state.doc, firstId, lastId);
    return ops.length ? tr(ops, null, "command") : null;
  };
}

/** Pin a page break before a paragraph — the "Organize Pages" repair pass uses
 *  this on a group that merged onto the previous page after a reorder. No-op if
 *  the block isn't a paragraph or already breaks. */
export function pinPageBreakBeforeCmd(blockId: string): Command {
  return (state) => {
    const ops = pinPageBreakBeforeOps(state.doc, blockId);
    return ops.length ? tr(ops, state.selection, "command") : null;
  };
}
