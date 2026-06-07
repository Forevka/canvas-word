// Composition root: model -> layout -> paint -> input -> editor core -> a11y.
// One-way data flow: input -> command -> transaction -> applyOp* -> new state
// -> incremental layout -> paint + caret + proxy reposition (same frame).

import type { CharStyle, Document, ParaStyle, TableBlock } from "./model/document";
import { BAND_CONTAINERS } from "./model/document";
import type { DocPosition, DocSelection } from "./model/position";
import { isCollapsed } from "./model/position";
import { applyOp, effectiveFractions, type Op } from "./model/ops";
import { bandParagraphs, blockById, containerListOf, locateParagraph, paragraphsOf, styleAtRuns, textOfRuns } from "./model/text";
import { createLayoutEngine, type LayoutEngine } from "./layout/engine";
import {
  caretRect,
  objectRect,
  selectionRects,
  type ColumnBoundaryHit,
  type GeoScope,
} from "./layout/geometry";
import type { LayoutTree } from "./layout/layoutTree";
import { createPaintLayer } from "./paint/renderer";
import { createSelectionController } from "./input/selectionController";
import { createObjectFrame } from "./input/objectController";
import { createImeProxy } from "./input/imeProxy";
import { createKeymapHandler, type StyleKey } from "./input/keymap";
import { htmlToFragment } from "./input/clipboard";
import { createA11yMirror } from "./a11y/mirror";
import {
  changeListLevel,
  insertText,
  insertFragment,
  deleteBackward,
  deleteForward,
  deleteImage,
  insertTableRowCmd,
  replaceBackAndInsert,
  setAlignment,
  setCharStyle as setCharStyleCmd,
  setImageProps,
  setParaProps,
  setTableColFractionsCmd,
  splitParagraph,
  toggleCharStyle,
} from "./editor/commands";
import type { Command, EditorState, Transaction } from "./editor/state";
import { UndoManager } from "./editor/undo";

export interface CurrentFormat {
  styleId: string | null;
  fontFamily: string | null;
  fontSizePx: number | null;
  lineHeight: number | null;
}

export interface SearchState {
  index: number; // 1-based current match (0 = none)
  total: number;
}

export interface Editor {
  focus(): void;
  getDocument(): Document;
  getSelection(): DocSelection | null;
  getSelectedObject(): string | null;
  dispatch(cmd: Command): void;
  toggleStyle(key: StyleKey): void;
  /** Absolute char patch: range -> restyle runs; collapsed -> pending style. */
  setCharStyle(patch: Partial<CharStyle>): void;
  /** Alignment routes to the selected image when one is selected. */
  align(align: ParaStyle["align"]): void;
  /** Formatting at the caret — drives toolbar control state. */
  currentFormat(): CurrentFormat;
  /** Format painter: capture caret formatting, apply on the next selection. */
  armFormatPainter(sticky: boolean): void;
  cancelFormatPainter(): void;
  /** Find & replace. search() highlights all matches and returns state. */
  search(query: string, opts?: { matchCase?: boolean; wholeWord?: boolean }): SearchState;
  searchNav(dir: 1 | -1): SearchState;
  searchReplaceCurrent(replacement: string): SearchState;
  searchReplaceAll(replacement: string): number;
  searchClear(): void;
  undo(): void;
  redo(): void;
  destroy(): void;
}

export interface EditorOptions {
  /** Reuse a pre-warmed layout engine (its caches) instead of a fresh one. */
  engine?: LayoutEngine;
  /** Fires after any selection or document change (toolbar sync). */
  onChange?: () => void;
}

export function createEditor(
  container: HTMLElement,
  initialDoc: Document,
  options: EditorOptions = {},
): Editor {
  const engine = options.engine ?? createLayoutEngine();
  const paint = createPaintLayer(container);
  const undoMgr = new UndoManager();

  let doc = initialDoc;
  let tree: LayoutTree = engine.layout(doc);
  let selection: DocSelection | null = null;
  let pendingStyle: Partial<CharStyle> | null = null;
  let activeStory: GeoScope | null = null; // header/footer story-edit scope
  let savedBodySelection: DocSelection | null = null; // restored on story exit
  paint.setTree(tree);

  const state = (): EditorState => ({ doc, selection, pendingStyle });
  const scope = (): GeoScope | undefined => activeStory ?? undefined;

  // ---- selection visuals + proxy follow ----------------------------------

  const refreshSelectionVisuals = (): void => {
    if (!selection) {
      paint.setSelectionRects([]);
      paint.setCaret(null);
      return;
    }
    if (isCollapsed(selection)) {
      paint.setSelectionRects([]);
      const caret = caretRect(tree, selection.focus, scope());
      paint.setCaret(caret);
      if (caret) {
        const at = paint.caretToContainer(caret);
        if (at) proxy.moveTo(at.left, at.top, caret.height);
      }
    } else {
      paint.setSelectionRects(selectionRects(tree, selection, scope()));
      paint.setCaret(null); // Word hides the caret while a range is selected
    }
  };

  const notifyChange = (): void => {
    options.onChange?.();
  };

  const setSelection = (next: DocSelection | null): void => {
    selection = next;
    pendingStyle = null; // moving the caret drops the pending typing style
    refreshSelectionVisuals();
    mirror.sync(state());
    notifyChange();
  };

  // ---- story mode (header/footer band editing) ----------------------------

  const relayout = (): void => {
    tree = engine.layout(doc, undefined, { rawBand: activeStory?.band ?? null });
    paint.setTree(tree);
  };

  const setStory = (next: GeoScope | null): void => {
    const changingBand = (activeStory?.band ?? null) !== (next?.band ?? null);
    if (!changingBand && activeStory?.pageIndex === next?.pageIndex) return;
    if (next) selectObject(null); // objects and band stories are exclusive modes
    if (next && !activeStory) savedBodySelection = selection;
    activeStory = next;
    pendingStyle = null;
    paint.setBandEditMode(next?.band ?? null);
    if (changingBand) {
      relayout(); // the edited band switches between raw and substituted text
      selection = next ? null : savedBodySelection;
    }
    refreshSelectionVisuals();
    mirror.sync(state());
  };

  // ---- object selection (images): frame, resize, alignment, delete --------

  let selectedObject: string | null = null;
  let resizeBase: { w: number; h: number } | null = null;

  const contentWidth = (): number =>
    doc.section.pageWidthPx - doc.section.marginPx.left - doc.section.marginPx.right;

  const objectFrame = createObjectFrame({
    getPageElement: (i) => paint.getPageElement(i),
    onResizePreview: (w, h) => {
      if (!selectedObject) return;
      const img = doc.blocks.find((b) => b.id === selectedObject);
      if (img?.kind !== "image") return;
      resizeBase ??= { w: img.widthPx, h: img.heightPx };
      dispatch(setImageProps(selectedObject, { widthPx: w, heightPx: h }, "transient"));
    },
    onResizeCommit: (w, h) => {
      if (!selectedObject || !resizeBase) {
        resizeBase = null;
        return;
      }
      const base = resizeBase;
      resizeBase = null;
      // Revert the transient preview, then ONE undoable op for the whole drag.
      dispatch(setImageProps(selectedObject, { widthPx: base.w, heightPx: base.h }, "transient"));
      dispatch(setImageProps(selectedObject, { widthPx: w, heightPx: h }));
    },
  });

  const refreshObjectFrame = (): void => {
    if (!selectedObject) {
      objectFrame.hide();
      return;
    }
    const rect = objectRect(tree, selectedObject);
    if (!rect) {
      selectedObject = null;
      objectFrame.hide();
      return;
    }
    objectFrame.show(rect, contentWidth());
  };

  const selectObject = (blockId: string | null): void => {
    if (blockId === selectedObject) {
      if (blockId) refreshObjectFrame();
      return;
    }
    selectedObject = blockId;
    if (blockId) {
      selection = null; // object selection replaces the text selection (Word)
      refreshSelectionVisuals();
    }
    refreshObjectFrame();
  };

  // ---- table column-boundary drag ------------------------------------------

  const startColumnDrag = (hit: ColumnBoundaryHit, ev: MouseEvent): void => {
    const table = doc.blocks.find((b) => b.id === hit.tableId);
    if (table?.kind !== "table") return;
    const base = effectiveFractions(table);
    const minFrac = 24 / hit.tableWidth; // columns never shrink below 24px
    const startX = ev.clientX;
    let lastFractions = base;

    const fractionsFor = (e: MouseEvent): number[] => {
      const df = (e.clientX - startX) / hit.tableWidth;
      const f = base.slice();
      const a = hit.boundaryIndex;
      const b = a + 1;
      const shift = Math.max(-(f[a]! - minFrac), Math.min(f[b]! - minFrac, df));
      f[a] = f[a]! + shift;
      f[b] = f[b]! - shift;
      return f;
    };

    const onMove = (e: MouseEvent): void => {
      lastFractions = fractionsFor(e);
      dispatch(setTableColFractionsCmd(hit.tableId, lastFractions, "transient"));
    };
    const onUp = (e: MouseEvent): void => {
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      lastFractions = fractionsFor(e);
      // Revert preview, commit one undoable op.
      dispatch(setTableColFractionsCmd(hit.tableId, base, "transient"));
      dispatch(setTableColFractionsCmd(hit.tableId, lastFractions));
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  };

  // ---- Tab navigation between table cells ----------------------------------

  const tabInTable = (backward: boolean): boolean => {
    if (!selection) return false;
    // Lists win: Tab at the START of a list paragraph changes its level.
    if (isCollapsed(selection) && selection.focus.offset === 0) {
      const blk = doc.blocks.find((b) => b.id === selection!.focus.blockId);
      if (blk?.kind === "paragraph" && blk.style.list) {
        dispatch(changeListLevel(backward ? -1 : 1));
        return true;
      }
    }
    const loc = locateParagraph(doc, selection.focus.blockId);
    if (loc?.kind !== "cell") return false;
    const table = containerListOf(doc, loc.where)[loc.bi] as TableBlock;
    const flat: { ri: number; ci: number }[] = [];
    table.rows.forEach((row, ri) => row.cells.forEach((_, ci) => flat.push({ ri, ci })));
    const pos = flat.findIndex((c) => c.ri === loc.ri && c.ci === loc.ci);
    const target = pos + (backward ? -1 : 1);
    if (target < 0) return true; // Shift+Tab in the first cell: stay (Word)
    if (target >= flat.length) {
      // Tab in the last cell appends a row and moves into it (Word behavior).
      dispatch(insertTableRowCmd("below"));
      return true;
    }
    const cell = table.rows[flat[target]!.ri]!.cells[flat[target]!.ci]!;
    const paras = cell.blocks.filter((b): b is import("./model/document").Paragraph => b.kind === "paragraph");
    const first = paras[0];
    const last = paras[paras.length - 1];
    if (!first || !last) return true; // image-only cell: consume Tab, nowhere to caret
    // Word selects the target cell's content.
    setSelection({
      anchor: { blockId: first.id, offset: 0 },
      focus: { blockId: last.id, offset: textOfRuns(last.runs).length },
    });
    return true;
  };

  const toggleStyle = (key: StyleKey): void => {
    if (selection && !isCollapsed(selection)) {
      dispatch(toggleCharStyle(key));
      return;
    }
    if (!selection) return;
    const block = blockById(doc, selection.focus.blockId);
    const inherited = block ? styleAtRuns(block.runs, selection.focus.offset) : undefined;
    const effective = { ...(inherited ?? {}), ...(pendingStyle ?? {}) } as Partial<CharStyle>;
    pendingStyle = { ...(pendingStyle ?? {}), [key]: !effective[key] };
    mirror.announce(`${key} ${pendingStyle[key] ? "on" : "off"}`);
  };

  // ---- transaction pipeline ------------------------------------------------

  const runOps = (ops: Op[]): Op[] => {
    const inverses: Op[] = [];
    for (const op of ops) {
      const res = applyOp(doc, op);
      doc = res.doc;
      inverses.unshift(res.inverse);
    }
    return inverses;
  };

  const afterMutation = (selectionAfter: DocSelection | null): void => {
    relayout();
    selection = selectionAfter;
    refreshSelectionVisuals();
    refreshObjectFrame(); // images move/resize with reflow; frame follows
    if (searchQuery) {
      runSearch(); // live re-search while the find bar is open
      paintSearch();
    }
    if (selection && isCollapsed(selection)) {
      const caret = caretRect(tree, selection.focus, scope());
      if (caret) paint.ensureVisible(caret);
    }
    mirror.sync(state());
    notifyChange();
  };

  const commit = (trn: Transaction): void => {
    const selectionBefore = selection;
    const inverses = runOps(trn.ops);
    if (trn.origin !== "transient") {
      undoMgr.record({
        ops: trn.ops,
        inverseOps: inverses,
        selectionBefore,
        selectionAfter: trn.selectionAfter,
        origin: trn.origin,
        time: Date.now(),
      });
    }
    afterMutation(trn.selectionAfter);
  };

  const dispatch = (cmd: Command): void => {
    const trn = cmd(state());
    if (trn) commit(trn);
  };

  const undo = (): void => {
    const entry = undoMgr.popUndo();
    if (!entry) return;
    runOps(entry.inverseOps);
    afterMutation(entry.selectionBefore);
  };

  const redo = (): void => {
    const entry = undoMgr.popRedo();
    if (!entry) return;
    runOps(entry.ops);
    afterMutation(entry.selectionAfter);
  };

  // ---- AutoCorrect (typographic) -------------------------------------------

  const autoCorrect = { quotes: true, dashes: true, symbols: true };

  const insertWithAutoCorrect = (data: string): void => {
    if (data.length === 1 && selection && isCollapsed(selection)) {
      const block = blockById(doc, selection.focus.blockId);
      const prev = block ? textOfRuns(block.runs).slice(0, selection.focus.offset) : "";
      if (autoCorrect.quotes && (data === '"' || data === "'")) {
        const before = prev.slice(-1);
        const open = before === "" || /[\s([{‘“—-]/.test(before);
        dispatch(insertText(data === '"' ? (open ? "“" : "”") : open ? "‘" : "’"));
        return;
      }
      if (autoCorrect.dashes && data === "-" && prev.endsWith("-") && !prev.endsWith("--")) {
        dispatch(replaceBackAndInsert(1, "—")); // -- → em dash
        return;
      }
      if (autoCorrect.symbols && data === ")") {
        if (/\(c$/i.test(prev)) return dispatch(replaceBackAndInsert(2, "©"));
        if (/\(r$/i.test(prev)) return dispatch(replaceBackAndInsert(2, "®"));
        if (/\(tm$/i.test(prev)) return dispatch(replaceBackAndInsert(3, "™"));
      }
    }
    dispatch(insertText(data));
  };

  // ---- Format painter --------------------------------------------------------

  let painter: { char: CharStyle; para: Partial<ParaStyle>; sticky: boolean } | null = null;
  let painterMouseUp: (() => void) | null = null;

  const cancelFormatPainter = (): void => {
    painter = null;
    delete container.dataset["painter"];
    container.style.cursor = "text";
    if (painterMouseUp) {
      window.removeEventListener("mouseup", painterMouseUp);
      painterMouseUp = null;
    }
  };

  const armFormatPainter = (sticky: boolean): void => {
    if (painter) {
      cancelFormatPainter(); // pressing the button again disarms
      return;
    }
    if (!selection) return;
    const block = blockById(doc, selection.focus.blockId);
    const char = block ? styleAtRuns(block.runs, selection.focus.offset) : undefined;
    if (!block || !char) return;
    const { namedStyle: _ns, list: _list, ...para } = block.style;
    painter = { char: { ...char }, para, sticky };
    container.dataset["painter"] = "1";
    container.style.cursor = "copy";
    // Arming runs in the toolbar button's `click` (after its mouseup), so the
    // next mouseup IS the user's apply gesture — no skip needed.
    painterMouseUp = (): void => {
      setTimeout(() => {
        if (!painter || !selection) return;
        // range gesture → char + para; bare click → para only (Word)
        if (!isCollapsed(selection)) dispatch(setCharStyleCmd(painter.char));
        dispatch(setParaProps(painter.para));
        if (!painter.sticky) cancelFormatPainter();
      }, 0);
    };
    window.addEventListener("mouseup", painterMouseUp);
  };

  // ---- Find & replace ---------------------------------------------------------

  interface Match {
    blockId: string;
    start: number;
    end: number;
  }
  let searchQuery: { q: string; matchCase: boolean; wholeWord: boolean } | null = null;
  let matches: Match[] = [];
  let matchIndex = -1;

  const isWordChar = (c: string | undefined): boolean => c !== undefined && /[\p{L}\p{N}]/u.test(c);

  const runSearch = (): void => {
    matches = [];
    if (!searchQuery || searchQuery.q.length === 0) return;
    const { q, matchCase, wholeWord } = searchQuery;
    const needle = matchCase ? q : q.toLowerCase();
    // Body + cell paragraphs; band stories are excluded (their selection needs
    // story scope — find lands in the body, like Word's default scope).
    const bandIds = new Set(
      BAND_CONTAINERS.flatMap((band) => bandParagraphs(doc, band)).map((p) => p.id),
    );
    for (const p of paragraphsOf(doc)) {
      if (bandIds.has(p.id)) continue;
      const text = textOfRuns(p.runs);
      const hay = matchCase ? text : text.toLowerCase();
      let from = 0;
      for (;;) {
        const i = hay.indexOf(needle, from);
        if (i < 0) break;
        const ok = !wholeWord || (!isWordChar(text[i - 1]) && !isWordChar(text[i + needle.length]));
        if (ok) matches.push({ blockId: p.id, start: i, end: i + needle.length });
        from = i + Math.max(1, needle.length);
      }
    }
  };

  const paintSearch = (): void => {
    const rects = matches.flatMap((m) =>
      selectionRects(tree, {
        anchor: { blockId: m.blockId, offset: m.start },
        focus: { blockId: m.blockId, offset: m.end },
      }),
    );
    paint.setSearchRects(rects);
  };

  const searchState = (): SearchState => ({ index: matchIndex + 1, total: matches.length });

  const gotoMatch = (i: number): void => {
    if (matches.length === 0) {
      matchIndex = -1;
      return;
    }
    matchIndex = ((i % matches.length) + matches.length) % matches.length;
    const m = matches[matchIndex]!;
    setSelection({
      anchor: { blockId: m.blockId, offset: m.start },
      focus: { blockId: m.blockId, offset: m.end },
    });
    const rect = caretRect(tree, { blockId: m.blockId, offset: m.start });
    if (rect) paint.ensureVisible(rect);
  };

  const search: Editor["search"] = (query, opts = {}) => {
    searchQuery = { q: query, matchCase: opts.matchCase ?? false, wholeWord: opts.wholeWord ?? false };
    runSearch();
    matchIndex = matches.length > 0 ? -1 : -1;
    paintSearch();
    if (matches.length > 0) gotoMatch(0);
    return searchState();
  };

  const searchNav: Editor["searchNav"] = (dir) => {
    if (matches.length > 0) gotoMatch(matchIndex + dir);
    return searchState();
  };

  const replaceMatchTr = (m: Match, replacement: string): Command => (state) => ({
    ops: [
      { type: "deleteRange", blockId: m.blockId, start: m.start, end: m.end },
      ...(replacement.length > 0
        ? ([{ type: "insertText", at: { blockId: m.blockId, offset: m.start }, text: replacement }] as Op[])
        : []),
    ],
    selectionAfter: {
      anchor: { blockId: m.blockId, offset: m.start },
      focus: { blockId: m.blockId, offset: m.start + replacement.length },
    },
    origin: "command",
  });
  const searchReplaceCurrent: Editor["searchReplaceCurrent"] = (replacement) => {
    if (matchIndex < 0 || !matches[matchIndex]) return searchState();
    dispatch(replaceMatchTr(matches[matchIndex]!, replacement));
    // afterMutation re-ran the search; land on the match now nearest that spot
    if (matches.length > 0) gotoMatch(Math.min(matchIndex, matches.length - 1));
    return searchState();
  };

  const searchReplaceAll: Editor["searchReplaceAll"] = (replacement) => {
    if (matches.length === 0) return 0;
    const count = matches.length;
    // Back-to-front so earlier offsets stay valid — ONE transaction, one undo.
    const all = [...matches].reverse();
    const cmd: Command = () => ({
      ops: all.flatMap((m): Op[] => [
        { type: "deleteRange", blockId: m.blockId, start: m.start, end: m.end },
        ...(replacement.length > 0
          ? ([{ type: "insertText", at: { blockId: m.blockId, offset: m.start }, text: replacement }] as Op[])
          : []),
      ]),
      selectionAfter: selection,
      origin: "command",
    });
    dispatch(cmd);
    return count;
  };

  const searchClear = (): void => {
    searchQuery = null;
    matches = [];
    matchIndex = -1;
    paint.setSearchRects([]);
  };

  // ---- IME composition: transient preview edits outside the undo stack -----

  let transient: { at: DocPosition; len: number } | null = null;

  const compositionStyle = (at: DocPosition): CharStyle | undefined => {
    const block = blockById(doc, at.blockId);
    const inherited = block ? styleAtRuns(block.runs, at.offset) : undefined;
    return inherited ? { ...inherited, underline: true } : undefined;
  };

  const onCompositionStart = (): void => {
    // A range selection is consumed by the composition (undoable, like typing over it).
    if (selection && !isCollapsed(selection)) dispatch(deleteBackward());
    transient = selection ? { at: selection.focus, len: 0 } : null;
  };

  const onCompositionUpdate = (data: string): void => {
    if (!transient) return;
    const ops: Op[] = [];
    if (transient.len > 0) {
      ops.push({
        type: "deleteRange",
        blockId: transient.at.blockId,
        start: transient.at.offset,
        end: transient.at.offset + transient.len,
      });
    }
    const style = compositionStyle(transient.at);
    if (data.length > 0) {
      const op: Op = { type: "insertText", at: transient.at, text: data };
      if (style) op.style = style;
      ops.push(op);
    }
    transient.len = data.length;
    const caretPos = {
      blockId: transient.at.blockId,
      offset: transient.at.offset + data.length,
    };
    commit({ ops, selectionAfter: { anchor: caretPos, focus: caretPos }, origin: "transient" });
  };

  const onCompositionEnd = (data: string): void => {
    if (!transient) return;
    // Remove the transient preview, then commit the final text as ONE undoable insert.
    if (transient.len > 0) {
      commit({
        ops: [
          {
            type: "deleteRange",
            blockId: transient.at.blockId,
            start: transient.at.offset,
            end: transient.at.offset + transient.len,
          },
        ],
        selectionAfter: { anchor: transient.at, focus: transient.at },
        origin: "transient",
      });
    }
    transient = null;
    if (data.length > 0) dispatch(insertText(data));
  };

  // ---- input layer wiring ---------------------------------------------------

  const proxy = createImeProxy(container, {
    onInsertText: (text) => insertWithAutoCorrect(text),
    onDeleteBackward: () => dispatch(deleteBackward()),
    onDeleteForward: () => dispatch(deleteForward()),
    onSplitParagraph: () => dispatch(splitParagraph()),
    onPaste: ({ html, text }) => {
      if (html) {
        const fragment = htmlToFragment(html);
        if (fragment) {
          dispatch(insertFragment(fragment));
          return;
        }
      }
      if (!text) return;
      // Plain text: insert through insertText so it inherits the caret style.
      const parts = text.replace(/\r\n?/g, "\n").split("\n");
      dispatch(insertText(parts[0] ?? "", "paste"));
      for (let i = 1; i < parts.length; i++) {
        dispatch(splitParagraph());
        if (parts[i]!.length > 0) dispatch(insertText(parts[i]!, "paste"));
      }
    },
    onCompositionStart,
    onCompositionUpdate,
    onCompositionEnd,
  });

  const mirror = createA11yMirror(proxy.el);

  const controller = createSelectionController({
    container,
    getTree: () => tree,
    getDoc: () => doc,
    getSelection: () => selection,
    setSelection,
    clientToPage: (x, y) => paint.clientToPage(x, y),
    focusProxy: () => proxy.focus(),
    onDeleteSelection: () => dispatch(deleteBackward()),
    getStory: () => activeStory,
    setStory,
    selectObject,
    hasSelectedObject: () => selectedObject !== null,
    deleteSelectedObject: () => {
      if (!selectedObject) return;
      const id = selectedObject;
      selectObject(null);
      dispatch(deleteImage(id));
    },
    startColumnDrag,
    onTab: tabInTable,
    jumpToBlock: (blockId: string): void => {
      // Word: Ctrl+click on a TOC entry moves the caret to the heading.
      if (!blockById(doc, blockId)) return;
      setSelection({ anchor: { blockId, offset: 0 }, focus: { blockId, offset: 0 } });
      const rect = caretRect(tree, { blockId, offset: 0 });
      if (rect) paint.ensureVisible(rect);
    },
  });

  const keymapHandler = createKeymapHandler({ dispatch, undo, redo, toggleStyle });
  container.addEventListener("keydown", keymapHandler);

  return {
    focus(): void {
      proxy.focus();
    },
    getDocument(): Document {
      return doc;
    },
    getSelection(): DocSelection | null {
      return selection;
    },
    getSelectedObject(): string | null {
      return selectedObject;
    },
    dispatch,
    toggleStyle,
    setCharStyle(patch: Partial<CharStyle>): void {
      if (selection && !isCollapsed(selection)) {
        dispatch(setCharStyleCmd(patch));
        return;
      }
      if (!selection) return;
      pendingStyle = { ...(pendingStyle ?? {}), ...patch }; // applies to next typed text
      mirror.announce("formatting set for next text");
    },
    currentFormat(): CurrentFormat {
      const focus = selection?.focus;
      const block = focus ? blockById(doc, focus.blockId) : undefined;
      const char = block && focus ? styleAtRuns(block.runs, focus.offset) : undefined;
      const effective = { ...(char ?? {}), ...(pendingStyle ?? {}) };
      return {
        styleId: block?.style.namedStyle ?? (block ? "Normal" : null),
        fontFamily: effective.fontFamily ?? null,
        fontSizePx: effective.fontSizePx ?? null,
        lineHeight: block?.style.lineHeight ?? null,
      };
    },
    align(align: ParaStyle["align"]): void {
      if (selectedObject) {
        if (align !== "justify") dispatch(setImageProps(selectedObject, { align }));
        return;
      }
      dispatch(setAlignment(align));
    },
    armFormatPainter,
    cancelFormatPainter,
    search,
    searchNav,
    searchReplaceCurrent,
    searchReplaceAll,
    searchClear,
    undo,
    redo,
    destroy(): void {
      cancelFormatPainter();
      searchClear();
      container.removeEventListener("keydown", keymapHandler);
      controller.destroy();
      objectFrame.destroy();
      mirror.destroy();
      proxy.destroy();
      paint.destroy();
    },
  };
}
