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
  findSdtRanges,
  insertText,
  insertFragment,
  deleteBackward,
  deleteForward,
  deleteImage,
  insertTableRowCmd,
  replaceBackAndInsert,
  sdtAtPosition,
  setAlignment,
  setCharStyle as setCharStyleCmd,
  setImageProps,
  setParaProps,
  setSdtContent,
  setTableColFractionsCmd,
  splitParagraph,
  toggleCharStyle,
  toggleSdtCheckbox,
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

  const SDT_LABELS: Record<string, string> = {
    richText: "Rich Text",
    plainText: "Text",
    checkbox: "Check Box",
    dropDown: "Drop-Down List",
    comboBox: "Combo Box",
    date: "Date",
  };

  /** Word's active-control chrome: gray frame + title tab around the control
   *  containing the caret. Cleared when the caret leaves. */
  const updateSdtAdornment = (): void => {
    const focus = selection?.focus;
    const id = focus ? sdtAtPosition(doc, focus) : null;
    const props = id ? doc.sdts?.[id] : undefined;
    if (!id || !props) {
      paint.setSdtAdornment(null);
      return;
    }
    const rects = findSdtRanges(doc, id).flatMap((r) =>
      selectionRects(
        tree,
        { anchor: { blockId: r.blockId, offset: r.start }, focus: { blockId: r.blockId, offset: r.end } },
        scope(),
      ),
    );
    paint.setSdtAdornment({ rects, label: props.alias ?? SDT_LABELS[props.type] ?? "Content control" });
  };

  const refreshSelectionVisuals = (): void => {
    updateSdtAdornment();
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
    // Word: entering a placeholder control selects its whole content, so the
    // first keystroke replaces the prompt text.
    if (next && isCollapsed(next)) {
      const focus = next.focus;
      const id = sdtAtPosition(doc, focus);
      const props = id ? doc.sdts?.[id] : undefined;
      if (id && props?.placeholder) {
        const r = findSdtRanges(doc, id).find(
          (rr) => rr.blockId === focus.blockId && focus.offset >= rr.start && focus.offset <= rr.end,
        );
        if (r && r.end > r.start) {
          next = {
            anchor: { blockId: r.blockId, offset: r.start },
            focus: { blockId: r.blockId, offset: r.end },
          };
        }
      }
    }
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

  // ---- Content controls (SDT): popups, typing rules ------------------------

  /** The control + props under the caret (focus side). */
  const sdtAtCaret = (): { id: string; props: NonNullable<Document["sdts"]>[string] } | null => {
    const focus = selection?.focus;
    const id = focus ? sdtAtPosition(doc, focus) : null;
    const props = id ? doc.sdts?.[id] : undefined;
    return id && props ? { id, props } : null;
  };

  let sdtPopup: HTMLDivElement | null = null;
  const closeSdtPopup = (): void => {
    sdtPopup?.remove();
    sdtPopup = null;
  };

  /** Dropdown list / combo / date picker beside the control (Word's chooser). */
  const openSdtPopup = (id: string): void => {
    closeSdtPopup();
    const props = doc.sdts?.[id];
    const range = findSdtRanges(doc, id)[0];
    if (!props || !range) return;
    const rect = caretRect(tree, { blockId: range.blockId, offset: range.start }, scope());
    if (!rect) return;
    const at = paint.caretToContainer(rect);
    if (!at) return;
    const panel = document.createElement("div");
    panel.style.cssText =
      `position:absolute;left:${at.left}px;top:${at.top + rect.height + 4}px;z-index:30;` +
      "background:#fff;border:1px solid #c8c8c8;border-radius:6px;box-shadow:0 4px 12px rgba(0,0,0,.22);" +
      "font:13px Arial;min-width:160px;max-height:240px;overflow:auto;padding:4px;";
    panel.addEventListener("mousedown", (e) => {
      e.stopPropagation(); // keep the press away from the selection controller
    });
    if (props.type === "date") {
      const input = document.createElement("input");
      input.type = "date";
      input.style.cssText = "font:13px Arial;border:1px solid #c8c8c8;border-radius:4px;padding:3px 6px;";
      const pick = (): void => {
        if (!input.value) return;
        const [y, m, d] = input.value.split("-").map(Number);
        const fmt = props.dateFormat ?? "M/d/yyyy";
        const text = fmt
          .replace(/yyyy/g, String(y))
          .replace(/MM/g, String(m!).padStart(2, "0"))
          .replace(/M(?!M)/g, String(m))
          .replace(/dd/g, String(d!).padStart(2, "0"))
          .replace(/d(?!d)/g, String(d));
        dispatch(setSdtContent(id, text));
        closeSdtPopup();
        proxy.focus();
      };
      input.addEventListener("change", pick);
      panel.appendChild(input);
    } else {
      for (const item of props.listItems ?? []) {
        const row = document.createElement("div");
        row.textContent = item.display;
        row.style.cssText = "padding:4px 10px;border-radius:4px;cursor:pointer;";
        row.addEventListener("mouseenter", () => (row.style.background = "#e8eaed"));
        row.addEventListener("mouseleave", () => (row.style.background = ""));
        row.addEventListener("click", () => {
          dispatch(setSdtContent(id, item.display));
          closeSdtPopup();
          proxy.focus();
        });
        panel.appendChild(row);
      }
      if ((props.listItems ?? []).length === 0) {
        const empty = document.createElement("div");
        empty.textContent = "(no list items)";
        empty.style.cssText = "padding:4px 10px;color:#80868b;";
        panel.appendChild(empty);
      }
    }
    container.appendChild(panel);
    sdtPopup = panel;
  };

  const onSdtPress = (pos: DocPosition): boolean => {
    closeSdtPopup();
    const id = sdtAtPosition(doc, pos);
    const props = id ? doc.sdts?.[id] : undefined;
    if (!id || !props) return false;
    if (props.type === "checkbox") {
      dispatch(toggleSdtCheckbox(id));
      return true; // consume: the click IS the toggle
    }
    if (props.type === "dropDown" || props.type === "comboBox" || props.type === "date") {
      // Open after the controller places the caret (same frame ordering).
      requestAnimationFrame(() => openSdtPopup(id));
    }
    return false;
  };
  container.addEventListener("mousedown", () => closeSdtPopup());

  const insertWithAutoCorrect = (data: string): void => {
    const sdt = sdtAtCaret();
    if (sdt) {
      const { id, props } = sdt;
      // Not text-editable: locked content, pure dropdowns, checkboxes.
      if (props.lockContent || props.type === "dropDown" || props.type === "checkbox") return;
      if (props.placeholder) {
        // First keystroke replaces the gray prompt with real content.
        dispatch(setSdtContent(id, data));
        return;
      }
    }
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

  /** True when the caret sits in a control whose content must not be edited. */
  const sdtBlocksEdit = (): boolean => {
    const sdt = sdtAtCaret();
    return !!sdt && (sdt.props.lockContent === true || sdt.props.type === "dropDown");
  };

  const proxy = createImeProxy(container, {
    onInsertText: (text) => insertWithAutoCorrect(text),
    onDeleteBackward: () => {
      if (sdtBlocksEdit()) return;
      dispatch(deleteBackward());
    },
    onDeleteForward: () => {
      if (sdtBlocksEdit()) return;
      dispatch(deleteForward());
    },
    onSplitParagraph: () => {
      if (sdtAtCaret()) return; // controls are inline — no paragraph splits inside
      dispatch(splitParagraph());
    },
    onPaste: ({ html, text }) => {
      if (sdtBlocksEdit()) return;
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
    onSdtPress,
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
