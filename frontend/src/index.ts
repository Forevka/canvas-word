// Composition root: model -> layout -> paint -> input -> editor core -> a11y.
// One-way data flow: input -> command -> transaction -> applyOp* -> new state
// -> incremental layout -> paint + caret + proxy reposition (same frame).

import type { Block, CharStyle, Document, ParaStyle, TableBlock } from "@cw/shared";
import { BAND_CONTAINERS } from "@cw/shared";
import type { DocPosition, DocSelection } from "@cw/shared";
import { isCollapsed } from "@cw/shared";
import { applyOp, containerBlocks, containerOf, effectiveFractions, locateImage, sliceRuns, type Op } from "@cw/shared";
import { bandParagraphs, blockById, containerListOf, locateParagraph, paragraphsOf, styleAtRuns, textOfRuns } from "@cw/shared";
import { createLayoutEngine, type LayoutEngine } from "./layout/engine";
import {
  caretRect,
  comparePositions,
  hitTest,
  hitTestObject,
  linkAt,
  objectRect,
  selectionRects,
  type ColumnBoundaryHit,
  type GeoScope,
  type Rect,
} from "./layout/geometry";
import type { LayoutTree, Page, PlacedBlock } from "./layout/layoutTree";
import { createPaintLayer } from "./paint/renderer";
import { createSelectionController } from "./input/selectionController";
import { createObjectFrame } from "./input/objectController";
import { createImeProxy } from "./input/imeProxy";
import { createKeymapHandler, type StyleKey } from "./input/keymap";
import { extractFragment, fragmentToHtml, fragmentToPlainText, htmlToFragment, type DocFragment } from "./input/clipboard";
import { showContextMenu, type ContextMenuHandle, type MenuEntry } from "./ui/contextMenu";
import { showSdtInspector, type SdtInspectorData, type SdtInspectorHandle } from "./ui/sdtInspector";
import { createA11yMirror } from "./a11y/mirror";
import {
  changeListLevel,
  deleteTableRowCmd,
  deleteTableColumnCmd,
  deleteTableCmd,
  findSdtRanges,
  insertContentControl,
  insertText,
  insertFragment,
  deleteBackward,
  deleteForward,
  deleteImage,
  insertTableColumnCmd,
  insertTableRowCmd,
  mergeCellsCmd,
  removeContentControl,
  replaceBackAndInsert,
  replaceSdtContent,
  replaceSdtBlockSpan,
  replaceSdtCellContent,
  sdtAtPosition,
  setAlignment,
  setCharStyle as setCharStyleCmd,
  setImageProps,
  setLinkCmd,
  setParaProps,
  setSdtContent,
  setTableColFractionsCmd,
  splitParagraph,
  toggleCharStyle,
  toggleList,
  toggleSdtCheckbox,
  unmergeCellCmd,
} from "./editor/commands";
import type { SdtType } from "@cw/shared";
import { ICONS } from "./ui/icons";
import type { Command, EditorState, Transaction } from "./editor/state";
import { UndoManager } from "./editor/undo";
import { ChangeRecorder, type ChangeSink } from "./sync/changeRecorder";
import type { Change, ChangeOrigin } from "@cw/shared";

export interface CurrentFormat {
  styleId: string | null;
  fontFamily: string | null;
  fontSizePx: number | null;
  lineHeight: number | null;
  /** Character toggles at the caret (incl. the pending style) — drive the
   *  ribbon's pressed-button state. */
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  highlight: boolean;
  superscript: boolean;
  subscript: boolean;
  /** Paragraph alignment, and which list (if any) the caret paragraph is in. */
  align: ParaStyle["align"] | null;
  listKind: "bullet" | "number" | null;
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
  /** Open the content-control inspector for the control at the caret (ribbon
   *  button). Returns false when the caret isn't inside a content control. */
  inspectContentControl(): boolean;
  /** Update every TOC entry's page number to its target's current page (the
   *  imported pre-calculated numbers are shown until the user asks for this).
   *  Returns the count of entries whose number changed. */
  recalculateToc(): number;
  /** Presentational zoom (1 = 100%, clamped to [.25, 5]). No relayout. */
  setZoom(zoom: number): void;
  getZoom(): number;
  /** Format painter: capture caret formatting, apply on the next selection. */
  armFormatPainter(sticky: boolean): void;
  cancelFormatPainter(): void;
  /** Clipboard, mirroring the Ctrl+C/X/V handlers for ribbon buttons. */
  copy(): void;
  cut(): void;
  paste(): void;
  /** Move the caret to a block's start and scroll it into view (outline pane,
   *  navigation). No-op if the block id isn't in the current document. */
  revealBlock(blockId: string): void;
  /** Select the whole document body (Select All button / Ctrl+A). */
  selectAll(): void;
  /** Layout summary for the status bar: total pages and the caret's 1-based page. */
  getLayoutInfo(): { pageCount: number; currentPage: number };
  /** Viewport rect of the selected image (anchor for a floating toolbar), or null. */
  getSelectedObjectRect(): { left: number; top: number; width: number; height: number } | null;
  /** Delete the selected image and clear the object selection. */
  deleteSelectedObject(): void;
  /** Find & replace. search() highlights all matches and returns state. */
  search(query: string, opts?: { matchCase?: boolean; wholeWord?: boolean }): SearchState;
  searchNav(dir: 1 | -1): SearchState;
  searchReplaceCurrent(replacement: string): SearchState;
  searchReplaceAll(replacement: string): number;
  searchClear(): void;
  undo(): void;
  redo(): void;
  /** The document-history change log recorded this session (ordered). The base
   *  snapshot taken at load + this log reconstructs the current document. */
  getChangeLog(): Change[];
  /** The local version: number of changes recorded since load. */
  getChangeHead(): number;
  /** Apply ops received from a remote collaborator: mutates the document and
   *  rebases the local caret, without recording to the change log or undo stack. */
  applyRemoteOps(ops: Op[]): void;
  destroy(): void;
}

export interface EditorOptions {
  /** Reuse a pre-warmed layout engine (its caches) instead of a fresh one. */
  engine?: LayoutEngine;
  /** Fires after any selection or document change (toolbar sync). */
  onChange?: () => void;
  /** Fires after the zoom changes (toolbar/wheel), for a zoom indicator. */
  onZoomChange?: (zoom: number) => void;
  /** Document id this editor session is editing — stamped onto every recorded
   *  Change. Defaults to "local". */
  docId?: string;
  /** Fires for each committed Change (the document-history log entry). The
   *  SyncClient subscribes here to ship edits to the server. */
  onChangeRecorded?: ChangeSink;
}

export function createEditor(
  container: HTMLElement,
  initialDoc: Document,
  options: EditorOptions = {},
): Editor {
  const engine = options.engine ?? createLayoutEngine();
  const paint = createPaintLayer(container);
  const undoMgr = new UndoManager();
  // Document history: every committed edit (and undo/redo, as forward ops) is
  // recorded as a Change. The base snapshot + this ordered log reconstructs any
  // version (see shared/replay).
  const recorder = new ChangeRecorder(options.docId ?? "local", options.onChangeRecorded);

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

  /** Vertical extent (top, bottom) a placed block occupies on its page —
   *  paragraphs measure their line stack, images/tables their box. */
  const placedBlockVBounds = (pb: PlacedBlock): [number, number] => {
    if (pb.table) return [pb.table.y, pb.table.y + pb.table.height];
    if (pb.image) return [pb.y, pb.y + pb.image.height];
    let top = pb.y;
    let bot = pb.y;
    for (const l of pb.lines) {
      const ly = pb.y + l.y;
      if (ly < top) top = ly;
      if (ly + l.height > bot) bot = ly + l.height;
    }
    return [top, bot];
  };

  /** Body block-level controls (those wrapping WHOLE paragraphs/tables, like
   *  Word's boundingBox appearance) draw ONE frame per page spanning the full
   *  content column from the first block's top to the last block's bottom — not
   *  ragged per-line rects. Returns null for inline / cell-hosted controls, which
   *  keep the text-shaped highlight.
   *
   *  The importer's block-level tag (tagBlockSdt) covers EVERY paragraph it
   *  touches in full — including all cells of a contained table — whereas an
   *  inline control leaves partial coverage. So "every range is whole-paragraph"
   *  means block-level; a cell paragraph's range maps back to its top-level table
   *  via locateParagraph, letting an SDT that wraps a heading + table be framed. */
  const blockLevelSdtRects = (id: string): Rect[] | null => {
    const ranges = findSdtRanges(doc, id);
    if (ranges.length === 0) return null;
    const paraLen = new Map<string, number>();
    for (const p of paragraphsOf(doc)) paraLen.set(p.id, textOfRuns(p.runs).length);
    let minIdx = Infinity;
    let maxIdx = -Infinity;
    for (const r of ranges) {
      const len = paraLen.get(r.blockId);
      if (len === undefined || r.start > 0 || r.end < len) return null; // partial ⇒ inline
      const loc = locateParagraph(doc, r.blockId);
      // Body only: a top-level paragraph, or a cell paragraph whose table is the
      // top-level block (loc.bi). Bands/footnotes can't hold the caret here.
      if (!loc || (loc.kind !== "top" && !(loc.kind === "cell" && loc.where === "body"))) return null;
      minIdx = Math.min(minIdx, loc.bi);
      maxIdx = Math.max(maxIdx, loc.bi);
    }
    if (!Number.isFinite(minIdx)) return null;
    const ids = new Set(containerBlocks(doc, "body").slice(minIdx, maxIdx + 1).map((b) => b.id));
    // Union the span's blocks' vertical extent per page (a span can break pages).
    const byPage = new Map<number, { top: number; bot: number; page: Page }>();
    for (const page of tree.pages) {
      for (const pb of page.blocks) {
        if (!ids.has(pb.blockId)) continue;
        const [t, b] = placedBlockVBounds(pb);
        const cur = byPage.get(page.index);
        if (cur) {
          cur.top = Math.min(cur.top, t);
          cur.bot = Math.max(cur.bot, b);
        } else byPage.set(page.index, { top: t, bot: b, page });
      }
    }
    if (byPage.size === 0) return null;
    const rects: Rect[] = [];
    for (const { top, bot, page } of byPage.values()) {
      rects.push({
        pageIndex: page.index,
        x: page.marginPx.left,
        y: top,
        width: page.widthPx - page.marginPx.left - page.marginPx.right,
        height: bot - top,
      });
    }
    return rects;
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
    // Block-level controls get a single bounding box (Word's boundingBox chrome);
    // inline ones fall back to the text-shaped per-line rects.
    const rects =
      blockLevelSdtRects(id) ??
      findSdtRanges(doc, id).flatMap((r) =>
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
    getZoom: () => paint.getZoom(),
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
    notifyChange(); // object selection drives the floating image toolbar
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
    const paras = cell.blocks.filter((b): b is import("@cw/shared").Paragraph => b.kind === "paragraph");
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
    notifyChange(); // pending toggle drives the ribbon's pressed state
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
      // Mirror the edit into the document-history log (transient IME previews are
      // excluded, exactly as for undo). origin is now ChangeOrigin-compatible.
      recorder.record(trn.ops, trn.origin as ChangeOrigin, trn.selectionAfter, Date.now());
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
    // Undo is a real forward edit in history terms — record the inverse ops it
    // applied so the log faithfully replays the same end state.
    recorder.record(entry.inverseOps, "undo", entry.selectionBefore, Date.now());
    afterMutation(entry.selectionBefore);
  };

  const redo = (): void => {
    const entry = undoMgr.popRedo();
    if (!entry) return;
    runOps(entry.ops);
    recorder.record(entry.ops, "redo", entry.selectionAfter, Date.now());
    afterMutation(entry.selectionAfter);
  };

  // Apply ops that arrived from another collaborator. Unlike commit(), these are
  // NOT recorded (the server's log already holds them) and do NOT enter the undo
  // stack (a user can't undo a peer's edit). The local caret is rebased through
  // each op's position mapper so it survives the remote insert/delete.
  const applyRemoteOps = (ops: Op[]): void => {
    let sel = selection;
    for (const op of ops) {
      const res = applyOp(doc, op);
      doc = res.doc;
      if (sel) {
        sel = {
          anchor: res.mapPosition(sel.anchor),
          focus: res.mapPosition(sel.focus),
          ...(sel.goalX !== undefined ? { goalX: sel.goalX } : {}),
        };
      }
    }
    afterMutation(sel);
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
    onAnchorJump: (anchorName: string, fromBlockId: string | null): void => {
      const target = resolveAnchorTarget(anchorName, fromBlockId);
      if (!target) return;
      setSelection({ anchor: { blockId: target, offset: 0 }, focus: { blockId: target, offset: 0 } });
      const rect = caretRect(tree, { blockId: target, offset: 0 });
      if (rect) paint.ensureVisible(rect);
    },
  });

  const keymapHandler = createKeymapHandler({ dispatch, undo, redo, toggleStyle });
  container.addEventListener("keydown", keymapHandler);

  // ---- clipboard (context-menu Cut/Copy/Paste) -----------------------------
  // Copy/cut serialize the model fragment to the async Clipboard API; paste
  // reads it back. The keyboard path still flows through the native copy/cut
  // events on the controller — these are the menu's gesture-driven equivalents.

  const orderedSelection = (): [DocPosition, DocPosition] | null => {
    if (!selection || isCollapsed(selection)) return null;
    const cmp = comparePositions(tree, selection.anchor, selection.focus, scope());
    return cmp <= 0 ? [selection.anchor, selection.focus] : [selection.focus, selection.anchor];
  };

  const copySelection = async (): Promise<void> => {
    const o = orderedSelection();
    if (!o) return;
    const fragment = extractFragment(paragraphsOf(doc), o[0], o[1]);
    const html = fragmentToHtml(fragment);
    const text = fragmentToPlainText(fragment);
    try {
      await navigator.clipboard.write([
        new ClipboardItem({
          "text/html": new Blob([html], { type: "text/html" }),
          "text/plain": new Blob([text], { type: "text/plain" }),
        }),
      ]);
    } catch {
      try {
        await navigator.clipboard.writeText(text);
      } catch {
        /* clipboard unavailable (no permission / insecure context) */
      }
    }
  };

  const pasteText = (text: string): void => {
    const parts = text.replace(/\r\n?/g, "\n").split("\n");
    dispatch(insertText(parts[0] ?? "", "paste"));
    for (let i = 1; i < parts.length; i++) {
      dispatch(splitParagraph());
      if (parts[i]!.length > 0) dispatch(insertText(parts[i]!, "paste"));
    }
  };

  const pasteFromClipboard = async (): Promise<void> => {
    if (sdtBlocksEdit()) return;
    try {
      const items = await navigator.clipboard.read();
      for (const it of items) {
        if (it.types.includes("text/html")) {
          const html = await (await it.getType("text/html")).text();
          const frag = htmlToFragment(html);
          if (frag) {
            dispatch(insertFragment(frag));
            return;
          }
        }
      }
      for (const it of items) {
        if (it.types.includes("text/plain")) {
          pasteText(await (await it.getType("text/plain")).text());
          return;
        }
      }
    } catch {
      try {
        pasteText(await navigator.clipboard.readText());
      } catch {
        /* clipboard read blocked */
      }
    }
  };

  // ---- contextual right-click menu -----------------------------------------
  // Composition matrix (sections concatenated, only present targets emitted):
  //   always              → Clipboard (Cut/Copy/Paste)
  //   hyperlink           → Link (Open/Edit/Remove)
  //   image object        → Image (Wrap/Align/Delete)
  //   content control     → Control (Toggle|Choose/Remove)
  //   editable text       → Font, Paragraph, Insert
  //   list paragraph      → List (level up/down, remove)
  //   table cell          → Table (Insert/Delete/Merge/Unmerge)
  //   header/footer band  → Band (Edit / Close)
  let contextMenu: ContextMenuHandle | null = null;
  const closeContextMenu = (): void => {
    contextMenu?.close();
    contextMenu = null;
  };

  // Content-control inspector: gather an SDT's content and show its properties
  // + a faithful preview. Block-level controls render their whole block range
  // (paragraphs, images, tables, blank lines) so the preview mirrors the page;
  // inline controls render just their run slice.
  let sdtInspector: SdtInspectorHandle | null = null;

  /** Render a block list to preview HTML that mirrors the document — runs with
   *  their styles, images as <img>, tables as <table>, empty paragraphs kept. */
  // When `objRefs` is supplied (top-level editable render), each image/table is
  // captured into it and emitted wrapped in a contenteditable=false div carrying
  // its ref index — so the Save round-trip restores the object verbatim instead
  // of dropping it. Nested (cell) renders pass no collector: their objects ride
  // along inside the parent table's single ref.
  const renderBlocksHtml = (
    blocks: Block[],
    objRefs?: Block[],
  ): { html: string; text: string; paraCount: number; charCount: number; hasObjects: boolean } => {
    let html = "";
    let text = "";
    let paraCount = 0;
    let charCount = 0;
    let hasObjects = false;
    const objHtml = (inner: string): string => {
      if (!objRefs) return inner;
      const k = objRefs.length;
      // The referenced block is pushed by the caller; here we only know its slot.
      return `<div data-cw-ref="${k}" contenteditable="false" style="position:relative;">${inner}</div>`;
    };
    for (const b of blocks) {
      if (b.kind === "paragraph") {
        paraCount++;
        const t = textOfRuns(b.runs);
        charCount += t.length;
        text += t + "\n";
        html += t.length === 0
          ? "<p>&nbsp;</p>"
          : fragmentToHtml({ blocks: [{ runs: b.runs, style: b.style }], inline: false });
      } else if (b.kind === "image") {
        hasObjects = true;
        const src = b.src.replace(/"/g, "&quot;");
        const img = `<img src="${src}" style="display:block;margin:6px auto;max-width:100%;width:${Math.round(b.widthPx)}px;height:auto;">`;
        html += objHtml(img);
        if (objRefs) objRefs.push(b);
        text += "[image]\n";
      } else {
        hasObjects = true;
        let tbl = `<table style="border-collapse:collapse;width:100%;margin:6px 0;">`;
        for (const row of b.rows) {
          tbl += "<tr>";
          for (const cell of row.cells) {
            const span = cell.colSpan && cell.colSpan > 1 ? ` colspan="${cell.colSpan}"` : "";
            tbl += `<td${span} style="border:1px solid #d0d4d9;padding:3px 6px;vertical-align:top;">${renderBlocksHtml(cell.blocks).html}</td>`;
          }
          tbl += "</tr>";
        }
        tbl += "</table>";
        html += objHtml(tbl);
        if (objRefs) objRefs.push(b);
        text += "[table]\n";
      }
    }
    return { html, text: text.replace(/\n+$/, ""), paraCount, charCount, hasObjects };
  };

  /** Parse the inspector's edited HTML back into model blocks, restoring the
   *  preserved objects: a node carrying data-cw-ref re-injects objRefs[k]
   *  verbatim; everything between objects is parsed as paragraph runs. Block ids
   *  are left empty — replaceSdtBlockSpan assigns fresh ones and re-tags runs. */
  const htmlToBlocksPreserving = (html: string, objRefs: Block[]): Block[] => {
    const parsed = new DOMParser().parseFromString(html, "text/html");
    const out: Block[] = [];
    let buffer = "";
    const flushText = (): void => {
      const frag = buffer.trim() === "" && !/&nbsp;|<br/i.test(buffer) ? null : htmlToFragment(buffer);
      buffer = "";
      if (!frag) return;
      for (const fb of frag.blocks) {
        out.push({ kind: "paragraph", id: "", revision: 0, runs: fb.runs, style: fb.style });
      }
    };
    const refOf = (el: HTMLElement): number | null => {
      const direct = el.getAttribute("data-cw-ref");
      if (direct !== null) return Number(direct);
      // contenteditable can wrap the marked div; honor a descendant ref too.
      const inner = el.querySelector?.("[data-cw-ref]")?.getAttribute("data-cw-ref");
      return inner != null ? Number(inner) : null;
    };
    for (const node of Array.from(parsed.body.childNodes)) {
      const el = node instanceof HTMLElement ? node : null;
      const k = el ? refOf(el) : null;
      if (k !== null && objRefs[k]) {
        flushText();
        out.push(objRefs[k]!);
      } else {
        buffer += el ? el.outerHTML : escapeForBuffer(node.textContent ?? "");
      }
    }
    flushText();
    return out;
  };
  const escapeForBuffer = (s: string): string => {
    const d = document.createElement("span");
    d.textContent = s;
    return d.innerHTML;
  };

  /** Split the inspector's edited HTML into one run-list per top-level block
   *  (paragraph), preserving order and empties — so a cell-hosted control's N
   *  paragraphs map 1:1 back to its N tagged cell ranges. */
  const editedBlockRuns = (html: string): import("@cw/shared").Run[][] => {
    const parsed = new DOMParser().parseFromString(html, "text/html");
    const out: import("@cw/shared").Run[][] = [];
    for (const node of Array.from(parsed.body.childNodes)) {
      if (node instanceof HTMLElement) {
        out.push(htmlToFragment(node.outerHTML)?.blocks[0]?.runs ?? []);
      } else if (node.nodeType === Node.TEXT_NODE && (node.textContent ?? "").trim() !== "") {
        out.push(htmlToFragment(`<p>${escapeForBuffer(node.textContent ?? "")}</p>`)?.blocks[0]?.runs ?? []);
      }
    }
    return out;
  };

  /** `blockLevel` controls round-trip through whole-block replacement (objects
   *  preserved by ref); inline controls round-trip through a paragraph fragment.
   *  `objRefs` holds the block-level span's images/tables in data-cw-ref order. */
  type SdtData = SdtInspectorData & { hasObjects: boolean; blockLevel: boolean; objRefs: Block[] };
  const sdtInspectorData = (id: string): SdtData | null => {
    const props = doc.sdts?.[id];
    if (!props) return null;
    const ranges = findSdtRanges(doc, id);
    if (ranges.length === 0) {
      return { id, props, html: "", text: "", paragraphCount: 0, charCount: 0, hasObjects: false, blockLevel: false, objRefs: [] };
    }
    const first = ranges[0]!;
    const last = ranges[ranges.length - 1]!;
    const firstBlock = blockById(doc, first.blockId);
    const inlinePartial =
      ranges.length === 1 &&
      firstBlock !== undefined &&
      (first.start > 0 || first.end < textOfRuns(firstBlock.runs).length);
    // Block-level control: render its whole top-level block span (so images and
    // blank lines between tagged paragraphs survive). Capture objects into
    // objRefs so an edit can restore them verbatim.
    if (!inlinePartial) {
      const fc = containerOf(doc, first.blockId);
      const lc = containerOf(doc, last.blockId);
      if (fc && lc && fc.where === lc.where) {
        const span = containerBlocks(doc, fc.where).slice(fc.index, lc.index + 1);
        const objRefs: Block[] = [];
        const r = renderBlocksHtml(span, objRefs);
        return { id, props, html: r.html, text: r.text, paragraphCount: r.paraCount, charCount: r.charCount, hasObjects: r.hasObjects, blockLevel: true, objRefs };
      }
    }
    // Inline / cell-hosted: per-range run slices (text only).
    const blocks = ranges
      .map((rr) => {
        const block = blockById(doc, rr.blockId);
        return block ? { runs: sliceRuns(block.runs, rr.start, rr.end), style: { ...block.style } } : null;
      })
      .filter((b): b is { runs: import("@cw/shared").Run[]; style: ParaStyle } => b !== null);
    const fragment: DocFragment = { blocks, inline: blocks.length === 1 };
    const text = fragmentToPlainText(fragment);
    return { id, props, html: fragmentToHtml(fragment), text, paragraphCount: blocks.length, charCount: text.length, hasObjects: false, blockLevel: false, objRefs: [] };
  };
  const openSdtInspector = (id: string): void => {
    const data = sdtInspectorData(id);
    if (!data) return;
    const props = data.props;
    // Editable for unlocked text controls. Block-level controls (incl. ones with
    // images/tables) round-trip through whole-block replacement that preserves
    // objects by ref, so they no longer need the read-only guard. Value-driven
    // and locked controls stay read-only.
    const editable =
      !props.lockContent &&
      props.type !== "checkbox" &&
      props.type !== "dropDown";
    sdtInspector?.close();
    sdtInspector = showSdtInspector(data, {
      editable,
      onSave: (html: string): boolean => {
        const before = doc;
        const ranges = findSdtRanges(doc, id);
        const multiBlock = new Set(ranges.map((r) => r.blockId)).size > 1;
        if (data.blockLevel) {
          // Body/band block-level control: whole-span replacement, objects by ref.
          dispatch(replaceSdtBlockSpan(id, htmlToBlocksPreserving(html, data.objRefs)));
        } else if (multiBlock) {
          // Cell-hosted control spanning several cells: rewrite each cell in place.
          const edited = editedBlockRuns(html);
          dispatch(replaceSdtCellContent(id, ranges.map((_, i) => edited[i] ?? [])));
        } else {
          // True inline control (one paragraph): fragment replacement.
          const fragment = htmlToFragment(html) ?? emptyFragmentLike(id);
          dispatch(replaceSdtContent(id, fragment));
        }
        return doc !== before; // dispatch swapped doc iff the edit applied
      },
    });
  };
  /** A one-empty-run fragment carrying the control's base style — used when the
   *  user clears the editable content entirely (htmlToFragment returns null). */
  const emptyFragmentLike = (id: string): DocFragment => {
    const r = findSdtRanges(doc, id)[0];
    const block = r ? blockById(doc, r.blockId) : undefined;
    const style = block ? (styleAtRuns(block.runs, r!.start + 1) ?? block.runs[0]?.style) : undefined;
    const para: ParaStyle = block?.style ?? {
      align: "left", lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0,
      indentFirstLinePx: 0, indentLeftPx: 0,
    };
    const fallbackChar: CharStyle = {
      fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false,
      underline: false, strikethrough: false, color: "#202124",
    };
    return {
      inline: true,
      blocks: [{ runs: [{ text: "", style: style ?? fallbackChar }], style: para }],
    };
  };
  /** Inspect the control at the caret (ribbon button). Returns false if none. */
  const inspectSdtAtCaret = (): boolean => {
    const focus = selection?.focus;
    const id = focus ? sdtAtPosition(doc, focus) : null;
    if (!id || !doc.sdts?.[id]) return false;
    openSdtInspector(id);
    return true;
  };

  const positionWithinSelection = (pos: DocPosition): boolean => {
    if (!selection || isCollapsed(selection)) return false;
    const cmp = comparePositions(tree, selection.anchor, selection.focus, scope());
    const [min, max] = cmp <= 0 ? [selection.anchor, selection.focus] : [selection.focus, selection.anchor];
    return (
      comparePositions(tree, min, pos, scope()) <= 0 &&
      comparePositions(tree, pos, max, scope()) <= 0
    );
  };

  const bandAtPoint = (pt: { pageIndex: number; y: number }): "header" | "footer" | null => {
    const pg = tree.pages[pt.pageIndex];
    if (!pg) return null;
    if (pt.y < pg.contentTopPx) return pg.headerSource || doc.section.header ? "header" : null;
    if (pt.y > pg.contentBottomPx) return pg.footerSource || doc.section.footer ? "footer" : null;
    return null;
  };

  /** First caret-capable paragraph of the cell holding `imageId` (so the Table
   *  section's commands, which read the caret's cell, apply to an image's cell). */
  const caretIntoImageCell = (imageId: string): void => {
    const loc = locateImage(doc, imageId);
    if (loc?.kind !== "cell") return;
    const table = doc.blocks[loc.bi] as TableBlock;
    const para = table.rows[loc.ri]!.cells[loc.ci]!.blocks.find((b) => b.kind === "paragraph");
    if (para) {
      selection = { anchor: { blockId: para.id, offset: 0 }, focus: { blockId: para.id, offset: 0 } };
    }
  };

  const listKindOf = (p: import("@cw/shared").Paragraph): "bullet" | "decimal" | null => {
    const ref = p.style.list;
    if (!ref) return null;
    const def = doc.lists?.[ref.listId];
    const level = def?.levels[Math.min(ref.level, def.levels.length - 1)];
    return level?.format === "bullet" ? "bullet" : "decimal";
  };

  const buildContextEntries = (pt: { pageIndex: number; x: number; y: number }): MenuEntry[] => {
    const item = (
      label: string,
      onClick: () => void,
      opts: { icon?: string; shortcut?: string; disabled?: boolean; danger?: boolean } = {},
    ): MenuEntry => ({ kind: "item", label, onClick, ...opts });
    const sep: MenuEntry = { kind: "sep" };

    const hasSel = !!selection && !isCollapsed(selection);
    const focus = selection?.focus ?? null;
    const para = focus ? blockById(doc, focus.blockId) : undefined;
    const loc = focus ? locateParagraph(doc, focus.blockId) : null;
    const imgId = selectedObject;
    const imgInCell = imgId ? locateImage(doc, imgId)?.kind === "cell" : false;
    const inCell = loc?.kind === "cell" || imgInCell;
    const sdtId = focus && !imgId ? sdtAtPosition(doc, focus) : null;
    const linkUrl = linkAt(tree, pt.pageIndex, pt.x, pt.y, scope());
    const band = bandAtPoint(pt);

    const entries: MenuEntry[] = [];

    // Clipboard — always.
    entries.push(
      item("Cut", () => void copySelection().then(() => dispatch(deleteBackward())), {
        shortcut: "Ctrl+X",
        disabled: !hasSel || sdtBlocksEdit(),
      }),
      item("Copy", () => void copySelection(), { shortcut: "Ctrl+C", disabled: !hasSel }),
      item("Paste", () => void pasteFromClipboard(), { shortcut: "Ctrl+V", disabled: sdtBlocksEdit() }),
    );

    // Link.
    if (linkUrl) {
      entries.push(
        sep,
        item("Open Hyperlink", () => window.open(linkUrl, "_blank", "noopener"), { icon: ICONS.link }),
        item("Edit Hyperlink…", () => {
          const u = prompt("Link URL:", linkUrl);
          if (u !== null) dispatch(setLinkCmd(u.trim() === "" ? null : u.trim()));
        }),
        item("Remove Hyperlink", () => dispatch(setLinkCmd(null)), { danger: true }),
      );
    }

    // Image.
    if (imgId) {
      entries.push(
        sep,
        {
          kind: "submenu",
          label: "Wrap Text",
          icon: ICONS.wrapSquare,
          items: [
            { kind: "item", label: "In Line with Text", icon: ICONS.wrapInline, onClick: () => dispatch(setImageProps(imgId, { wrap: "block", align: "center" })) },
            { kind: "item", label: "Square", icon: ICONS.wrapSquare, onClick: () => dispatch(setImageProps(imgId, { wrap: "square", align: "left" })) },
          ],
        },
        {
          kind: "submenu",
          label: "Align",
          icon: ICONS.alignLeft,
          items: [
            { kind: "item", label: "Left", icon: ICONS.alignLeft, onClick: () => dispatch(setImageProps(imgId, { align: "left" })) },
            { kind: "item", label: "Center", icon: ICONS.alignCenter, onClick: () => dispatch(setImageProps(imgId, { align: "center" })) },
            { kind: "item", label: "Right", icon: ICONS.alignRight, onClick: () => dispatch(setImageProps(imgId, { align: "right" })) },
          ],
        },
        item("Delete Image", () => {
          selectObject(null);
          dispatch(deleteImage(imgId));
        }, { icon: ICONS.image, danger: true }),
      );
    }

    // Content control.
    if (sdtId) {
      const props = doc.sdts?.[sdtId];
      if (props) {
        entries.push(sep);
        if (props.type === "checkbox") {
          entries.push(item("Toggle Check Box", () => dispatch(toggleSdtCheckbox(sdtId)), { icon: ICONS.sdtCheckbox }));
        }
        if ((props.type === "dropDown" || props.type === "comboBox") && (props.listItems?.length ?? 0) > 0) {
          entries.push({
            kind: "submenu",
            label: "Choose Item",
            icon: ICONS.sdtDropdown,
            items: props.listItems!.map((li) => ({ kind: "item", label: li.display, onClick: () => dispatch(setSdtContent(sdtId, li.display)) })),
          });
        }
        entries.push(
          item("Properties & Edit Content…", () => openSdtInspector(sdtId), { icon: ICONS.sdtText }),
        );
        if (!props.lockControl) {
          entries.push(item("Remove Content Control", () => dispatch(removeContentControl(sdtId, false)), { icon: ICONS.sdtRemove, danger: true }));
        }
      }
    }

    // Text formatting — not for a selected image.
    if (!imgId && focus) {
      entries.push(
        sep,
        item("Bold", () => toggleStyle("bold"), { shortcut: "Ctrl+B" }),
        item("Italic", () => toggleStyle("italic"), { shortcut: "Ctrl+I" }),
        item("Underline", () => toggleStyle("underline"), { shortcut: "Ctrl+U" }),
        {
          kind: "submenu",
          label: "Alignment",
          icon: ICONS.alignLeft,
          items: [
            { kind: "item", label: "Left", icon: ICONS.alignLeft, onClick: () => dispatch(setAlignment("left")) },
            { kind: "item", label: "Center", icon: ICONS.alignCenter, onClick: () => dispatch(setAlignment("center")) },
            { kind: "item", label: "Right", icon: ICONS.alignRight, onClick: () => dispatch(setAlignment("right")) },
            { kind: "item", label: "Justify", icon: ICONS.alignJustify, onClick: () => dispatch(setAlignment("justify")) },
          ],
        },
        item("Bullets", () => dispatch(toggleList("bullet")), { icon: ICONS.bullets }),
        item("Numbering", () => dispatch(toggleList("decimal")), { icon: ICONS.numbering }),
        sep,
        item("Insert Hyperlink…", () => {
          const u = prompt("Link URL:");
          if (u !== null && u.trim() !== "") dispatch(setLinkCmd(u.trim()));
        }, { icon: ICONS.link }),
        {
          kind: "submenu",
          label: "Insert Content Control",
          icon: ICONS.sdtText,
          items: (["richText", "checkbox", "dropDown", "date"] as SdtType[]).map((type) => ({
            kind: "item" as const,
            label: (
              { richText: "Rich Text", checkbox: "Check Box", dropDown: "Drop-Down List", date: "Date Picker" } as Record<string, string>
            )[type] ?? type,
            onClick: () => {
              const props: Parameters<typeof insertContentControl>[1] =
                type === "dropDown"
                  ? { listItems: [{ display: "Item 1", value: "Item 1" }, { display: "Item 2", value: "Item 2" }] }
                  : type === "date"
                    ? { dateFormat: "M/d/yyyy" }
                    : {};
              dispatch(insertContentControl(type, props));
            },
          })),
        },
      );
    }

    // List.
    if (!imgId && para?.style.list) {
      const kind = listKindOf(para) ?? "bullet";
      entries.push(
        sep,
        item("Increase List Level", () => dispatch(changeListLevel(1))),
        item("Decrease List Level", () => dispatch(changeListLevel(-1))),
        item("Remove List", () => dispatch(toggleList(kind)), { danger: true }),
      );
    }

    // Table.
    if (inCell) {
      entries.push(
        sep,
        {
          kind: "submenu",
          label: "Insert",
          icon: ICONS.rowBelow,
          items: [
            { kind: "item", label: "Row Above", icon: ICONS.rowAbove, onClick: () => dispatch(insertTableRowCmd("above")) },
            { kind: "item", label: "Row Below", icon: ICONS.rowBelow, onClick: () => dispatch(insertTableRowCmd("below")) },
            { kind: "item", label: "Column Left", icon: ICONS.colLeft, onClick: () => dispatch(insertTableColumnCmd("left")) },
            { kind: "item", label: "Column Right", icon: ICONS.colRight, onClick: () => dispatch(insertTableColumnCmd("right")) },
          ],
        },
        {
          kind: "submenu",
          label: "Delete",
          icon: ICONS.deleteRow,
          items: [
            { kind: "item", label: "Row", icon: ICONS.deleteRow, danger: true, onClick: () => dispatch(deleteTableRowCmd()) },
            { kind: "item", label: "Column", icon: ICONS.deleteCol, danger: true, onClick: () => dispatch(deleteTableColumnCmd()) },
            { kind: "item", label: "Table", icon: ICONS.deleteTable, danger: true, onClick: () => dispatch(deleteTableCmd()) },
          ],
        },
        item("Merge Cells", () => dispatch(mergeCellsCmd()), { icon: ICONS.mergeCells, disabled: !hasSel }),
        item("Unmerge Cell", () => dispatch(unmergeCellCmd()), { icon: ICONS.unmergeCells }),
      );
    }

    // Header/footer band.
    if (activeStory) {
      entries.push(sep, item("Close Header/Footer", () => setStory(null)));
    } else if (band) {
      entries.push(
        sep,
        item(`Edit ${band === "header" ? "Header" : "Footer"}`, () => setStory({ band, pageIndex: pt.pageIndex })),
      );
    }

    return entries;
  };

  /** Resolve an in-document anchor (TOC entry / cross-ref) to a target block
   *  via the modeled bookmarks (docx w:bookmarkStart). */
  const resolveAnchorTarget = (anchorName: string, _fromBlockId: string | null): string | null => {
    const target = doc.bookmarks?.[anchorName];
    return target && blockById(doc, target) ? target : null;
  };

  /** Imported TOC entries keep their pre-calculated page numbers; this rewrites
   *  each to the target heading's CURRENT page (from the live layout). An entry
   *  is "label <tab> number" with an in-document anchor link to its heading. */
  const recalculateToc = (): number => {
    // blockId → displayed page number of the FIRST page it appears on (body only).
    const pageOfBlock = new Map<string, number>();
    const scan = (blocks: import("./layout/layoutTree").PlacedBlock[], pageNum: number): void => {
      for (const pb of blocks) {
        if (!pageOfBlock.has(pb.blockId)) pageOfBlock.set(pb.blockId, pageNum);
        if (pb.table) for (const row of pb.table.rows) for (const cell of row.cells) scan(cell.blocks, pageNum);
      }
    };
    for (const pg of tree.pages) scan(pg.blocks, pg.number);

    const edits: { blockId: string; start: number; end: number; text: string }[] = [];
    for (const b of doc.blocks) {
      if (b.kind !== "paragraph") continue;
      const anchor = b.runs.find((r) => r.style.link?.startsWith("#"))?.style.link?.slice(1);
      if (!anchor) continue;
      const targetId = doc.bookmarks?.[anchor];
      if (!targetId) continue;
      const num = pageOfBlock.get(targetId);
      if (num === undefined) continue;
      const full = textOfRuns(b.runs);
      const m = full.match(/\t(\d+)(\s*)$/); // trailing tab + page number
      if (!m) continue;
      if (m[1] === String(num)) continue; // already current
      const numStart = full.length - m[2]!.length - m[1]!.length;
      edits.push({ blockId: b.id, start: numStart, end: numStart + m[1]!.length, text: String(num) });
    }
    if (edits.length === 0) return 0;

    dispatch((state) => {
      const ops: import("@cw/shared").Op[] = [];
      for (const e of edits) {
        const block = blockById(state.doc, e.blockId);
        if (!block) continue;
        const style = styleAtRuns(block.runs, e.start) ?? block.runs[0]?.style;
        ops.push({ type: "deleteRange", blockId: e.blockId, start: e.start, end: e.end });
        ops.push({ type: "insertText", at: { blockId: e.blockId, offset: e.start }, text: e.text, ...(style ? { style } : {}) });
      }
      return { ops, selectionAfter: state.selection, origin: "command" };
    });
    return edits.length;
  };

  const onContextMenu = (ev: MouseEvent): void => {
    const pt = paint.clientToPage(ev.clientX, ev.clientY);
    if (!pt) return;
    ev.preventDefault();
    closeContextMenu();
    proxy.focus();

    // Word: right-click places focus unless it lands inside an existing range.
    const imageId = hitTestObject(tree, pt.pageIndex, pt.x, pt.y)?.blockId ?? null;
    if (imageId) {
      selectObject(imageId);
      caretIntoImageCell(imageId); // lets the Table section act on the image's cell
    } else {
      selectObject(null);
      const pos = hitTest(tree, pt.pageIndex, pt.x, pt.y, scope());
      if (pos && !positionWithinSelection(pos)) setSelection({ anchor: pos, focus: pos });
    }
    refreshSelectionVisuals();

    const entries = buildContextEntries(pt);
    if (entries.length > 0) contextMenu = showContextMenu(ev.clientX, ev.clientY, entries);
  };
  container.addEventListener("contextmenu", onContextMenu);

  const applyZoom = (next: number, anchorClientY?: number): void => {
    const before = paint.getZoom();
    paint.setZoom(next);
    const after = paint.getZoom();
    if (after === before) return;
    refreshObjectFrame(); // the selection frame's geometry is zoom-scaled
    // Keep the anchored point (cursor, or viewport center) stationary.
    const rect = container.getBoundingClientRect();
    const anchorY = anchorClientY ?? rect.top + rect.height / 2;
    const yInContent = container.scrollTop + (anchorY - rect.top);
    container.scrollTop = (yInContent * after) / before - (anchorY - rect.top);
    options.onZoomChange?.(after);
  };

  // Ctrl/Cmd + wheel zooms (Word/browser convention), anchored on the cursor so
  // the point under the pointer stays put across the zoom.
  container.addEventListener(
    "wheel",
    (ev: WheelEvent) => {
      if (!ev.ctrlKey && !ev.metaKey) return;
      ev.preventDefault();
      applyZoom(paint.getZoom() * (ev.deltaY < 0 ? 1.1 : 1 / 1.1), ev.clientY);
    },
    { passive: false },
  );

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
    getChangeLog(): Change[] {
      return recorder.changes();
    },
    getChangeHead(): number {
      return recorder.head();
    },
    applyRemoteOps,
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
      notifyChange(); // keep the ribbon's font/size controls in sync
    },
    currentFormat(): CurrentFormat {
      const focus = selection?.focus;
      const block = focus ? blockById(doc, focus.blockId) : undefined;
      const char = block && focus ? styleAtRuns(block.runs, focus.offset) : undefined;
      const effective = { ...(char ?? {}), ...(pendingStyle ?? {}) };
      const list = block?.style.list;
      let listKind: CurrentFormat["listKind"] = null;
      if (list) {
        const def = doc.lists?.[list.listId];
        const level = def?.levels[Math.min(list.level, def.levels.length - 1)];
        if (level) listKind = level.format === "bullet" ? "bullet" : "number";
      }
      return {
        styleId: block?.style.namedStyle ?? (block ? "Normal" : null),
        fontFamily: effective.fontFamily ?? null,
        fontSizePx: effective.fontSizePx ?? null,
        lineHeight: block?.style.lineHeight ?? null,
        bold: effective.bold === true,
        italic: effective.italic === true,
        underline: effective.underline === true,
        strikethrough: effective.strikethrough === true,
        highlight: effective.highlightColor !== undefined && effective.highlightColor !== null,
        superscript: effective.verticalAlign === "super",
        subscript: effective.verticalAlign === "sub",
        align: block?.style.align ?? null,
        listKind,
      };
    },
    align(align: ParaStyle["align"]): void {
      if (selectedObject) {
        if (align !== "justify") dispatch(setImageProps(selectedObject, { align }));
        return;
      }
      dispatch(setAlignment(align));
    },
    inspectContentControl: inspectSdtAtCaret,
    recalculateToc,
    setZoom: (z: number): void => applyZoom(z),
    getZoom: () => paint.getZoom(),
    armFormatPainter,
    cancelFormatPainter,
    copy: (): void => {
      // extractFragment runs synchronously inside copySelection, so the
      // clipboard write captures the selection before any later edit.
      void copySelection();
    },
    cut: (): void => {
      void copySelection();
      dispatch(deleteBackward());
    },
    paste: (): void => {
      void pasteFromClipboard();
    },
    revealBlock: (blockId: string): void => {
      if (!blockById(doc, blockId)) return;
      setSelection({ anchor: { blockId, offset: 0 }, focus: { blockId, offset: 0 } });
      const rect = caretRect(tree, { blockId, offset: 0 });
      if (rect) paint.ensureVisible(rect, "center");
    },
    selectAll: (): void => {
      const paras = doc.blocks.filter((b): b is import("@cw/shared").Paragraph => b.kind === "paragraph");
      const first = paras[0];
      const last = paras[paras.length - 1];
      if (!first || !last) return;
      setSelection({
        anchor: { blockId: first.id, offset: 0 },
        focus: { blockId: last.id, offset: textOfRuns(last.runs).length },
      });
      proxy.focus();
    },
    getLayoutInfo: (): { pageCount: number; currentPage: number } => {
      const pageCount = tree.pages.length;
      let currentPage = 1;
      if (selection) {
        const r = caretRect(tree, selection.focus, scope());
        if (r) currentPage = r.pageIndex + 1;
      }
      return { pageCount, currentPage };
    },
    getSelectedObjectRect: (): { left: number; top: number; width: number; height: number } | null => {
      if (!selectedObject) return null;
      const r = objectRect(tree, selectedObject);
      if (!r) return null;
      const ph = paint.getPageElement(r.pageIndex);
      if (!ph) return null;
      const z = paint.getZoom();
      const pr = ph.getBoundingClientRect();
      return { left: pr.left + r.x * z, top: pr.top + r.y * z, width: r.width * z, height: r.height * z };
    },
    deleteSelectedObject: (): void => {
      if (!selectedObject) return;
      const id = selectedObject;
      selectObject(null);
      dispatch(deleteImage(id));
    },
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
      closeContextMenu();
      sdtInspector?.close();
      container.removeEventListener("keydown", keymapHandler);
      container.removeEventListener("contextmenu", onContextMenu);
      controller.destroy();
      objectFrame.destroy();
      mirror.destroy();
      proxy.destroy();
      paint.destroy();
    },
  };
}
