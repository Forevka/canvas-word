// Composition root: model -> layout -> paint -> input -> editor core -> a11y.
// One-way data flow: input -> command -> transaction -> applyOp* -> new state
// -> incremental layout -> paint + caret + proxy reposition (same frame).

import type { Block, CharStyle, Document, ParaStyle, TableBlock } from "@cw/shared";
import { BAND_CONTAINERS, parseTocInstruction } from "@cw/shared";
import type { BookmarkRange, DocPosition, DocSelection, UserInfo } from "@cw/shared";
import { isCollapsed, colorForId, userDisplayName, freshId, DEFAULT_CHAR_STYLE } from "@cw/shared";
import type { ResolvedBehavior, ResolvedTheme } from "./config";
import { ZOOM_STEP } from "./uiConstants";
import { applyOp, containerBlocks, containerOf, effectiveFractions, locateImage, sliceRuns, type Op } from "@cw/shared";
import { bandParagraphs, blockById, buildTableGrid, containerListOf, gridOriginOfCell, locateParagraph, normalizeRect, paragraphsOf, styleAtRuns, textOfRuns } from "@cw/shared";
import type { CellBorders } from "@cw/shared";
import { createLayoutEngine, type LayoutEngine } from "./layout/engine";
import {
  caretRect,
  cellRangeRects,
  comparePositions,
  hitTest,
  hitTestCell,
  hitTestSelectableObject,
  linkAt,
  objectRect,
  selectionRects,
  type ColumnBoundaryHit,
  type GeoScope,
  type Rect,
} from "./layout/geometry";
import type { LayoutTree, Page, PlacedBlock } from "./layout/layoutTree";
import { computeTocEdits } from "./recalc/recalcToc";
import { createSdtPopup } from "./editor/sdtPopup";
import { createCommentController } from "./editor/commentController";
import { createPaintLayer, type RemoteCaret, type PaintLayerOptions } from "./paint/renderer";
import { createChildDocument, type ChildDocument, type StyleContext } from "./child/childDocument";
export type { ChildDocument, ChildContent, ChildRenderOptions, ChildEditorHandle, StyleContext } from "./child/childDocument";
import { createSelectionController } from "./input/selectionController";
import { createObjectFrame } from "./input/objectController";
import { createImeProxy } from "./input/imeProxy";
import { createKeymapHandler, type StyleKey } from "./input/keymap";
import { extractFragment, fragmentToHtml, fragmentToPlainText, htmlToFragment, tableRectToClipboard, type DocFragment } from "./input/clipboard";
import { parseOoxmlFragment } from "./import/docx/fragment";
import { mediaUrl, registerMediaBytes } from "./media/store";
import { importDocx } from "./import/docx/importDocx";
import { showContextMenu, type ContextMenuHandle, type MenuEntry } from "./ui/contextMenu";
import { showSdtInspector, type SdtInspectorData, type SdtInspectorHandle } from "./ui/sdtInspector";
import { showFieldConstructor } from "./ui/fieldConstructor";
import { showTocProperties } from "./ui/tocProperties";
import { showStyleManager, type StyleManagerHandle } from "./ui/styleManager";
import { showTableProperties, type BorderStyleName, type TablePropertiesHandle } from "./ui/tableProperties";
import { createA11yMirror } from "./a11y/mirror";
import {
  changeListLevel,
  applyTableStyle,
  deleteTableRowCmd,
  deleteTableColumnCmd,
  deleteTableCmd,
  findSdtRanges,
  findFieldRanges,
  fieldAtPosition,
  insertContentControl,
  insertText,
  insertFragment,
  deleteBackward,
  deleteForward,
  deleteImage,
  findTableById,
  insertTableColumnCmd,
  insertTableRowCmd,
  fieldAtBlock,
  insertFieldCmd,
  editFieldCmd,
  updateFieldCmd,
  updateTocFieldCmd,
  setTocSwitchesCmd,
  mergeCellsCmd,
  replaceFieldResultCmd,
  setCellsBordersCmd,
  setCellsShadingCmd,
  removeContentControl,
  replaceBackAndInsert,
  replaceSdtContent,
  replaceSdtBlockSpan,
  replaceSdtCellContent,
  sdtAtPosition,
  sdtStackAtPosition,
  setAlignment,
  setCharStyle as setCharStyleCmd,
  setImageProps,
  setImageLayer,
  bringImageToFront,
  sendImageToBack,
  moveAnchoredImage,
  setLinkCmd,
  setParaProps,
  setSdtContent,
  setTableColFractionsCmd,
  setTableWidthModeCmd,
  splitParagraph,
  toggleCharStyle,
  toggleList,
  toggleSdtCheckbox,
  unmergeCellCmd,
  setTableWidthModeAtSelectionCmd,
} from "./editor/commands";
import type { SdtType } from "@cw/shared";
import { ICONS } from "./ui/icons";
import type { CellSelection, Command, EditMode, EditorState, Transaction } from "./editor/state";
import { UndoManager } from "./editor/undo";
import { ChangeRecorder, type ChangeSink } from "./sync/changeRecorder";
import type { Change, ChangeOrigin } from "@cw/shared";
import {
  applyReviewOp,
  emptyReview,
  gcStructuralReviewLayer,
  rebaseReview,
  type Fragment,
  type ReviewLayer,
  type ReviewOp,
  type ReviewOpEnvelope,
} from "@cw/shared";
import { intercept } from "./review/intercept";
import { decorate } from "./review/decorate";
import { acceptSuggestion, rejectSuggestion, acceptAllSuggestions, rejectAllSuggestions, type Resolution } from "@cw/shared";

/** Editor mode: edit (normal), suggest (edits become tracked-change records), or
 *  view (read-only — every mutation is a no-op; the old `readonly:true`).
 *  Defined in ./editor/state (imported above) and re-exported for the public API. */
export type { EditMode };

export type { ReviewOpEnvelope };

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
  /** Caret/selection context — drives which ribbon buttons are enabled. */
  imageSelected: boolean;
  inTable: boolean;
  inContentControl: boolean;
}

export interface SearchState {
  index: number; // 1-based current match (0 = none)
  total: number;
}

export interface Editor {
  focus(): void;
  getDocument(): Document;
  getSelection(): DocSelection | null;
  /** Move the caret/selection programmatically (null clears it). Unlike
   *  revealBlock/revealBookmark this does NOT scroll — it just sets the anchor so
   *  a following command (insert/format) targets the range. Used by agent tools. */
  setSelection(sel: DocSelection | null): void;
  /** The current layout tree (pages, placed blocks, line boxes, fragments — all
   *  page-local geometry). Read-only snapshot for diagnostics/agent inspection of
   *  text placement and rendering. */
  getLayoutTree(): LayoutTree;
  /** A live snapshot of this document's style context (stylesheet, list defs,
   *  content-control + field maps, page section) — what a child document shares. */
  getStyleContext(): StyleContext;
  /** Create a child document that shares this editor's live styles/fonts/theme and
   *  renders or edits a content slice on canvas (see ./child/childDocument). */
  createChild(): ChildDocument;
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
  /** Drawing-grid overlay (light gridline mesh on every page) + snap-to-grid for
   *  dragging anchored objects. Presentational only — no layout/model change. */
  setShowGrid(show: boolean): void;
  getShowGrid(): boolean;
  setSnapToGrid(snap: boolean): void;
  getSnapToGrid(): boolean;
  /** Grid step in document px (96dpi); drives both the mesh and snapping. */
  setGridSpacing(px: number): void;
  getGridSpacing(): number;
  /** Non-printing formatting marks (space dots, tab arrows, pilcrows, line-break
   *  arrows). Presentational overlay only — no layout/model change. */
  setShowFormattingMarks(show: boolean): void;
  getShowFormattingMarks(): boolean;
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
  /** Select a bookmark's range and scroll it into view (Bookmarks panel "Go To").
   *  No-op if the bookmark is missing or anchored in hidden/unplaced content. */
  revealBookmark(name: string): void;
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
  /** Upsert a remote collaborator's caret (rebased through edits, rendered with
   *  their name). `selection` null hides their caret. */
  setPeerPresence(siteId: string, user: UserInfo | undefined, selection: DocSelection | null): void;
  /** Remove a collaborator's caret (they left). */
  removePeer(siteId: string): void;
  /** True when the editor was created in view-only mode (mutations are no-ops). */
  isReadonly(): boolean;
  // ---- review layer (track changes + comments) ----------------------------
  /** Current editor mode. */
  getMode(): EditMode;
  /** Switch mode. Returns false if `mode` isn't in `allowedModes`. */
  setMode(mode: EditMode): boolean;
  /** Read-only snapshot of the review overlay. */
  getReview(): ReviewLayer;
  /** Replace the review overlay (e.g. a rehydrated snapshot loaded on join) and
   *  repaint. Used by the collab join flow after the document is mounted. */
  seedReview(layer: ReviewLayer): void;
  /** Scroll to and select a suggestion or comment thread by id (panel "go to"). */
  revealReview(id: string): void;
  acceptSuggestion(id: string): void;
  rejectSuggestion(id: string): void;
  acceptAllSuggestions(): void;
  rejectAllSuggestions(): void;
  /** The roster of @-mentionable users (set in the constructor / setKnownUsers). */
  getKnownUsers(): UserInfo[];
  /** Update the @-mentionable user roster at runtime. */
  setKnownUsers(users: UserInfo[]): void;
  /** Add a comment thread anchored to the current selection (start===end ⇒ a
   *  point comment). `body` is a rich fragment; `mentions` are tagged users.
   *  Returns the new thread id, or null if there's no selection. */
  addComment(body: Fragment, mentions?: UserInfo[]): string | null;
  /** Open the floating comment composer anchored to the current selection
   *  (Google-Docs-style bubble). No-op in view mode or with no selection. */
  startComment(): void;
  replyToComment(threadId: string, body: Fragment, mentions?: UserInfo[]): void;
  resolveThread(threadId: string, resolved?: boolean): void;
  /** Apply a review op that arrived from a collaborator (already anchor-rebased
   *  to the local doc by the caller). Not recorded/broadcast. */
  applyRemoteReviewOp(op: ReviewOp): void;
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
  /** Fires whenever the local selection/caret moves (for presence broadcast). */
  onSelectionChange?: (selection: DocSelection | null) => void;
  /** View-only mode: the document renders and stays selectable/copyable, but
   *  every mutation (typing, paste, undo/redo, structural edits) is a no-op.
   *  Remote collaborator edits still apply — a read-only client tracks a live
   *  session, it just can't author. Equivalent to `mode: "view"`. */
  readonly?: boolean;
  /** Initial editor mode. Defaults to "view" when `readonly`, else "edit". */
  mode?: EditMode;
  /** Modes the user may switch to (the mode picker is constrained to these and
   *  setMode rejects others). Omit ⇒ all three. */
  allowedModes?: EditMode[];
  /** Local user identity — attributed onto suggestions/comments authored here.
   *  Without it, suggest mode falls back to anonymous attribution. */
  user?: UserInfo;
  /** Users that can be @-mentioned in comments (the embedder owns the roster).
   *  Updatable at runtime via setKnownUsers. */
  knownUsers?: UserInfo[];
  /** Seed the review overlay (e.g. a layer loaded from the backend). */
  review?: ReviewLayer;
  /** Fires after the review overlay changes (suggestion/comment added, resolved,
   *  rebased) — for panels + the embedder. */
  onReviewChanged?: (review: ReviewLayer) => void;
  /** Fires for each locally-authored review op (suggest mode, accept/reject,
   *  comments). The SyncClient ships these on the review channel; persistence
   *  appends them. */
  onReviewOpRecorded?: (env: ReviewOpEnvelope) => void;
  /** Fires when the mode changes (picker / setMode). */
  onModeChanged?: (mode: EditMode) => void;
  /** Resolve a custom field's content from the host (e.g. its backend). Invoked by
   *  the "Update Field" context-menu action for a non-built-in field; the result is
   *  imported and spliced in as the field's new result. Absent ⇒ no refresh. */
  resolveField?: FieldResolver;
  /** Paint-surface tuning passed to the canvas layer — e.g. `{ chrome: false }`
   *  to drop page shadow/gaps for an embedded child-document editor. */
  paintOptions?: PaintLayerOptions;
  /** Resolved color theme (per-instance). Threaded into the paint layer; omit ⇒
   *  the library default look. */
  theme?: ResolvedTheme;
  /** Resolved behavior tuning (zoom step/clamp, indent step). Omit ⇒ defaults. */
  behavior?: ResolvedBehavior;
}

/** Request passed to a custom-field resolver. */
export interface FieldResolveRequest {
  /** The field's id (Document.fields key). */
  fieldId: string;
  /** Field keyword, uppercased (e.g. "MYCHART"). */
  name: string;
  /** The verbatim field instruction (e.g. ` MYCHART "sales-2026" `). */
  instruction: string;
  /** The collaboration doc id, when the session has one. */
  docId?: string;
}

/** What a field resolver returns:
 *  - a full **.docx** (ArrayBuffer / Uint8Array / Blob) — RECOMMENDED: imported
 *    through the same pipeline as opening a document, so images, tables, lists and
 *    styles all come through (media is content-addressed and survives export);
 *  - or an **OOXML fragment string** (w:p / w:tbl, or a w:document) for simple,
 *    text-only results (no embedded media). */
export type FieldResult = string | ArrayBuffer | Uint8Array | Blob;

/** Host hook: produce a field's result for the given request. */
export type FieldResolver = (req: FieldResolveRequest) => Promise<FieldResult>;

/** Turn a resolver result into model blocks: a .docx goes through the full import
 *  pipeline (worker + media re-homing, so images render and export); an OOXML
 *  fragment string is parsed media-free. */
async function fieldResultToBlocks(result: FieldResult): Promise<Block[]> {
  if (typeof result === "string") return result.trim() === "" ? [] : parseOoxmlFragment(result);
  // ArrayBuffer / Blob pass straight to importDocx; wrap a Uint8Array in a Blob
  // (sidesteps SharedArrayBuffer typing and any subarray view offset).
  const input =
    result instanceof ArrayBuffer || result instanceof Blob ? result : new Blob([new Uint8Array(result)]);
  return (await importDocx(input)).doc.blocks;
}

/** Deep-clone a block list with ONE content control's run markers removed, so the
 *  control's content can be edited as a plain mini-document in a child editor — no
 *  nested control frame, and no chance of producing a second copy of the SDT. The
 *  commit (replaceSdt*) re-applies the marker, re-wrapping the whole result. */
function stripSdtMarker(blocks: Block[], sdtId: string): Block[] {
  const clone = structuredClone(blocks);
  const strip = (path: string[] | undefined): string[] | undefined => {
    if (!path) return undefined;
    const next = path.filter((x) => x !== sdtId);
    return next.length > 0 ? next : undefined;
  };
  const walk = (bs: Block[]): void => {
    for (const b of bs) {
      const bp = strip(b.sdtPath);
      if (bp) b.sdtPath = bp;
      else delete b.sdtPath;
      if (b.kind === "paragraph") {
        for (const r of b.runs) {
          const rp = strip(r.style.sdtPath);
          if (rp) r.style.sdtPath = rp;
          else delete r.style.sdtPath;
        }
      } else if (b.kind === "table") {
        for (const row of b.rows) for (const c of row.cells) walk(c.blocks);
      }
    }
  };
  walk(clone);
  return clone;
}

export function createEditor(
  container: HTMLElement,
  initialDoc: Document,
  options: EditorOptions = {},
): Editor {
  const engine = options.engine ?? createLayoutEngine();
  // Editor mode. "view" short-circuits the mutation pipeline (commit/undo/redo)
  // so the document can be read/selected/copied but never authored; remote ops
  // bypass it. "suggest" routes edits through the review interceptor. `readonly`
  // captures the INITIAL view state for construction-time affordance gating;
  // the mutation gates below check `mode` so a runtime switch to view is honored.
  let mode: EditMode = options.mode ?? (options.readonly ? "view" : "edit");
  const allowedModes = options.allowedModes;
  const readonly = mode === "view";
  // Multiplicative zoom step (toolbar +/- and Ctrl+wheel) — per-instance.
  const zoomStep = options.behavior?.zoomStep ?? ZOOM_STEP;
  const paint = createPaintLayer(container, {
    ...options.paintOptions,
    ...(options.theme ? { theme: options.theme } : {}),
    ...(options.behavior ? { zoomMin: options.behavior.zoomMin, zoomMax: options.behavior.zoomMax } : {}),
  });
  const undoMgr = new UndoManager();
  // Document history: every committed edit (and undo/redo, as forward ops) is
  // recorded as a Change. The base snapshot + this ordered log reconstructs any
  // version (see shared/replay).
  const recorder = new ChangeRecorder(options.docId ?? "local", options.onChangeRecorded);
  // Review overlay (track changes + comments) — a SIBLING of `doc`, never folded
  // into it. Anchors ride mapPosition exactly like bookmarks (rebaseReviewLayer).
  let review: ReviewLayer = options.review ?? emptyReview(options.docId ?? "local", recorder.head());

  let doc = initialDoc;
  let tree: LayoutTree = engine.layout(doc);
  let selection: DocSelection | null = null;
  let cellSelection: CellSelection | null = null; // rectangular table-cell selection
  let pendingStyle: Partial<CharStyle> | null = null;
  let activeStory: GeoScope | null = null; // header/footer story-edit scope
  let savedBodySelection: DocSelection | null = null; // restored on story exit
  paint.setTree(tree);

  const state = (): EditorState => ({ doc, selection, cellSelection, pendingStyle });
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
  /** Union the vertical extent of a set of top-level body blocks into one framed
   *  box per page (a span can break across pages). */
  const boxRectsForBlockIds = (ids: Set<string>): Rect[] => {
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
    const rects: Rect[] = [];
    for (const { top, bot, page } of byPage.values()) {
      rects.push({ pageIndex: page.index, x: page.marginPx.left, y: top, width: page.widthPx - page.marginPx.left - page.marginPx.right, height: bot - top });
    }
    return rects;
  };

  const blockLevelSdtRects = (id: string): Rect[] | null => {
    const body = containerBlocks(doc, "body");
    // A control is block-level only when top-level body blocks (paragraph/table/
    // image) carry its id on Block.sdtPath — then frame the whole contiguous span
    // (incl. a wrapped table). Everything else — including an INLINE control that
    // happens to fill a whole cell paragraph — is text-shaped: return null so the
    // caller frames the run rects, not the containing block/table.
    const blockIdxs: number[] = [];
    body.forEach((b, i) => { if (b.sdtPath?.includes(id)) blockIdxs.push(i); });
    if (blockIdxs.length === 0) return null;
    const lo = Math.min(...blockIdxs);
    const hi = Math.max(...blockIdxs);
    const ids = new Set(body.slice(lo, hi + 1).map((b) => b.id));
    const rects = boxRectsForBlockIds(ids);
    return rects.length > 0 ? rects : null;
  };

  /** Frame + label for one control id: a single bounding box for block-level
   *  controls (Word's boundingBox chrome), else text-shaped per-line rects. */
  const sdtLayerFor = (id: string): { rects: Rect[]; label: string } | null => {
    const props = doc.sdts?.[id];
    if (!props) return null;
    const rects =
      blockLevelSdtRects(id) ??
      findSdtRanges(doc, id).flatMap((r) =>
        selectionRects(tree, { anchor: { blockId: r.blockId, offset: r.start }, focus: { blockId: r.blockId, offset: r.end } }, scope()),
      );
    if (rects.length === 0) return null;
    return { rects, label: props.alias ?? SDT_LABELS[props.type] ?? "Content control" };
  };

  // Content-control chrome: an OUTER→INNER stack of frames (concentric, with a
  // breadcrumb tab) for the nested controls under the caret OR the pointer. Hover
  // takes precedence over the caret so the user can inspect any control by pointing.
  let caretSdtChain: string[] = [];
  let hoverSdtChain: string[] = [];
  const sdtChainKey = (ids: string[]): string => ids.join(">");
  const renderSdtAdornment = (): void => {
    const chain = hoverSdtChain.length > 0 ? hoverSdtChain : caretSdtChain;
    paint.setSdtAdornment(chain.length > 0 ? chain.map(sdtLayerFor).filter((l): l is { rects: Rect[]; label: string } => l !== null) : null);
  };

  /** Word's active-control chrome, recomputed from the caret. Cleared when the
   *  caret leaves all controls (unless the pointer is hovering one). */
  const updateSdtAdornment = (): void => {
    const focus = selection?.focus;
    caretSdtChain = focus ? sdtStackAtPosition(doc, focus) : [];
    renderSdtAdornment();
  };

  let lastHoverKey = "";
  /** Highlight the nested control chain under the pointer (skipped while dragging
   *  or when an object is selected). */
  const updateHoverAdornment = (clientX: number, clientY: number, buttons: number): void => {
    let chain: string[] = [];
    if (buttons === 0 && !selectedObject) {
      const pt = paint.clientToPage(clientX, clientY);
      if (pt && pt.inside) {
        const pos = hitTest(tree, pt.pageIndex, pt.x, pt.y, scope());
        if (pos) chain = sdtStackAtPosition(doc, pos);
      }
    }
    const key = sdtChainKey(chain);
    if (key === lastHoverKey) return;
    lastHoverKey = key;
    hoverSdtChain = chain;
    renderSdtAdornment();
  };

  /** Block-region field (contiguous top-level blocks sharing Block.fieldId): one
   *  framed box per page spanning the content column. */
  const blockLevelFieldRects = (id: string): Rect[] => {
    const ids = new Set(containerBlocks(doc, "body").filter((b) => b.fieldId === id).map((b) => b.id));
    if (ids.size === 0) return [];
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
    const rects: Rect[] = [];
    for (const { top, bot, page } of byPage.values()) {
      rects.push({ pageIndex: page.index, x: page.marginPx.left, y: top, width: page.widthPx - page.marginPx.left - page.marginPx.right, height: bot - top });
    }
    return rects;
  };

  /** Word's field highlight: gray field shading + a labelled tab around the field
   *  (inline runs, or a block region) containing the caret. Cleared when it leaves. */
  const updateFieldAdornment = (): void => {
    const focus = selection?.focus;
    const inlineId = focus ? fieldAtPosition(doc, focus) : null;
    const blockFieldId = !inlineId && focus ? blockById(doc, focus.blockId)?.fieldId ?? null : null;
    const id = inlineId ?? blockFieldId;
    const def = id ? doc.fields?.[id] : undefined;
    if (!id || !def) {
      paint.setFieldAdornment(null);
      return;
    }
    const rects = inlineId
      ? findFieldRanges(doc, id).flatMap((r) =>
          selectionRects(
            tree,
            { anchor: { blockId: r.blockId, offset: r.start }, focus: { blockId: r.blockId, offset: r.end } },
            scope(),
          ),
        )
      : blockLevelFieldRects(id);
    paint.setFieldAdornment({ rects, label: def.name });
  };

  /** Pixel rects for the active cell selection, normalized to whole merged cells. */
  const cellSelectionRects = (): Rect[] => {
    if (!cellSelection) return [];
    const found = findTableById(doc, cellSelection.tableId);
    if (!found) return [];
    const grid = buildTableGrid(found.table);
    const rect = normalizeRect(grid, {
      r0: cellSelection.anchor.row,
      c0: cellSelection.anchor.col,
      r1: cellSelection.focus.row,
      c1: cellSelection.focus.col,
    });
    return cellRangeRects(tree, cellSelection.tableId, rect.r0, rect.c0, rect.r1, rect.c1);
  };

  // Assigned in the review section below; called from refreshSelectionVisuals.
  let updateCommentAffordance: () => void = () => {};

  const applySelectionVisuals = (): void => {
    updateSdtAdornment();
    updateFieldAdornment();
    // A rectangular cell selection paints filled cell rects and hides the caret;
    // it supersedes any lingering text selection visual.
    if (cellSelection) {
      paint.setSelectionRects(cellSelectionRects());
      paint.setCaret(null);
      return;
    }
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

  const refreshSelectionVisuals = (): void => {
    applySelectionVisuals();
    updateCommentAffordance(); // float the comment chip beside a suggest-mode selection
  };

  const notifyChange = (): void => {
    options.onChange?.();
  };

  // ---- remote collaborator presence (live carets) -------------------------
  interface RemotePeer {
    user: UserInfo | undefined;
    color: string;
    selection: DocSelection | null;
  }
  const remotePeers = new Map<string, RemotePeer>();

  const paintRemoteCarets = (): void => {
    const list: RemoteCaret[] = [];
    for (const [siteId, peer] of remotePeers) {
      if (!peer.selection) continue;
      const label = peer.user ? userDisplayName(peer.user) : "";
      const rect = caretRect(tree, peer.selection.focus, scope());
      // A range selection also draws translucent highlight rects (text, table
      // cells, SDT content — anything expressible as a DocSelection).
      const rects = isCollapsed(peer.selection) ? [] : selectionRects(tree, peer.selection, scope());
      if (rect || rects.length) list.push({ siteId, color: peer.color, label, rect, rects });
    }
    paint.setRemoteCarets(list);
  };

  // Rebase peer caret positions through an applied op (same mapper as the local
  // selection), so their carets travel with the text everyone is editing.
  const rebasePeers = (mapPosition: (p: DocPosition) => DocPosition): void => {
    for (const peer of remotePeers.values()) {
      if (peer.selection) {
        peer.selection = {
          anchor: mapPosition(peer.selection.anchor),
          focus: mapPosition(peer.selection.focus),
        };
      }
    }
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
    cellSelection = null; // a text caret/selection supersedes a cell selection
    pendingStyle = null; // moving the caret drops the pending typing style
    refreshSelectionVisuals();
    mirror.sync(state());
    notifyChange();
    options.onSelectionChange?.(selection);
  };

  /** Set (or clear) the rectangular table-cell selection. Repaints the cell rects
   *  and updates the menu/toolbar context. */
  const setCellSelection = (next: CellSelection | null): void => {
    if (!next && !cellSelection) return;
    cellSelection = next;
    refreshSelectionVisuals();
    mirror.sync(state());
    options.onSelectionChange?.(selection);
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
      // locateImage (not doc.blocks) so images inside table cells — e.g. inside a
      // content control — resize too; otherwise the handles drag but nothing moves.
      const img = locateImage(doc, selectedObject)?.image;
      if (!img) return;
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
    // Anchored (out-of-flow) images may bleed past the margins, so they resize up
    // to the full page width; in-flow images stay within the content box.
    const sel = doc.blocks.find((b) => b.id === selectedObject);
    const maxW = sel?.kind === "image" && sel.anchor ? doc.section.pageWidthPx : contentWidth();
    objectFrame.show(rect, maxW);
  };

  const selectObject = (blockId: string | null): void => {
    // View-only: never raise the image frame (resize handles) — clearing (null)
    // still runs so any stale frame can be torn down.
    if (readonly && blockId !== null) return;
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
    // Manual column resize cancels autofit (Word): pin the table to fixed widths,
    // keeping the current column proportions, then drag it as a fixed table.
    if (table.widthMode && table.widthMode !== "fixed") {
      dispatch(setTableWidthModeCmd(hit.tableId, "fixed", effectiveFractions(table)));
    }
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

  /** Bookmark ranges are offset-precise, so they travel with edits — mapped
   *  through each applied op exactly like the selection and peer carets. */
  const rebaseBookmarks = (mapPosition: (p: DocPosition) => DocPosition): void => {
    if (!doc.bookmarks) return;
    const next: Record<string, BookmarkRange> = {};
    let changed = false;
    for (const [name, r] of Object.entries(doc.bookmarks)) {
      const start = mapPosition(r.start);
      const end = mapPosition(r.end);
      if (start.blockId !== r.start.blockId || start.offset !== r.start.offset || end.blockId !== r.end.blockId || end.offset !== r.end.offset) {
        changed = true;
      }
      next[name] = { start, end };
    }
    if (changed) doc = { ...doc, bookmarks: next };
  };

  // Review anchors ride the SAME position-mapper as bookmarks (REVIEW.md §5.1):
  // suggestion + comment ranges travel through every applied core op, then GC
  // drops records whose text was removed.
  const rebaseReviewLayer = (mapPosition: (p: DocPosition) => DocPosition): void => {
    review = rebaseReview(review, mapPosition);
  };

  const notifyReviewChanged = (): void => {
    options.onReviewChanged?.(review);
  };

  const reviewAuthor = (): UserInfo => options.user ?? { id: "anon", firstName: "Anonymous", lastName: "" };

  /** Apply one locally-authored review op: mutate the layer and broadcast/persist
   *  it via the sink (with the core version it depends on). Returns its inverse
   *  for the undo coupling. */
  const applyLocalReviewOp = (op: ReviewOp): ReviewOp => {
    const res = applyReviewOp(review, op);
    review = res.layer;
    options.onReviewOpRecorded?.({ op, dependsOnSeq: recorder.head(), ...(options.user ? { author: options.user } : {}) });
    return res.inverse;
  };

  /** Recompute + repaint review decorations from the current layer + layout
   *  tree (metric-neutral → repaint only, like search highlights). */
  const refreshReviewDecorations = (): void => {
    paint.setReviewDecorations(decorate(review, tree, scope()));
  };

  const runOps = (ops: Op[]): Op[] => {
    const inverses: Op[] = [];
    for (const op of ops) {
      const res = applyOp(doc, op);
      doc = res.doc;
      rebasePeers(res.mapPosition); // keep peer carets aligned with the edited text
      rebaseBookmarks(res.mapPosition);
      rebaseReviewLayer(res.mapPosition); // suggestion/comment anchors travel too
      inverses.unshift(res.inverse);
    }
    // Structural records are block-keyed (no text range to ride mapPosition), so
    // GC them by block existence once the doc reaches its new shape.
    review = gcStructuralReviewLayer(review, doc);
    return inverses;
  };

  const afterMutation = (selectionAfter: DocSelection | null): void => {
    const prevSelection = selection;
    relayout();
    selection = selectionAfter;
    cellSelection = null; // grid coords are invalidated by any structural change
    refreshSelectionVisuals();
    refreshObjectFrame(); // images move/resize with reflow; frame follows
    if (searchQuery) {
      runSearch(); // live re-search while the find bar is open
      paintSearch();
    }
    // Only chase the caret into view when the mutation actually MOVED it.
    // Edits that leave the caret where it is — e.g. dragging a table column
    // grip, which keeps the selection untouched — must not yank the viewport
    // back to a caret parked on another page (the resize "scroll jump" bug).
    const caretMoved =
      !prevSelection ||
      !selection ||
      prevSelection.focus.blockId !== selection.focus.blockId ||
      prevSelection.focus.offset !== selection.focus.offset;
    if (selection && isCollapsed(selection) && caretMoved) {
      const caret = caretRect(tree, selection.focus, scope());
      if (caret) paint.ensureVisible(caret);
    }
    paintRemoteCarets(); // peers' carets re-measured against the new layout
    refreshReviewDecorations(); // suggestion/comment overlays re-measured too
    mirror.sync(state());
    notifyChange();
    options.onSelectionChange?.(selection);
  };

  /** Raw commit: applies core ops + already-decided review ops WITHOUT mode
   *  interception. Couples the review ops into the undo entry so Ctrl+Z reverses
   *  text + records as one action. Used by the mode-aware commit (after
   *  intercept) and by accept/reject. */
  const commitCore = (trn: Transaction, reviewOps: ReviewOp[]): void => {
    const selectionBefore = selection;
    const inverses = runOps(trn.ops);
    // Review ops apply AFTER core ops, so their anchors are in post-core coords.
    const reviewInverses: ReviewOp[] = [];
    for (const rop of reviewOps) reviewInverses.unshift(applyLocalReviewOp(rop));
    if (trn.origin !== "transient" && (trn.ops.length > 0 || reviewOps.length > 0)) {
      undoMgr.record({
        ops: trn.ops,
        inverseOps: inverses,
        reviewOps,
        reviewInverses,
        selectionBefore,
        selectionAfter: trn.selectionAfter,
        origin: trn.origin,
        time: Date.now(),
      });
      // Mirror the core edit into the document-history log (empty for a pure
      // deletion-suggestion; the recorder no-ops on empty ops).
      recorder.record(trn.ops, trn.origin as ChangeOrigin, trn.selectionAfter, Date.now());
    }
    if (reviewOps.length > 0) notifyReviewChanged();
    afterMutation(trn.selectionAfter);
  };

  const commit = (trn: Transaction): void => {
    if (mode === "view") return; // view-only: drop every local mutation
    if (mode === "suggest" && trn.origin !== "transient") {
      // Rewrite destructive edits into non-destructive review records.
      const r = intercept(trn, review, doc, reviewAuthor(), Date.now());
      commitCore(r.core, r.reviewOps);
      return;
    }
    commitCore(trn, []);
  };

  const dispatch = (cmd: Command): void => {
    const trn = cmd(state());
    if (trn) commit(trn);
  };

  const undo = (): void => {
    if (mode === "view") return; // undo/redo bypass commit() — gate them directly
    const entry = undoMgr.popUndo();
    if (!entry) return;
    runOps(entry.inverseOps);
    // Review inverses AFTER the text inverses, so anchors are valid when records
    // are restored (reverse application order — they were unshifted on record).
    if (entry.reviewInverses?.length) {
      for (const rop of entry.reviewInverses) applyLocalReviewOp(rop);
      notifyReviewChanged();
    }
    // Undo is a real forward edit in history terms — record the inverse ops it
    // applied so the log faithfully replays the same end state.
    recorder.record(entry.inverseOps, "undo", entry.selectionBefore, Date.now());
    afterMutation(entry.selectionBefore);
  };

  const redo = (): void => {
    if (mode === "view") return;
    const entry = undoMgr.popRedo();
    if (!entry) return;
    runOps(entry.ops);
    if (entry.reviewOps?.length) {
      for (const rop of entry.reviewOps) applyLocalReviewOp(rop);
      notifyReviewChanged();
    }
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
      rebasePeers(res.mapPosition); // peers' carets travel with the remote edit too
      rebaseBookmarks(res.mapPosition);
      rebaseReviewLayer(res.mapPosition); // remote insert shifts everyone's pins
      if (sel) {
        sel = {
          anchor: res.mapPosition(sel.anchor),
          focus: res.mapPosition(sel.focus),
          ...(sel.goalX !== undefined ? { goalX: sel.goalX } : {}),
        };
      }
    }
    review = gcStructuralReviewLayer(review, doc); // remote removeBlock kills its record
    if (ops.length > 0) notifyReviewChanged();
    afterMutation(sel);
  };

  // ---- review actions (track changes + comments) --------------------------

  /** Pure review-layer action (comments, status) — no core ops. Undoable as a
   *  coupled review-only entry; repaints decorations (no relayout). */
  const recordReviewOnly = (ops: ReviewOp[]): void => {
    if (mode === "view" || ops.length === 0) return;
    const reviewInverses: ReviewOp[] = [];
    for (const op of ops) reviewInverses.unshift(applyLocalReviewOp(op));
    undoMgr.record({
      ops: [],
      inverseOps: [],
      reviewOps: ops,
      reviewInverses,
      selectionBefore: selection,
      selectionAfter: selection,
      origin: "command",
      time: Date.now(),
    });
    notifyReviewChanged();
    refreshReviewDecorations();
  };

  const commitResolution = (r: Resolution | null): void => {
    if (!r) return;
    commitCore({ ops: r.ops, selectionAfter: selection, origin: "command" }, r.reviewOps);
  };

  /** Ordered [start,end] of the current selection (anchor/focus in doc order). */
  const orderedSelectionRange = (): { start: DocPosition; end: DocPosition } | null => {
    if (!selection) return null;
    const cmp = comparePositions(tree, selection.anchor, selection.focus, scope());
    return cmp <= 0
      ? { start: selection.anchor, end: selection.focus }
      : { start: selection.focus, end: selection.anchor };
  };

  // The @-mentionable roster = the embedder's configured base PLUS whoever is
  // live-editing the document right now (reusing presence: remotePeers already
  // carries each peer's UserInfo, maintained by setPeerPresence/removePeer).
  let knownUsers: UserInfo[] = options.knownUsers ?? [];
  const dedupeById = (users: UserInfo[]): UserInfo[] => {
    const seen = new Map<string, UserInfo>();
    for (const u of users) if (u.id && !seen.has(u.id)) seen.set(u.id, u);
    return [...seen.values()];
  };
  const mentionableUsers = (): UserInfo[] =>
    dedupeById([
      ...knownUsers,
      ...[...remotePeers.values()].map((p) => p.user).filter((u): u is UserInfo => !!u),
    ]);
  const mentionField = (mentions?: UserInfo[]) => (mentions && mentions.length > 0 ? { mentions } : {});

  const addComment = (body: Fragment, mentions?: UserInfo[]): string | null => {
    if (mode === "view") return null;
    const range = orderedSelectionRange();
    if (!range) return null;
    const author = reviewAuthor();
    const now = Date.now();
    const thread = {
      id: freshId(),
      anchor: { start: range.start, end: range.end },
      status: "open" as const,
      comments: [{ id: freshId(), author, body, createdAt: now, ...mentionField(mentions) }],
    };
    recordReviewOnly([{ type: "addThread", t: thread }]);
    return thread.id;
  };

  const replyToComment = (threadId: string, body: Fragment, mentions?: UserInfo[]): void => {
    recordReviewOnly([
      { type: "addComment", threadId, c: { id: freshId(), author: reviewAuthor(), body, createdAt: Date.now(), ...mentionField(mentions) } },
    ]);
  };

  const resolveThread = (threadId: string, resolved = true): void => {
    recordReviewOnly([{ type: "setThreadStatus", threadId, status: resolved ? "resolved" : "open" }]);
  };

  // Google-Docs-style comment affordance: in suggest mode, selecting text shows a
  // small floating "comment" chip next to the selection; clicking it expands into
  // the full composer. No ribbon icon needed.
  const comments = createCommentController({
    container,
    caretToContainer: (rect) => paint.caretToContainer(rect),
    getSelection: () => selection,
    getCellSelection: () => cellSelection,
    getTree: () => tree,
    getDoc: () => doc,
    getMode: () => mode,
    scope,
    reviewAuthor,
    mentionableUsers,
    addComment,
    focusProxy: () => proxy.focus(),
  });
  // Public: open the composer directly (kept for API parity).
  const startComment = comments.openComposer;
  // The chip refresh is forward-declared (called from refreshSelectionVisuals + setMode).
  updateCommentAffordance = comments.updateAffordance;

  const applyRemoteReviewOp = (op: ReviewOp): void => {
    review = applyReviewOp(review, op).layer;
    notifyReviewChanged();
    refreshReviewDecorations();
  };

  const seedReview = (layer: ReviewLayer): void => {
    review = layer;
    notifyReviewChanged();
    refreshReviewDecorations();
  };

  const setMode = (next: EditMode): boolean => {
    if (allowedModes && !allowedModes.includes(next)) return false;
    if (next === mode) return true;
    mode = next;
    pendingStyle = null; // clear pending preview on switch
    if (mode !== "suggest") {
      comments.hideChip();
      comments.closeComposer();
    }
    updateCommentAffordance(); // reflect the new mode immediately
    options.onModeChanged?.(mode);
    notifyChange();
    return true;
  };

  // Paint any seeded review overlay once at startup (pages pick it up on mount).
  if (review.suggestions.length > 0 || review.threads.length > 0) refreshReviewDecorations();

  const setPeerPresence = (siteId: string, user: UserInfo | undefined, selection: DocSelection | null): void => {
    remotePeers.set(siteId, { user, color: colorForId(user?.id ?? siteId), selection });
    paintRemoteCarets();
  };

  const removePeer = (siteId: string): void => {
    remotePeers.delete(siteId);
    paintRemoteCarets();
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

  // Dropdown / combo / date chooser shown beside a content control (see ./sdtPopup).
  const sdtPopupCtl = createSdtPopup({
    container,
    caretToContainer: (rect) => paint.caretToContainer(rect),
    getDoc: () => doc,
    getTree: () => tree,
    scope,
    dispatch,
    focusProxy: () => proxy.focus(),
  });
  container.addEventListener("mousedown", () => sdtPopupCtl.close());

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
    if (readonly) return; // view-only: the painter would only ever apply no-op edits
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
    readonly,
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
    getCellSelection: () => cellSelection,
    setCellSelection,
    clientToPage: (x, y) => paint.clientToPage(x, y),
    getGridSpacing: () => paint.getGridSpacing(),
    isSnapToGrid: () => paint.getSnapToGrid(),
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
    applyObjectMove: (blockId, x, y, transient) =>
      dispatch(moveAnchoredImage(blockId, x, y, transient ? "transient" : "command")),
    onTab: tabInTable,
    onSdtPress: sdtPopupCtl.handlePress,
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

  // The clipboard flavors for the active selection: a rectangular cell selection
  // serializes as an HTML table + tab-separated text; otherwise the text range.
  const clipboardPayload = (): { html: string; text: string } | null => {
    if (cellSelection) {
      const found = findTableById(doc, cellSelection.tableId);
      if (!found) return null;
      return tableRectToClipboard(found.table, {
        r0: cellSelection.anchor.row,
        c0: cellSelection.anchor.col,
        r1: cellSelection.focus.row,
        c1: cellSelection.focus.col,
      });
    }
    const o = orderedSelection();
    if (!o) return null;
    const fragment = extractFragment(paragraphsOf(doc), o[0], o[1]);
    return { html: fragmentToHtml(fragment), text: fragmentToPlainText(fragment) };
  };

  const copySelection = async (): Promise<void> => {
    const payload = clipboardPayload();
    if (!payload) return;
    const { html, text } = payload;
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

  /** Plain-text summary of a block list — drives the inspector's empty check,
   *  copy button, and char/paragraph counts. The preview itself is canvas-rendered
   *  by a child document (real fonts/metrics), not an HTML approximation. */
  const blocksSummary = (blocks: Block[]): { text: string; paraCount: number; charCount: number } => {
    let text = "";
    let paraCount = 0;
    let charCount = 0;
    for (const b of blocks) {
      if (b.kind === "paragraph") {
        paraCount++;
        const t = textOfRuns(b.runs);
        charCount += t.length;
        text += t + "\n";
      } else {
        text += (b.kind === "image" ? "[image]" : "[table]") + "\n";
      }
    }
    return { text: text.replace(/\n+$/, ""), paraCount, charCount };
  };

  /** The control's content as model blocks + how it must be committed. Block-level
   *  controls preview/edit their whole top-level block span (paragraphs, images,
   *  tables, blank lines); inline / cell-hosted controls use per-range paragraph
   *  slices. The blocks are rendered by a child document, and read back from the
   *  child editor on Save — no HTML round-trip. */
  type SdtData = SdtInspectorData & { blocks: Block[]; blockLevel: boolean; multiBlock: boolean };
  const sdtInspectorData = (id: string): SdtData | null => {
    const props = doc.sdts?.[id];
    if (!props) return null;
    const ranges = findSdtRanges(doc, id);
    if (ranges.length === 0) {
      return { id, props, text: "", paragraphCount: 0, charCount: 0, blocks: [], blockLevel: false, multiBlock: false };
    }
    const first = ranges[0]!;
    const last = ranges[ranges.length - 1]!;
    const firstBlock = blockById(doc, first.blockId);
    const inlinePartial =
      ranges.length === 1 &&
      firstBlock !== undefined &&
      (first.start > 0 || first.end < textOfRuns(firstBlock.runs).length);
    // Block-level control: its whole top-level block span (images + blank lines
    // between tagged paragraphs survive).
    if (!inlinePartial) {
      const fc = containerOf(doc, first.blockId);
      const lc = containerOf(doc, last.blockId);
      if (fc && lc && fc.where === lc.where) {
        const span = containerBlocks(doc, fc.where).slice(fc.index, lc.index + 1);
        const s = blocksSummary(span);
        return { id, props, text: s.text, paragraphCount: s.paraCount, charCount: s.charCount, blocks: span, blockLevel: true, multiBlock: false };
      }
    }
    // Inline / cell-hosted: one paragraph per tagged range (text only).
    const paras: import("@cw/shared").Paragraph[] = ranges
      .map((rr): import("@cw/shared").Paragraph | null => {
        const block = blockById(doc, rr.blockId);
        return block
          ? { kind: "paragraph", id: `cw-sdt-${rr.blockId}-${rr.start}`, revision: 0, runs: sliceRuns(block.runs, rr.start, rr.end), style: { ...block.style } }
          : null;
      })
      .filter((b): b is import("@cw/shared").Paragraph => b !== null);
    const s = blocksSummary(paras);
    const multiBlock = new Set(ranges.map((r) => r.blockId)).size > 1;
    return { id, props, text: s.text, paragraphCount: s.paraCount, charCount: s.charCount, blocks: paras, blockLevel: false, multiBlock };
  };
  const openSdtInspector = (id: string): void => {
    const data = sdtInspectorData(id);
    if (!data) return;
    const props = data.props;
    // Editable for unlocked text controls; value-driven (checkbox/dropDown) and
    // locked controls stay read-only.
    const editable =
      !props.lockContent &&
      props.type !== "checkbox" &&
      props.type !== "dropDown";
    sdtInspector?.close();
    // A child document sharing this editor's live styles renders the preview and
    // hosts the canvas-native editor; Save reads the edited blocks straight back.
    const child = createChildDocument({ getStyleContext, makeEditor: createEditor });
    const commit = (blocks: Block[]): boolean => {
      const before = doc;
      const ranges = findSdtRanges(doc, id);
      if (data.blockLevel) {
        // Body/band block-level control: whole-span replacement (objects survive).
        dispatch(replaceSdtBlockSpan(id, blocks));
      } else if (data.multiBlock) {
        // Cell-hosted control spanning several cells: rewrite each range in place.
        dispatch(replaceSdtCellContent(id, ranges.map((_, i) => {
          const b = blocks[i];
          return b && b.kind === "paragraph" ? b.runs : [];
        })));
      } else {
        // True inline control (one paragraph): fragment replacement.
        const paras = blocks.filter((b): b is import("@cw/shared").Paragraph => b.kind === "paragraph");
        const fragment: DocFragment = paras.length
          ? { inline: paras.length === 1, blocks: paras.map((p) => ({ runs: p.runs, style: p.style })) }
          : emptyFragmentLike(id);
        dispatch(replaceSdtContent(id, fragment));
      }
      return doc !== before; // dispatch swapped doc iff the edit applied
    };
    // Edit/preview the content in isolation: a clone with THIS control's marker
    // stripped, so the surface shows plain content (no nested control frame) and
    // can't yield a second copy of the SDT. The commit re-wraps it.
    const editBlocks = stripSdtMarker(data.blocks, id);
    sdtInspector = showSdtInspector(data, {
      editable,
      // Images only round-trip through the block-level (whole-span) commit path.
      allowImages: editable && data.blockLevel,
      renderPreview: (host) => child.render(host, { kind: "blocks", blocks: editBlocks }),
      mountEditor: (host) => {
        const h = child.mountEditor(host, { kind: "blocks", blocks: editBlocks });
        return {
          getBlocks: (): Block[] => h.getBlocks(),
          insertImage: (bytes, mime, imgOpts) => h.insertImage(bytes, mime, imgOpts),
        };
      },
      onSave: commit,
      onClose: () => child.destroy(),
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
    const fallbackChar: CharStyle = { ...DEFAULT_CHAR_STYLE };
    return {
      inline: true,
      blocks: [{ runs: [{ text: "", style: style ?? fallbackChar }], style: para }],
    };
  };
  /** Open the field constructor with its result preview rendered by a child
   *  document (the field's result in the document's real font, not plain text). */
  const openFieldConstructor = (
    config: Pick<Parameters<typeof showFieldConstructor>[0], "initial" | "baseStyle" | "onApply">,
  ): void => {
    const child = createChildDocument({ getStyleContext, makeEditor: createEditor });
    showFieldConstructor({
      ...config,
      renderResult: (host, runs) => child.render(host, { kind: "runs", runs }),
      onClose: () => child.destroy(),
    });
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

  // ---- Style Manager -----------------------------------------------------

  let styleMgr: StyleManagerHandle | null = null;
  const openStyleManager = (initialSelection?: { kind: "paragraph" | "character" | "list" | "table"; id: string }): void => {
    if (styleMgr) { styleMgr.refresh(); return; }
    styleMgr = showStyleManager({
      editor: {
        getDocument: () => doc,
        dispatch,
        createChild: () => createChildDocument({ getStyleContext, makeEditor: createEditor }),
        focus: () => proxy.focus(),
      },
      ...(initialSelection ? { initialSelection } : {}),
      onClose: () => { styleMgr = null; },
    });
  };

  // ---- Borders & Shading modal -------------------------------------------

  let tableProps: TablePropertiesHandle | null = null;

  /** True when a page point lands within the active rectangular cell selection. */
  const pointInCellSelection = (pt: { pageIndex: number; x: number; y: number }): boolean => {
    if (!cellSelection) return false;
    const hit = hitTestCell(tree, pt.pageIndex, pt.x, pt.y);
    if (!hit || hit.tableId !== cellSelection.tableId) return false;
    const found = findTableById(doc, cellSelection.tableId);
    if (!found) return false;
    const rect = normalizeRect(buildTableGrid(found.table), {
      r0: cellSelection.anchor.row,
      c0: cellSelection.anchor.col,
      r1: cellSelection.focus.row,
      c1: cellSelection.focus.col,
    });
    return hit.row >= rect.r0 && hit.row <= rect.r1 && hit.col >= rect.c0 && hit.col <= rect.c1;
  };

  /** A single-cell selection at the caret's cell (when no rectangular cell
   *  selection is active), so the modal can target the cell the caret sits in. */
  const singleCellAtCaret = (): CellSelection | null => {
    const focus = selection?.focus;
    if (!focus) return null;
    const loc = locateParagraph(doc, focus.blockId);
    if (loc?.kind !== "cell") return null;
    const table = containerListOf(doc, loc.where)[loc.bi];
    if (!table || table.kind !== "table") return null;
    const origin = gridOriginOfCell(buildTableGrid(table), loc.ri, loc.ci);
    if (!origin) return null;
    return { tableId: table.id, anchor: { row: origin.row, col: origin.col }, focus: { row: origin.row, col: origin.col } };
  };

  /** Seed the modal's border controls from an existing edge, else a 1px black line. */
  const borderSeed = (b: CellBorders | undefined): { color: string; widthPx: number; style: BorderStyleName } => {
    const e = b && (b.top ?? b.right ?? b.bottom ?? b.left);
    return { color: e?.color ?? "#000000", widthPx: e?.widthPx ?? 1, style: e?.style ?? "single" };
  };

  const openTableProperties = (): void => {
    const captured = cellSelection ?? singleCellAtCaret();
    if (!captured) return;
    const found = findTableById(doc, captured.tableId);
    if (!found) return;
    const grid = buildTableGrid(found.table);
    const rect = normalizeRect(grid, {
      r0: captured.anchor.row,
      c0: captured.anchor.col,
      r1: captured.focus.row,
      c1: captured.focus.col,
    });
    const topLeft = grid.slots[rect.r0]?.[rect.c0]?.cell;
    const multiCell = rect.r1 > rect.r0 || rect.c1 > rect.c0;
    const seed = borderSeed(topLeft?.borders);
    tableProps?.close();
    tableProps = showTableProperties(
      {
        color: seed.color,
        widthPx: seed.widthPx,
        style: seed.style,
        shading: topLeft?.shading ?? null,
        rangeLabel: multiCell ? "Selected cells" : "1 cell",
        multiCell,
      },
      {
        applyBorders: (spec, edges) => {
          dispatch(setCellsBordersCmd(spec, edges, captured));
          keepCellSelection(captured); // border edits don't change the grid — keep the highlight
        },
        applyShading: (fill) => {
          dispatch(setCellsShadingCmd(fill, captured));
          keepCellSelection(captured);
        },
      },
    );
  };

  /** Re-assert a cell selection after a property edit that left the grid intact
   *  (afterMutation clears it defensively), so the highlight stays put while the
   *  modal is open. */
  const keepCellSelection = (cs: CellSelection): void => {
    cellSelection = cs;
    refreshSelectionVisuals();
  };

  const buildContextEntries = (pt: { pageIndex: number; x: number; y: number }): MenuEntry[] => {
    const item = (
      label: string,
      onClick: () => void,
      opts: { icon?: string; shortcut?: string; disabled?: boolean; danger?: boolean } = {},
    ): MenuEntry => ({ kind: "item", label, onClick, ...opts });
    const sep: MenuEntry = { kind: "sep" };

    const hasSel = (!!selection && !isCollapsed(selection)) || !!cellSelection;

    // View-only (checked dynamically so a runtime switch to view mode applies):
    // only the non-mutating actions (copy text, follow a link).
    if (mode === "view") {
      const url = linkAt(tree, pt.pageIndex, pt.x, pt.y, scope());
      const ro: MenuEntry[] = [
        item("Copy", () => void copySelection(), { shortcut: "Ctrl+C", disabled: !hasSel }),
      ];
      if (url) ro.push(sep, item("Open Hyperlink", () => window.open(url, "_blank", "noopener"), { icon: ICONS.link }));
      return ro;
    }

    // Suggesting: don't expose the full editing menu (paste, table/object edits
    // and other structural actions aren't tracked as suggestions in V1 and would
    // mutate the document directly). Offer only the tracking-safe + non-mutating
    // actions — Cut becomes a tracked deletion, plus Comment (Google-Docs style).
    if (mode === "suggest") {
      const url = linkAt(tree, pt.pageIndex, pt.x, pt.y, scope());
      const m: MenuEntry[] = [
        item("Cut", () => void copySelection().then(() => dispatch(deleteBackward())), { shortcut: "Ctrl+X", disabled: !hasSel }),
        item("Copy", () => void copySelection(), { shortcut: "Ctrl+C", disabled: !hasSel }),
        sep,
        item("Comment", () => startComment(), { icon: ICONS.comment, disabled: !hasSel }),
      ];
      if (url) m.push(sep, item("Open Hyperlink", () => window.open(url, "_blank", "noopener"), { icon: ICONS.link }));
      return m;
    }

    const hasCellSel = !!cellSelection;
    const focus = selection?.focus ?? null;
    const para = focus ? blockById(doc, focus.blockId) : undefined;
    const loc = focus ? locateParagraph(doc, focus.blockId) : null;
    const imgId = selectedObject;
    const imgInCell = imgId ? locateImage(doc, imgId)?.kind === "cell" : false;
    const inCell = loc?.kind === "cell" || imgInCell || hasCellSel;
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

    // Comment on the selection (edit mode has no floating chip — suggest mode does).
    entries.push(sep, item("Comment", () => startComment(), { icon: ICONS.comment, disabled: !hasSel }));

    // Styles — open the manager, seeded to the caret paragraph's style.
    {
      const styleId = para?.kind === "paragraph" ? para.style.namedStyle : undefined;
      entries.push(item("Styles…", () => openStyleManager(styleId ? { kind: "paragraph", id: styleId } : undefined), { icon: ICONS.stylePencil }));
    }

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

    // Field — regenerate a TOC (Word's F9) or refresh a custom field from the host
    // via the resolveField hook (parse its OOXML, splice in as the new result).
    const field = focus ? fieldAtBlock(doc, focus.blockId) : null;
    if (field?.kind === "toc") {
      entries.push(
        sep,
        item("Update Field (TOC)", () => dispatch(updateTocFieldCmd())),
        item("Table of Contents options…", () =>
          showTocProperties({
            initial: parseTocInstruction(doc.tocInstruction ?? ' TOC \\o "1-3" \\h \\z '),
            onApply: (sw) => dispatch(setTocSwitchesCmd(sw)),
          }),
        ),
      );
    } else if (field?.kind === "custom") {
      const def = field.def;
      const resolve = options.resolveField;
      entries.push(
        sep,
        item(
          `Update Field (${def.name})`,
          () => {
            if (!resolve) return;
            void (async () => {
              try {
                const result = await resolve({
                  fieldId: def.id,
                  name: def.name,
                  instruction: def.instruction,
                  ...(options.docId ? { docId: options.docId } : {}),
                });
                const blocks = await fieldResultToBlocks(result);
                if (blocks.length > 0) dispatch(replaceFieldResultCmd(def.id, blocks));
              } catch (err) {
                console.error("resolveField failed", err);
              }
            })();
          },
          { disabled: !resolve },
        ),
      );
    }

    // Inline built-in field (PAGE/NUMPAGES/DATE/TIME/IF) at the caret — edit its
    // definition in the constructor, or recompute its result (Word's F9).
    const caretStyle = (pos: { blockId: string; offset: number }): CharStyle => {
      const b = blockById(doc, pos.blockId);
      const s = b ? styleAtRuns(b.runs, pos.offset) ?? b.runs[0]?.style : undefined;
      if (s) return s;
      for (const blk of doc.blocks) if (blk.kind === "paragraph" && blk.runs[0]) return blk.runs[0]!.style;
      return { fontFamily: "Arial, sans-serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" };
    };
    const inlineFieldId = focus ? fieldAtPosition(doc, focus) : null;
    const inlineFieldDef = inlineFieldId ? doc.fields?.[inlineFieldId] : undefined;
    if (focus && inlineFieldDef?.kind === "builtin" && inlineFieldDef.spec) {
      const def = inlineFieldDef;
      const spec = def.spec!;
      const f = focus;
      entries.push(sep, item(`Edit Field (${def.name})…`, () =>
        openFieldConstructor({ initial: spec, baseStyle: caretStyle(f), onApply: (s) => dispatch(editFieldCmd(def.id, s)) }),
      ));
      if (spec.type !== "PAGE" && spec.type !== "NUMPAGES") {
        entries.push(item(`Update Field (${def.name})`, () => dispatch(updateFieldCmd(def.id))));
      }
    } else if (focus) {
      entries.push(sep, item("Insert Field…", () =>
        openFieldConstructor({ baseStyle: caretStyle(focus), onApply: (s) => dispatch(insertFieldCmd(s)) }),
      ));
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
            { kind: "item", label: "In Line with Text", icon: ICONS.wrapInline, onClick: () => dispatch(setImageProps(imgId, { wrap: "block", align: "center", anchor: null })) },
            { kind: "item", label: "Square", icon: ICONS.wrapSquare, onClick: () => dispatch(setImageProps(imgId, { wrap: "square", align: "left", anchor: null })) },
            { kind: "sep" },
            { kind: "item", label: "Behind Text", onClick: () => dispatch(setImageLayer(imgId, true)) },
            { kind: "item", label: "In Front of Text", onClick: () => dispatch(setImageLayer(imgId, false)) },
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
        item("Bring to Front", () => dispatch(bringImageToFront(imgId))),
        item("Send to Back", () => dispatch(sendImageToBack(imgId))),
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
        item("Merge Cells", () => dispatch(mergeCellsCmd()), { icon: ICONS.mergeCells, disabled: !hasSel && !hasCellSel }),
        item("Unmerge Cell", () => dispatch(unmergeCellCmd()), { icon: ICONS.unmergeCells }),
        (() => {
          const captured = cellSelection ?? singleCellAtCaret();
          const curMode = (captured ? findTableById(doc, captured.tableId)?.table.widthMode : undefined) ?? "fixed";
          const fit = (label: string, mode: "fixed" | "autofitContents" | "autofitWindow"): MenuEntry => ({
            kind: "item",
            label: curMode === mode ? `${label} ✓` : label,
            onClick: () => dispatch(setTableWidthModeAtSelectionCmd(mode)),
          });
          return {
            kind: "submenu",
            label: "AutoFit",
            icon: ICONS.table,
            items: [
              fit("AutoFit to Contents", "autofitContents"),
              fit("AutoFit to Window", "autofitWindow"),
              fit("Fixed Column Width", "fixed"),
            ],
          } as MenuEntry;
        })(),
        sep,
        item("Borders & Shading…", () => openTableProperties(), { icon: ICONS.borders }),
        (() => {
          const captured = cellSelection ?? singleCellAtCaret();
          const curStyleId = captured ? findTableById(doc, captured.tableId)?.table.styleId : undefined;
          const styleItems: MenuEntry[] = Object.values(doc.tableStyles ?? {}).map((ts) => ({
            kind: "item",
            label: ts.id === curStyleId ? `${ts.name} ✓` : ts.name,
            onClick: () => dispatch(applyTableStyle(ts.id)),
          }));
          return {
            kind: "submenu",
            label: "Table Style",
            icon: ICONS.stylePencil,
            items: [
              ...(styleItems.length > 0 ? [...styleItems, { kind: "sep" } as MenuEntry] : []),
              { kind: "item", label: "New Table Style…", onClick: () => openStyleManager(curStyleId ? { kind: "table", id: curStyleId } : undefined) },
              { kind: "item", label: "Manage Styles…", onClick: () => openStyleManager(curStyleId ? { kind: "table", id: curStyleId } : undefined) },
            ],
          } as MenuEntry;
        })(),
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
    const target = doc.bookmarks?.[anchorName]?.start.blockId;
    return target && blockById(doc, target) ? target : null;
  };

  /** Imported TOC entries keep their pre-calculated page numbers; this rewrites
   *  each to the target heading's CURRENT page (from the live layout). An entry
   *  is "label <tab> number" with an in-document anchor link to its heading. */
  const recalculateToc = (): number => {
    // The page-number mapping lives in the shared, headless core (also used by the
    // backend recalc route). The editor owns only the dispatch wrapping so the
    // rewrite is undoable and replays to collaborators.
    const edits = computeTocEdits(doc, tree);
    if (edits.length === 0) return 0;

    dispatch((state) => {
      const ops: import("@cw/shared").Op[] = [];
      for (const e of edits) {
        const block = blockById(state.doc, e.blockId);
        if (!block) continue;
        // Re-resolve style against current state for collab-safety, else the
        // style captured by the core (against the laid-out doc).
        const style = styleAtRuns(block.runs, e.start) ?? e.style;
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

    // Word: right-click places focus unless it lands inside an existing range —
    // including a rectangular cell selection, which must survive so Merge /
    // Borders act on the whole dragged block, not just the clicked cell.
    if (!pointInCellSelection(pt)) {
      // Off-page right-clicks are clamped onto the edge — don't let them "hit" a
      // background image there (mirrors the left-click rule).
      const imageId = pt.inside ? hitTestSelectableObject(tree, pt.pageIndex, pt.x, pt.y, scope())?.blockId ?? null : null;
      if (imageId) {
        selectObject(imageId);
        caretIntoImageCell(imageId); // lets the Table section act on the image's cell
      } else {
        selectObject(null);
        const pos = hitTest(tree, pt.pageIndex, pt.x, pt.y, scope());
        if (pos && !positionWithinSelection(pos)) setSelection({ anchor: pos, focus: pos });
      }
      refreshSelectionVisuals();
    }

    const entries = buildContextEntries(pt);
    if (entries.length > 0) contextMenu = showContextMenu(ev.clientX, ev.clientY, entries);
  };
  container.addEventListener("contextmenu", onContextMenu);

  // Hover highlighting for content controls (incl. nested) — point at any control
  // to see its frame(s) and breadcrumb, without moving the caret.
  const onSdtHoverMove = (ev: MouseEvent): void => updateHoverAdornment(ev.clientX, ev.clientY, ev.buttons);
  const onSdtHoverLeave = (): void => {
    if (hoverSdtChain.length === 0) return;
    hoverSdtChain = [];
    lastHoverKey = "";
    renderSdtAdornment();
  };
  container.addEventListener("mousemove", onSdtHoverMove);
  container.addEventListener("mouseleave", onSdtHoverLeave);

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
      applyZoom(paint.getZoom() * (ev.deltaY < 0 ? zoomStep : 1 / zoomStep), ev.clientY);
    },
    { passive: false },
  );

  // Soft keyboard: when it opens, the VISUAL viewport shrinks but the document's
  // native scroll doesn't know the caret is now behind it. On a visualViewport
  // resize/scroll, re-center a focused, collapsed caret so it stays readable
  // above the keyboard. (Android resizes the layout viewport so centering is
  // exact; iOS overlays the keyboard, where centering keeps the caret in the
  // upper half, clear of a typical keyboard.)
  const vv = window.visualViewport;
  const onViewportChange = (): void => {
    if (!proxy.hasFocus() || !selection || !isCollapsed(selection)) return;
    const caret = caretRect(tree, selection.focus, scope());
    if (caret) paint.ensureVisible(caret, "center");
  };
  if (vv) {
    vv.addEventListener("resize", onViewportChange);
    vv.addEventListener("scroll", onViewportChange);
  }

  // Pinch-to-zoom (touch): two fingers zoom the DOCUMENT through the same
  // anchor-aware applyZoom the wheel uses — not the browser chrome (#app has
  // touch-action that disables native pinch there). The pinch MIDPOINT Y is the
  // zoom anchor, so the point between the fingers stays put. applyZoom calls are
  // throttled to one per frame to avoid relayout/paint thrash.
  const pinchPts = new Map<number, { x: number; y: number }>();
  let pinchDist = 0;
  let pinchZoom = 1;
  let pinchRaf: number | null = null;
  let pinchNext: { scale: number; midY: number } | null = null;
  const pinchPair = (): [{ x: number; y: number }, { x: number; y: number }] | null => {
    if (pinchPts.size !== 2) return null;
    const it = pinchPts.values();
    return [it.next().value!, it.next().value!];
  };
  const flushPinch = (): void => {
    pinchRaf = null;
    if (!pinchNext) return;
    applyZoom(pinchZoom * pinchNext.scale, pinchNext.midY);
    pinchNext = null;
  };
  const onPinchDown = (ev: PointerEvent): void => {
    if (ev.pointerType === "mouse") return;
    pinchPts.set(ev.pointerId, { x: ev.clientX, y: ev.clientY });
    const pair = pinchPair();
    if (pair) {
      pinchDist = Math.hypot(pair[0].x - pair[1].x, pair[0].y - pair[1].y);
      pinchZoom = paint.getZoom();
    }
  };
  const onPinchMove = (ev: PointerEvent): void => {
    if (!pinchPts.has(ev.pointerId)) return;
    pinchPts.set(ev.pointerId, { x: ev.clientX, y: ev.clientY });
    const pair = pinchPair();
    if (!pair || pinchDist === 0) return;
    ev.preventDefault();
    const dist = Math.hypot(pair[0].x - pair[1].x, pair[0].y - pair[1].y);
    pinchNext = { scale: dist / pinchDist, midY: (pair[0].y + pair[1].y) / 2 };
    if (pinchRaf === null) pinchRaf = requestAnimationFrame(flushPinch);
  };
  const onPinchUp = (ev: PointerEvent): void => {
    if (!pinchPts.delete(ev.pointerId)) return;
    if (pinchPts.size < 2) pinchDist = 0; // a finger lifted: end this pinch
  };
  container.addEventListener("pointerdown", onPinchDown);
  container.addEventListener("pointermove", onPinchMove, { passive: false });
  container.addEventListener("pointerup", onPinchUp);
  container.addEventListener("pointercancel", onPinchUp);

  // Live style context for child documents — read off the CURRENT doc each call
  // so children reflect edits (e.g. a redefined style) the moment they re-render.
  const getStyleContext = (): StyleContext => {
    const ctx: StyleContext = { section: doc.section };
    if (doc.stylesheet) ctx.stylesheet = doc.stylesheet;
    if (doc.lists) ctx.lists = doc.lists;
    if (doc.sdts) ctx.sdts = doc.sdts;
    if (doc.fields) ctx.fields = doc.fields;
    ctx.defaultStyleId = doc.stylesheet?.defaultStyleId;
    // Proxy over this document's media/relationships: the session store is shared,
    // so a child resolves the parent's images and registers new ones back into it.
    ctx.media = { resolve: (id) => mediaUrl(id), register: (bytes, mime) => registerMediaBytes(bytes, mime) };
    return ctx;
  };

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
    setSelection(sel: DocSelection | null): void {
      setSelection(sel);
    },
    getLayoutTree(): LayoutTree {
      return tree;
    },
    getStyleContext,
    createChild: (): ChildDocument =>
      createChildDocument({
        getStyleContext,
        makeEditor: createEditor,
        ...(options.theme ? { theme: options.theme } : {}),
        ...(options.behavior ? { behavior: options.behavior } : {}),
      }),
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
    setPeerPresence,
    removePeer,
    isReadonly: (): boolean => mode === "view",
    getMode: (): EditMode => mode,
    setMode,
    getReview: (): ReviewLayer => review,
    seedReview,
    // The effective roster: configured base + live editors (presence-merged).
    getKnownUsers: (): UserInfo[] => mentionableUsers(),
    setKnownUsers: (users: UserInfo[]): void => { knownUsers = users; },
    acceptSuggestion: (id: string): void => commitResolution(acceptSuggestion(doc, review, id)),
    rejectSuggestion: (id: string): void => commitResolution(rejectSuggestion(doc, review, id)),
    acceptAllSuggestions: (): void => commitResolution(acceptAllSuggestions(doc, review)),
    rejectAllSuggestions: (): void => commitResolution(rejectAllSuggestions(doc, review)),
    addComment,
    startComment,
    replyToComment,
    resolveThread,
    applyRemoteReviewOp,
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
        imageSelected: selectedObject !== null,
        inTable: !!cellSelection || (focus ? locateParagraph(doc, focus.blockId)?.kind === "cell" : false),
        inContentControl: focus && !selectedObject ? sdtAtPosition(doc, focus) !== null : false,
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
    setShowGrid: (show: boolean): void => paint.setShowGrid(show),
    getShowGrid: () => paint.getShowGrid(),
    setSnapToGrid: (snap: boolean): void => paint.setSnapToGrid(snap),
    getSnapToGrid: () => paint.getSnapToGrid(),
    setGridSpacing: (px: number): void => paint.setGridSpacing(px),
    getGridSpacing: () => paint.getGridSpacing(),
    setShowFormattingMarks: (show: boolean): void => paint.setShowFormattingMarks(show),
    getShowFormattingMarks: () => paint.getShowFormattingMarks(),
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
    revealBookmark: (name: string): void => {
      const range = doc.bookmarks?.[name];
      if (!range || !blockById(doc, range.start.blockId)) return;
      const rect = caretRect(tree, range.start, scope());
      if (!rect) return; // anchored in hidden/unplaced content — nothing to show
      setSelection({ anchor: range.start, focus: range.end });
      paint.ensureVisible(rect, "center");
    },
    revealReview: (id: string): void => {
      const anchor = (review.suggestions.find((s) => s.id === id) ?? review.threads.find((t) => t.id === id))?.anchor;
      if (!anchor || !blockById(doc, anchor.start.blockId)) return;
      const rect = caretRect(tree, anchor.start, scope());
      if (!rect) return; // anchored in unplaced content
      setSelection({ anchor: anchor.start, focus: anchor.end });
      paint.ensureVisible(rect, "center");
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
      tableProps?.close();
      container.removeEventListener("keydown", keymapHandler);
      container.removeEventListener("contextmenu", onContextMenu);
      container.removeEventListener("mousemove", onSdtHoverMove);
      container.removeEventListener("mouseleave", onSdtHoverLeave);
      if (vv) {
        vv.removeEventListener("resize", onViewportChange);
        vv.removeEventListener("scroll", onViewportChange);
      }
      if (pinchRaf !== null) cancelAnimationFrame(pinchRaf);
      container.removeEventListener("pointerdown", onPinchDown);
      container.removeEventListener("pointermove", onPinchMove);
      container.removeEventListener("pointerup", onPinchUp);
      container.removeEventListener("pointercancel", onPinchUp);
      controller.destroy();
      objectFrame.destroy();
      mirror.destroy();
      proxy.destroy();
      paint.destroy();
    },
  };
}
