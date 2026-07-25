import { bakeReview, colorForId, configureIds, deserializeDocument, freshId, reconstruct, serializeDocument, DEFAULT_CHAR_STYLE, DOCX_MIME, PDF_MIME, PX_PER_INCH, ptToPx as sharedPtToPx, pxToPt as sharedPxToPt, type Change, type DocSelection, type Document, type Fragment, type ParaStyle, type ReviewLayer } from "@cw/shared";
import { resolveConfig } from "./config";
import { emptyParagraphFor } from "./builder/blockFactory";
import { mediaIdsOf, mediaStore, mediaUrl, pruneOrphanMedia, registerMediaBytes, registerMediaRetention, rehydrateDocMedia } from "./media/store";
import { SyncClient } from "./sync/SyncClient";
import { createEditor, type CurrentFormat, type EditMode, type InspectorProbe, type ReviewOpEnvelope } from "./index";
import type { Command } from "./editor/state";
import type { ShapeDash } from "@cw/shared";
import { createLayoutEngine } from "./layout/engine";
import { sampleDoc } from "./model/sampleDoc";
import { stressDoc } from "./model/stressDoc";
import { importDocx, type ImportResult, type ImportPhase } from "./import/docx/importDocx";
import { exportDocument, type ExportFormat, type ExportWarning } from "./export/exportDocument";
import { loadArabicFallbackFont, loadCjkFallbackFont, loadHebrewFallbackFont, loadEditorFonts } from "./export/shared/editorFonts";
import { ARABIC_FONT_FAMILY, CJK_FONT_FAMILY, HEBREW_FONT_FAMILY } from "./fonts/clones";
import { fontsProgress } from "./app/loadProgress";
import { toolbarFonts } from "./fonts/clones";
import { createFontRegistry, normalizeFamily } from "./fonts/customRegistry";
import type { EditorHandle, WordCanvasRuntime } from "./app/runtime";
import { buildShell } from "./app/shell";
import { ensureWordCanvasStyles } from "./ui/styles";
import { showContextMenu, type MenuEntry } from "./ui/contextMenu";
import { createFloatingFormatBar } from "./ui/floatingFormatBar";
import { createContextToolbarManager } from "./ui/contextToolbar";
import { openSurface, type SurfaceHandle } from "./ui/surfaceManager";
import { showInputDialog } from "./ui/inputDialog";
import { showCommandPalette, type PaletteCommand } from "./ui/commandPalette";
import { showShortcutsCheatsheet } from "./ui/shortcutsCheatsheet";
import { createImageContextToolbar } from "./ui/imageContextToolbar";
import { createShapeContextToolbar } from "./ui/shapeContextToolbar";
import { createLinkContextToolbar } from "./ui/linkContextToolbar";
import { createTableContextToolbar } from "./ui/tableContextToolbar";
import { createEquationContextToolbar } from "./ui/equationContextToolbar";
import { createReviewContextToolbar } from "./ui/reviewContextToolbar";
import { createNoteContextToolbar } from "./ui/noteContextToolbar";
import { createTocContextToolbar } from "./ui/tocContextToolbar";
import { createListContextToolbar } from "./ui/listContextToolbar";
import { createInsertMenuToolbar } from "./ui/insertMenuToolbar";
import { createCustomContextToolbar } from "./ui/customContextToolbar";
import type { FloatingToolbarButtonSpec, ToolbarContext } from "./config";
import { showStyleManager, type StyleManagerHandle } from "./ui/styleManager";
import { showDevPanel, type DevPanelHandle } from "./ui/devPanel";
import { showPageLayout, type PageLayoutHandle } from "./ui/pageLayout";
import { showShapeSizePosition, type ShapeSizePositionHandle } from "./ui/shapeSizePosition";
import { showOrganizePages, type OrganizePagesHandle } from "./ui/organizePages";
import { showParagraphDialog, type ParagraphDialogHandle } from "./ui/paragraphDialog";
// The equation editor pulls in the whole LaTeX toolchain (parser + serializer +
// symbol tables, several hundred KB of source) and the symbol picker — both are
// click-driven dialogs, so they load lazily (same pattern as the WebMCP polyfill
// and agent chat) and stay out of the initial editor chunk. Types are erased at
// runtime, so the type-only import below costs nothing.
import type { SymbolPickerHandle } from "./ui/symbolPicker";
import { showFontDialog } from "./ui/fontDialog";
import { loadCollabDocument, loadCollabReview, publishDocument } from "./sync/collab";
import { attachMentionAutocomplete } from "./review/mentions";
import { showBusy } from "./app/busyOverlay";
import { showNotice } from "./app/notice";
import { emitBuilderCode } from "./codegen/emit";
// Ribbon/toolbar command set + style helpers + icons — the chrome, not the core.
// (Hoisted here from the toolbar block; ES imports must be top-level, and the
// toolbar code now lives inside mountEditorApp.)
import { anchorBeforeId, type RibbonActionContext, type RibbonApi, type RibbonButtonSpec } from "./ribbon";
import { chordMatches, resolveCommandBindings, type EditorCommand } from "./commands";
import {
  insertText as insertTextCmd,
  setCustomBlockData as setCustomBlockDataCmd,
  insertImage,
  insertImageInCell,
  setShapeProps,
  bringShapeToFront,
  sendShapeToBack,
  setShapeLayer,
  insertEquation,
  insertInlineEquation,
  insertSymbolCmd,
  insertTable,
  insertTableRowCmd,
  insertTableColumnCmd,
  deleteTableRowCmd,
  deleteTableColumnCmd,
  deleteTableCmd,
  mergeCellsCmd,
  unmergeCellCmd,
  setTableWidthModeAtSelectionCmd,
  setTablePreferredWidthAtSelectionCmd,
  setTableAlignAtSelectionCmd,
  setCellVAlignCmd,
  setCellTextDirectionCmd,
  setRowHeightAtSelectionCmd,
  setRowPropsCmd,
  setTablePropsAtSelectionCmd,
  tableAtSelection,
  setImageProps,
  applyNamedStyle,
  applyCharStyle,
  updateStyleToSelection,
  renameNamedStyle,
  restoreParagraphsCmd,
  styleInstanceRanges,
  hasDirectFormattingAt,
  type ParaSnapshot,
  setParaProps,
  addBookmarkCmd,
  removeBookmarkCmd,
  renameBookmarkCmd,
  toggleList,
  changeListLevel,
  setEquationAlignCmd,
  applyListStyleCmd,
  toggleMultilevelList,
  toggleHighlight,
  toggleVerticalAlign,
  setDirection,
  changeCaseCmd,
  adjustIndentCmd,
  clearCharFormatting,
  type CaseMode,
  setLinkCmd,
  insertPageBreak,
  insertSectionBreak,
  insertTocCmd,
  reorderPageGroupCmd,
  updateTocFieldCmd,
  insertFootnoteCmd,
  insertEndnoteCmd,
  deleteNoteCmd,
  insertContentControl,
  wrapImageInContentControl,
  removeContentControl,
  applyPageSetup,
  pageSetupAt,
} from "./editor/commands";
import { defaultStylesheet, styleById, styleType } from "@cw/shared";
import { bulletListDefinition, numberListDefinition, paragraphsOf, textOfRuns } from "@cw/shared";
import { containerListOf, locateParagraph } from "@cw/shared";
import type { RowProps, TableBlock, TableCell } from "@cw/shared";
import { ICONS } from "./ui/icons";
import { readRecent, pushRecent } from "./ui/recentList";

// Dev hook for in-browser verification (break-rule scans, perf probes, and the
// snapshot/replay round-trip). Assigned near the end of mountEditorApp.
declare global {
  interface Window {
    __cw?: unknown;
  }
}

// Ruler tick GEOMETRY shared by the horizontal ruler (draw) and the vertical
// ruler (drawV), kept here so the two stay pixel-identical. Colors + font are
// per-instance (config.theme.ruler).
const RULER_TICK_START = 4; // band edge the ticks hang from
const RULER_TICK_LEN = 10; // major (inch) tick extent
const RULER_HALF_TICK_LEN = 8; // half-inch tick extent

// Mount ONE editor into runtime.container and hand its control surface to
// runtime.onReady. Called once per WordCanvas instance — every binding below is
// per-call, so multiple editors coexist on a single page. The CSS is class-based
// (see ui/styles.ts) and all off-container artifacts are tracked for destroy().
export async function mountEditorApp(runtime: WordCanvasRuntime): Promise<void> {
  // Resolve config FIRST: custom fonts must be registered (so cloneFamilyFor /
  // metrics know them) and loaded before the first layout. pretext measures with the
  // same font strings the paint layer draws with, so a late font swap would desync
  // them. The editor renders the bundled metric clones (Calibri→Carlito, …) plus any
  // custom fonts, so layout matches the PDF/DOCX exporters exactly.
  const config = resolveConfig({
    ...(runtime.theme ? { theme: runtime.theme } : {}),
    ...(runtime.overrideDefaultStyles ? { overrideDefaultStyles: runtime.overrideDefaultStyles } : {}),
    ...(runtime.behavior ? { behavior: runtime.behavior } : {}),
    ...(runtime.fonts ? { fonts: runtime.fonts } : {}),
    ...(runtime.cjk ? { cjk: runtime.cjk } : {}),
    ...(runtime.develop !== undefined ? { develop: runtime.develop } : {}),
    ...(runtime.organizePages !== undefined ? { organizePages: runtime.organizePages } : {}),
    ...(runtime.chrome !== undefined ? { chrome: runtime.chrome } : {}),
    ...(runtime.floatingToolbar !== undefined ? { floatingToolbar: runtime.floatingToolbar } : {}),
    ...(runtime.contextToolbars !== undefined ? { contextToolbars: runtime.contextToolbars } : {}),
  });
  // This editor instance's own font registry — threaded into its layout engine and
  // paint layer (below) so its custom fonts can't be clobbered by another WordCanvas
  // instance on the page, and into its exports via config.fonts.
  const fontRegistry = createFontRegistry();
  const loadedFamilies = await loadEditorFonts(
    (loaded, total) => runtime.onLoadProgress?.(fontsProgress(loaded, total)),
    config.fonts,
  );
  // Register only families whose required Regular face actually loaded, so a custom
  // font that failed (CORS/404) falls back to a clone consistently in the editor —
  // the same fallback the exporter makes when its own fetch fails. (Layout reads
  // the registry below, so registration must happen before the first engine.layout.)
  fontRegistry.register({ fonts: config.fonts.fonts.filter((f) => loadedFamilies.has(normalizeFamily(f.family))) });
  await document.fonts.ready;
  // CJK fallback + locale are passed to this instance's layout engine (below), which
  // re-asserts them at the top of every layout pass — no process-global juggling.

  // Give this editor session a unique siteId so every block/cell id it mints is
  // disjoint from other collaborating clients' ids (see shared/ids). A short
  // random token is enough; the server later assigns the authoritative ordering.
  configureIds(crypto.randomUUID().slice(0, 8));

  ensureWordCanvasStyles();
  const shell = buildShell(runtime.container);
  const app = shell.app;

  // Config was resolved at the top of mountEditorApp (fonts must register before the
  // first layout). Everything below reads from `config`; omitting an option keeps the
  // built-in look. Theme + behavior thread into the editor via editorOpts; the
  // typography override seeds new/blank docs (a loaded .docx keeps its own).
  // Recolor the gray gutter behind the pages + the ruler troughs (inline overrides
  // the static .cw-app/.cw-ruler stylesheet rules, so it stays per-instance).
  app.style.background = config.theme.canvasBackground;
  shell.ruler.style.background = config.theme.canvasBackground;
  shell.vruler.style.background = config.theme.canvasBackground;

  // Per-instance teardown. Global window/document listeners are registered with
  // this signal, and body-level floating panels (popovers, find bar, image bar,
  // page-setup, the colour <input>) are tracked in `detachables` — destroy()
  // aborts the signal and removes them, so nothing leaks outside shell.root.
  const teardown = new AbortController();
  const detachables: Element[] = [];
  // Container-width compact ribbon observer (assigned in the ribbon block below).
  let compactRO: ResizeObserver | null = null;
  // Responsive layout state (row 12 / critique R1-R4). autoFit = keep the page
  // fit-to-width until the user manually zooms; the outline auto-collapse below
  // narrow widths is remembered so it reopens when there's room again.
  let autoFit = false;
  let applyingFit = false; // true only while WE call setZoom, so onZoomChange can tell manual zoom apart
  let autoCollapsedOutline = false;
  let setOutlineOpen: (v: boolean) => void = () => {};
  let isOutlineOpen: () => boolean = () => false;
  // Navigator (row 17): the bookmarks block renders its list into this panel, and
  // toggleBookmarks switches the rail to the Marks tab, once the rail is built.
  let navMarksPanel: HTMLElement | null = null;
  let navShow: (tab: string) => void = () => {};

  // Dark-mode chrome (critique V2): reflect the OS colour scheme onto
  // :root[data-theme], which the dark stylesheet keys off. Marks its own writes
  // with data-theme-auto so a host that sets data-theme itself keeps control.
  if (typeof matchMedia === "function") {
    const de = document.documentElement;
    const mq = matchMedia("(prefers-color-scheme: dark)");
    const applyTheme = (): void => {
      if (de.dataset["theme"] && de.dataset["themeAuto"] !== "1") return; // host-managed
      de.dataset["theme"] = mq.matches ? "dark" : "light";
      de.dataset["themeAuto"] = "1";
    };
    applyTheme();
    mq.addEventListener("change", applyTheme, { signal: teardown.signal });
  }

  // Ctrl+/ — keyboard shortcuts cheat sheet (critique A3). Embedder commands with
  // a keybinding are appended live, so the sheet reflects the actual keymap.
  window.addEventListener(
    "keydown",
    (e) => {
      if (!(e.ctrlKey || e.metaKey) || e.altKey || e.key !== "/") return;
      e.preventDefault();
      const cmds = (runtime.commands ?? [])
        .filter((c) => c.keybinding && c.label)
        .map((c) => ({ chord: (Array.isArray(c.keybinding) ? c.keybinding[0] : c.keybinding) as string, label: c.label as string }));
      showShortcutsCheatsheet(cmds.length ? [{ title: "Commands", items: cmds }] : []);
    },
    { signal: teardown.signal },
  );

// Backend base URL gates online features. Present ⇒ sync/publish/share; absent ⇒
// fully offline editor (the kill-switch).
const BACKEND_HTTP: string | null = runtime.backendUrl ?? null;
const BACKEND_WS: string | null = BACKEND_HTTP ? BACKEND_HTTP.replace(/^http/, "ws") : null;
const online = BACKEND_HTTP !== null;
// View-only viewer: hide the editing chrome (ribbon, ruler, replace) and run the
// editor core in readonly mode. Selection, copy, find, zoom, and live remote
// edits all still work.
const readonly = runtime.readonly ?? false;
// Join an existing session only makes sense online.
let collabId: string | null = online ? (runtime.collabId ?? null) : null;
let shareLink: string | null = null;

const reportImport = (r: ImportResult, ms: number): void => {
  console.log(`[docx-import] blocks=${r.doc.blocks.length} warnings=${r.warnings.length} total=${ms.toFixed(1)}ms`);
  for (const w of r.warnings) console.warn(`[docx-import] ${w.code}: ${w.message}`);
};

// Human labels for the import worker's progress phases (shown in the busy overlay).
const IMPORT_PHASE_LABEL: Record<ImportPhase, string> = {
  unzip: "Unzipping",
  styles: "Reading styles",
  parse: "Parsing content",
  map: "Building document",
};

let collabVersion = 0;
let sync: SyncClient | null = null;
let doc: Document;
if (collabId !== null && BACKEND_HTTP) {
  const busy = showBusy("Loading shared document…");
  try {
    const loaded = await loadCollabDocument(BACKEND_HTTP, collabId);
    collabVersion = loaded.version;
    doc = loaded.doc;
    // We joined an existing session: the URL in the address bar IS the share
    // link, so the Share button reveals it instead of re-publishing.
    shareLink = location.href;
    console.log(`[collab] loaded ${collabId} at v${collabVersion}`);
  } catch (e) {
    // Dead/invalid link — don't hang the whole load. Drop collab state, strip
    // ?collab (so a reload doesn't re-fail), and fall back to a blank editor.
    console.error("[collab] load failed", e);
    showNotice(`Could not open the shared document: ${e instanceof Error ? e.message : String(e)}`, "warning");
    collabId = null;
    shareLink = null;
    try {
      const u = new URL(location.href);
      u.searchParams.delete("collab");
      history.replaceState(null, "", u.toString());
    } catch {
      /* ignore URL errors */
    }
    doc = sampleDoc();
  } finally {
    busy.done();
  }
} else {
  doc = sampleDoc();
}

// A document with no stylesheet of its own adopts the configured default styles
// (the `overrideDefaultStyles` typography, or the library default). A loaded .docx
// always arrives WITH its own stylesheet, so its w:docDefaults/Normal win.
if (!doc.stylesheet) doc = { ...doc, stylesheet: config.stylesheet };

const engine = createLayoutEngine(fontRegistry, {
  ...(config.cjk.fallbackFont !== undefined ? { cjkFallback: config.cjk.fallbackFont } : {}),
  ...(config.cjk.locale !== undefined ? { cjkLocale: config.cjk.locale } : {}),
  ...(config.cjk.arabicFallbackFont !== undefined ? { arabicFallback: config.cjk.arabicFallbackFont } : {}),
  ...(config.cjk.hebrewFallbackFont !== undefined ? { hebrewFallback: config.cjk.hebrewFallbackFont } : {}),
});
const t0 = performance.now();
const tree = engine.layout(doc); // cold: every paragraph hits prepareRichInline
const t1 = performance.now();
engine.layout(doc); // warm: 100% prepare-cache hits, pure line-walk + pagination
const t2 = performance.now();
console.log(
  `[canvas-word] blocks=${doc.blocks.length} pages=${tree.pages.length} ` +
    `layout cold=${(t1 - t0).toFixed(1)}ms warm=${(t2 - t1).toFixed(1)}ms`,
);

let syncToolbar: () => void = () => {};
let syncZoom: (zoom: number) => void = () => {};
let refreshOutline: () => void = () => {};
let refreshStatus: () => void = () => {};
// Document-identity / save-state chrome (top-left of the ribbon). Forward-
// declared like the refreshers above so replaceDocument() / onChange (defined
// before the toolbar block builds the widget) can drive them; they stay no-ops
// when the ribbon is disabled (chromeless mount).
let refreshSaveState: () => void = () => {};
// Import fidelity (Move 3): the lossy/adapted-mapping warnings from the last .docx
// open, surfaced as a passive chrome badge — but ONLY in its warning state. `null`
// means no import has happened yet (the in-memory sample, a blank doc); an array
// (possibly empty) means an import completed. The badge renders nothing for `null`
// or an empty array (a clean import is silent by design — a permanent "faithful"
// check would be unearned chrome and would train the eye to ignore this region);
// it appears only when there are warnings to show. Set in openDocxFile; read via
// refreshFidelity.
let importWarnings: ImportResult["warnings"] | null = null;
let refreshFidelity: () => void = () => {};
let syncQat: () => void = () => {}; // enable/disable the quick-access undo/redo
let resetSaveBaseline: () => void = () => {};
let setDocTitleFromFile: (name: string | null) => void = () => {};
// Filesystem-safe base name for a downloaded export, derived from the document
// title (assigned in the toolbar block; falls back to "document" when chromeless).
let currentDocBaseName: () => string = () => "document";
let refreshRuler: () => void = () => {};
let toggleRuler: () => boolean = () => false;
let refreshVRuler: () => void = () => {};
let toggleVRuler: () => boolean = () => false;
let refreshContextToolbars: () => void = () => {};
// rAF-coalesced variant for the drag-frequency selection hook (see onSelectionChange).
let scheduleRefreshContextToolbars: () => void = () => {};
let openLinkDialog: (anchor: HTMLElement, url?: string) => void = () => {};
let refreshBookmarks: () => void = () => {};
let toggleBookmarks: () => void = () => {};
let refreshReview: () => void = () => {};
let toggleReview: () => void = () => {};
// Inspector (Move 2): a right-docked, selection-aware property sheet. Forward-
// declared so the header button and the selection/change hooks can drive it.
let toggleInspector: () => void = () => {};
let refreshInspector: () => void = () => {};
// Minimal-chrome command bar (Move 1): reflects the caret's bold/italic/list/style
// state on its compact buttons. No-op unless chrome === "minimal".
let refreshMinimalBar: () => void = () => {};
let syncMode: () => void = () => {};
// Develop-mode Document-tree inspector hooks (assigned when the panel is open;
// no-ops otherwise, so non-develop mounts pay nothing). refreshDevPanel re-reads
// the tree on edits; inspectorHoverSink routes the canvas→tree hover signal.
let refreshDevPanel: () => void = () => {};
let inspectorHoverSink: (blockId: string | null) => void = () => {};
let inspectorProbeSink: (probe: InspectorProbe | null) => void = () => {};
// Close the inspector (it captures the live editor) when the editor is rebuilt on
// document replacement. Assigned when the panel opens; no-op otherwise.
let closeDevPanel: () => void = () => {};
// Close the Paragraph dialog when the editor is rebuilt on document replacement —
// an orphaned dialog holds stale paragraph state and would dispatch its OK patch
// into the newly created editor. Assigned when the dialog opens; no-op otherwise.
let closeParaDialog: () => void = () => {};
// Ships a locally-authored review op on the collab review channel (assigned by
// the sync wiring; no-op offline). Forward-declared so editorOpts can reference it.
let onLocalReviewOp: (env: ReviewOpEnvelope) => void = () => {};
// The action context handed to custom ribbon buttons' onClick (handle + macro
// helpers). Assigned once the editor handle is built; custom buttons can't fire
// before mount completes, so it's always set by click time.
let ribbonCtx: RibbonActionContext | null = null;
// Registered custom commands (the `commands` option), by id. Populated once the
// editor handle + ribbonCtx are built; used by handle.runCommand and the custom
// keybinding listener.
const commandRegistry = new Map<string, EditorCommand>();
const runRegisteredCommand = (id: string): boolean => {
  const cmd = commandRegistry.get(id);
  if (!cmd || !ribbonCtx) return false;
  try {
    // Commands may be async: catch a synchronous throw here AND a rejected
    // promise below, so neither becomes an unhandled rejection.
    const r = cmd.run(ribbonCtx);
    if (r && typeof (r as Promise<void>).then === "function") {
      void (r as Promise<void>).catch((err) => console.error(`[canvas-word] command "${id}" run() rejected`, err));
    }
  } catch (err) {
    console.error(`[canvas-word] command "${id}" run() threw`, err);
  }
  return true;
};
// Opens the Size & Position dialog for the selected shape. Assigned once the ribbon
// machinery is set up (below); the right-click menu reaches it through the editor
// option hook so both surfaces share one implementation.
let openShapeSizePosition: () => void = () => {};
const editorOpts = {
  engine,
  paintOptions: { fontRegistry },
  readonly,
  theme: config.theme,
  behavior: config.behavior,
  // Right-click "Size & Position…" on a selected shape routes here (B2).
  onEditShapeSizePosition: () => openShapeSizePosition(),
  ...(runtime.mode ? { mode: runtime.mode } : {}),
  ...(runtime.allowedModes ? { allowedModes: runtime.allowedModes } : {}),
  ...(runtime.user ? { user: runtime.user } : {}),
  ...(runtime.knownUsers ? { knownUsers: runtime.knownUsers } : {}),
  ...(runtime.resolveField ? { resolveField: runtime.resolveField } : {}),
  ...(collabId !== null ? { docId: collabId } : {}),
  // Right-click "Fill…" / "Outline…" for a selected shape open the SAME shared
  // colour popovers the ribbon and floating toolbar use, anchored to the shape's
  // on-screen rect (issue #244 C1).
  onShapeFillMenu: () => {
    const r = editor.getSelectedShapeRect();
    if (r) shapeFillPopover(r);
  },
  onShapeOutlineMenu: () => {
    const r = editor.getSelectedShapeRect();
    if (r) shapeOutlinePopover(r);
  },
  // In a collab session, ship each recorded local edit to the server.
  onChangeRecorded: (change: Change) => {
    sync?.localEdit(change.ops);
  },
  // Review overlay: re-emit changes to the embedder + ship review ops to peers.
  onReviewChanged: (review: ReviewLayer) => {
    runtime.onEvent?.({ type: "reviewChanged", review });
    refreshReview();
  },
  onModeChanged: (mode: EditMode) => {
    runtime.onEvent?.({ type: "modeChanged", mode });
    syncMode();
  },
  onReviewOpRecorded: (env: ReviewOpEnvelope) => onLocalReviewOp(env),
  // Broadcast the local caret to collaborators on every move.
  onSelectionChange: (sel: DocSelection | null) => {
    sync?.localPresence(sel);
    // Cell (multi-row/col) selection updates through this hook only — its setter
    // skips the heavy onChange broadcast (it fires per drag-move) — so refresh the
    // context toolbars here too, else the table bar wouldn't appear until a scroll.
    // Coalesced (rAF) because this fires continuously during a drag and refresh reads
    // layout; onChange stays synchronous for typing/caret immediacy.
    scheduleRefreshContextToolbars();
    refreshInspector();
    refreshMinimalBar();
    refreshDevPanel();
  },
  // Develop mode only: the inspector turns this signal on while open (dormant
  // otherwise) — route the hovered block id to the tree for the reverse highlight.
  onInspectorHover: (blockId: string | null) => inspectorHoverSink(blockId),
  onInspectorProbe: (probe: InspectorProbe | null) => inspectorProbeSink(probe),
  // Right-click "Insert/Edit Hyperlink" routes here so it opens the SAME styled
  // link popover the ribbon uses (critique 2.8: no more native prompt(), and one
  // link UI instead of two). Anchored at the caret via a transient element.
  onEditLink: (current: string | null): void => {
    const r = editor.getSelectionAnchorRect();
    const anchor = document.createElement("div");
    anchor.style.cssText = `position:fixed;left:${Math.round(r?.left ?? 120)}px;top:${Math.round(r?.top ?? 90)}px;` +
      `width:${Math.round(r?.width ?? 0)}px;height:${Math.round(r?.height ?? 16)}px;pointer-events:none;`;
    document.body.appendChild(anchor);
    openLinkDialog(anchor, current ?? ""); // openPop reads the rect synchronously
    anchor.remove();
  },
  onChange: () => {
    syncToolbar();
    refreshOutline();
    refreshStatus();
    refreshSaveState();
    refreshRuler();
    refreshVRuler();
    refreshContextToolbars();
    refreshBookmarks();
    refreshInspector();
    refreshMinimalBar();
    refreshDevPanel();
  },
  onZoomChange: (z: number) => {
    if (!applyingFit) autoFit = false; // a manual zoom (controls / Ctrl+wheel) exits fit-width mode
    syncZoom(z);
    refreshStatus();
    refreshRuler();
    refreshVRuler();
    refreshContextToolbars();
  },
};
let editor = createEditor(app, doc, editorOpts);

// Media retention: declare which mediaIds this mount's CURRENT document
// references (read live at prune time — `doc` is reassigned on replacement).
// pruneOrphanMedia() (called on document replacement) evicts the rest.
const unregisterMediaRetention = registerMediaRetention(() => mediaIdsOf(doc));

// When the bundled CJK fallback is active, load it lazily (it's multi-MB; blocking
// first paint on it would penalize Latin-only documents) and re-lay-out once it
// arrives so CJK widths re-measure against the bundled face instead of the browser's
// interim system substitute. A custom fallback family is already loaded eagerly by
// loadEditorFonts; an explicit "" opt-out skips this entirely.
if (config.cjk.fallbackFont === CJK_FONT_FAMILY) {
  void loadCjkFallbackFont().then((ok) => {
    // The fetch is uncancellable; if the editor was destroyed while it was in
    // flight, skip — refreshFonts() would touch a torn-down editor/root.
    if (!ok || teardown.signal.aborted) return;
    editor.refreshFonts();
  });
}

// Same lazy-load pattern for the bundled Arabic fallback. Arabic text needs
// contextual joining (GSUB) that browser system fallbacks rarely supply correctly
// in canvas; once the bundled face loads, re-layout re-measures Arabic runs against
// it so the screen matches the PDF export exactly.
if (config.cjk.arabicFallbackFont === ARABIC_FONT_FAMILY) {
  void loadArabicFallbackFont().then((ok) => {
    if (!ok || teardown.signal.aborted) return;
    editor.refreshFonts();
  });
}

// Same lazy-load pattern for the bundled Hebrew fallback; once it loads, re-layout
// re-measures Hebrew runs against it so the screen matches the PDF export.
if (config.cjk.hebrewFallbackFont === HEBREW_FONT_FAMILY) {
  void loadHebrewFallbackFont().then((ok) => {
    if (!ok || teardown.signal.aborted) return;
    editor.refreshFonts();
  });
}

// Drawing-grid view state (persisted across document replacement, since the
// editor — and its paint layer — is rebuilt on docx open / setDocument). Seeded
// from the embedder's `view` options. Ruler visibility is DOM-class state on the
// shell, which survives a rebuild, so it isn't tracked here.
const gridView = {
  show: runtime.view?.grid ?? false,
  snap: runtime.view?.snapToGrid ?? false,
  spacingPx: runtime.view?.gridSpacingPx ?? config.behavior.gridSpacingPx,
};
// Formatting-marks visibility — same "survives a paint-layer rebuild" treatment
// as the grid (re-seeded via applyGridView below).
let showFormattingMarks = runtime.view?.formattingMarks ?? false;
const applyGridView = (): void => {
  editor.setGridSpacing(gridView.spacingPx);
  editor.setShowGrid(gridView.show);
  editor.setSnapToGrid(gridView.snap);
  editor.setShowFormattingMarks(showFormattingMarks);
};
applyGridView();

// Initial ruler visibility (both default on; always hidden in readonly, whose
// only ruler interaction — indent drag — is an edit). Shared by the ruler blocks
// and the View-tab toggle buttons below.
const showHRuler = !readonly && (runtime.view?.ruler ?? true);
const showVRuler = !readonly && (runtime.view?.verticalRuler ?? true);

// Initial zoom (what the reader opens at). One-time at mount — later docx opens
// reset to 100%, like a fresh document.
if (runtime.view?.zoom != null) editor.setZoom(runtime.view.zoom);

// Listeners notified whenever a remote change is applied (the activity panel,
// Part 3, registers here).
const remoteChangeListeners: Array<(change: Change) => void> = [];

// Build a SyncClient for `docId` wired to the current editor + identity + the
// presence/event callbacks. Used by both the join path and goOnline.
const connectSync = (docId: string, startVersion: number): SyncClient => {
  const s = new SyncClient({
    wsUrl: BACKEND_WS!,
    docId,
    editor,
    startVersion,
    user: runtime.user,
    onUserEntered: (siteId, user) => runtime.onEvent?.({ type: "userEntered", siteId, user }),
    onUserLeave: (siteId, user) => runtime.onEvent?.({ type: "userLeave", siteId, user }),
    onPresence: (participants) => runtime.onEvent?.({ type: "presence", participants }),
    onRemote: (change: Change) => {
      for (const l of remoteChangeListeners) l(change);
    },
  });
  s.connect();
  // Ship locally-authored review ops on this socket's review channel.
  onLocalReviewOp = (env) => s.localReviewOp(env);
  return s;
};

// Race-safe review rehydration for a JOINED session: hold inbound review ops,
// wait until the socket is in the room, fetch the rehydrated snapshot, seed it,
// then replay the held ops (idempotent) so a review op landing during the
// HTTP-load → ws-handshake window isn't lost. REVIEW.md §6.
const rehydrateReviewOnJoin = (s: SyncClient, docId: string): void => {
  if (!BACKEND_HTTP) return;
  s.holdReview();
  void (async () => {
    await s.whenOpen().catch(() => {});
    try {
      editor.seedReview(await loadCollabReview(BACKEND_HTTP, docId));
    } catch (e) {
      console.error("[collab] review rehydrate failed", e);
    } finally {
      s.markReviewReady(); // always release the hold, even if the fetch failed
    }
  })();
};

if (collabId !== null && BACKEND_WS) {
  sync = connectSync(collabId, collabVersion);
  rehydrateReviewOnJoin(sync, collabId);
  console.log(`[collab] connected to ${BACKEND_WS}/ws?doc=${collabId}`);
}

// Assigned by the toolbar block; shows the built-in share dialog (used when the
// embedder didn't supply onShareLink).
let showShareDialog: (link: string) => void = () => {};
// Assigned by the activity panel block; toggles the who/when/what-kind drawer.
let toggleActivity: () => void = () => {};

// Hand a share link to the embedder's handler, or fall back to the built-in dialog.
const surfaceShareLink = (link: string, docId: string): void => {
  if (runtime.onShareLink) runtime.onShareLink(link, docId);
  else showShareDialog(link);
};

// Publish the current document ONCE and switch this editor into a live collab
// session, reflecting it in the URL and surfacing a share link. Online only.
//
// Idempotent: once published this session (or joined via a ?collab link, which
// pre-sets shareLink), later calls just re-surface the cached link — no second
// POST /docs, no WebSocket teardown. Live edits already converge through the
// sync socket, so the server snapshot never needs re-posting. Used by the Share
// button and the public share() API.
const goOnlineWithCurrentDoc = async (): Promise<string> => {
  if (!BACKEND_HTTP || !BACKEND_WS) throw new Error("cannot share: offline (no backendUrl)");
  // Already shared — reuse the existing link instead of forking a new document.
  if (collabId && shareLink) {
    surfaceShareLink(shareLink, collabId);
    return shareLink;
  }
  const busy = showBusy("Creating share link…");
  let docId: string;
  try {
    ({ docId } = await publishDocument(BACKEND_HTTP, editor.getDocument(), runtime.user));
  } finally {
    busy.done();
  }
  sync?.destroy();
  collabId = docId;
  collabVersion = 0;
  sync = connectSync(docId, 0);
  // Reflect in the address bar so a reload re-joins (best-effort; harmless if the
  // host controls routing differently).
  try {
    const u = new URL(location.href);
    u.searchParams.set("collab", docId);
    history.replaceState(null, "", u.toString());
    shareLink = u.toString();
  } catch {
    shareLink = `${location.origin}${location.pathname}?collab=${docId}`;
  }
  console.log(`[collab] published, sharing ${docId}`);
  runtime.onEvent?.({ type: "shared", docId, url: shareLink });
  surfaceShareLink(shareLink, docId);
  return shareLink;
};

// Replace the open document (docx import): tear down and rebuild the editor.
// The layout engine instance is reused (cheap to keep its allocations), but its
// caches MUST be dropped first — see engine.reset() below.
const replaceDocument = (next: typeof doc): void => {
  closeDevPanel(); // the inspector captures the outgoing editor — drop it first
  closeParaDialog(); // ditto for the Paragraph dialog (stale editor reference)
  editor.destroy();
  // Drop the outgoing document's layout caches. The shared engine keys cached
  // lines by (block id, revision, width); the docx importer re-mints ids from i0
  // at revision 0, so the new document collides with the old one and would render
  // the two merged unless we clear first.
  engine.reset();
  doc = next;
  // The outgoing document's media (bytes + blob: URLs) would otherwise stay
  // resident for the rest of the session — the editor was just rebuilt (fresh
  // undo history), so this is the undo-safe moment to evict what the incoming
  // document doesn't reference.
  pruneOrphanMedia();
  editor = createEditor(app, doc, editorOpts);
  applyGridView(); // re-seed grid state onto the fresh paint layer
  refreshOutline();
  refreshStatus();
  refreshRuler();
  refreshVRuler();
  resetSaveBaseline(); // fresh editor ⇒ change head is back at 0; re-baseline "clean"
  window.__cw = { doc, tree: undefined, engine, editor, createLayoutEngine, sampleDoc, stressDoc, persist };
};

// A freshly opened/set document is a NEW document: drop any live collab session
// so the next Share publishes a fresh doc/link (fork), and clear ?collab.
// Publishing is lazy — it happens on the first Share click, not on open.
const detachCollabSession = (): void => {
  sync?.destroy();
  sync = null;
  collabId = null;
  shareLink = null;
  collabVersion = 0;
  try {
    const u = new URL(location.href);
    if (u.searchParams.has("collab")) {
      u.searchParams.delete("collab");
      history.replaceState(null, "", u.toString());
    }
  } catch {
    /* ignore URL errors */
  }
};

// Programmatic document swap (the builder / data-binding path): same teardown
// as a docx open, but the caller hands a ready Document. Clone defensively so
// later caller-side mutations (or a re-used builder) can't alias editor state,
// and preserve the viewport (zoom + scroll) so live rebuilds don't jump.
const setDocumentFromApi = (next: Document): void => {
  const clone = structuredClone(next);
  // A doc with no stylesheet adopts the configured defaults before we mint any
  // empty paragraph from it (a real .docx always brings its own — see config).
  if (!clone.stylesheet) clone.stylesheet = config.stylesheet;
  if (clone.blocks.length === 0) clone.blocks.push(emptyParagraphFor(clone, freshId()));
  const scrollTop = app.scrollTop;
  const zoom = editor.getZoom();
  replaceDocument(clone);
  editor.setZoom(zoom);
  requestAnimationFrame(() => {
    app.scrollTop = Math.min(scrollTop, Math.max(0, app.scrollHeight - app.clientHeight));
  });
  detachCollabSession();
};

const openDocxFile = async (file: File | ArrayBuffer): Promise<void> => {
  const i0 = performance.now();
  const busy = showBusy("Opening document…");
  try {
    const result = await importDocx(file, {
      onProgress: (phase, pct) => busy.update(`${IMPORT_PHASE_LABEL[phase]}…`, pct),
    });
    reportImport(result, performance.now() - i0);
    replaceDocument(result.doc);
    importWarnings = result.warnings; // drive the fidelity badge (Move 3)
    refreshFidelity();
    setDocTitleFromFile(file instanceof File ? file.name : null); // reflect the opened file's name
    detachCollabSession();
  } catch (e) {
    console.error("[docx-import]", e);
    const name = file instanceof File ? file.name : "document";
    showNotice(`Could not open "${name}": ${e instanceof Error ? e.message : String(e)}`, "warning");
  } finally {
    busy.done();
  }
};

// Bake the review overlay into a clean doc (default: reject-all = original
// baseline) so the review-blind writer emits a plain file, then export it.
// Defined at mount scope so both the toolbar's Export buttons and the public
// handle's exportDocx()/exportPdf() share one path.
const exportBaked = (format: ExportFormat): Promise<{ bytes: Uint8Array; warnings: ExportWarning[] }> => {
  const baked = bakeReview(editor.getDocument(), editor.getReview(), "reject");
  return exportDocument(baked, format, config.fonts, config.cjk);
};
const blobFor = (format: ExportFormat, bytes: Uint8Array): Blob =>
  new Blob([bytes as BlobPart], { type: format === "pdf" ? PDF_MIME : DOCX_MIME });
/** Export to a Blob — backs the public handle's exportDocx()/exportPdf(). */
const exportBlob = async (format: ExportFormat): Promise<Blob> => {
  const { bytes } = await exportBaked(format);
  return blobFor(format, bytes);
};

// ---- demo toolbar (app chrome, not editor core) ----------------------------

// One function assigned by the find-bar block below; the ribbon's Find button
// and the global Ctrl+F shortcut share it.
let openFind: () => void = () => {};

// View-only (or an embedder that opts out via view.toolbar): visually hide the
// editing ribbon, but keep the block running so the side drawers it wires up
// (outline navigation, bookmarks) still build and show.
const toolbar = shell.toolbar;
const showToolbar = !readonly && (runtime.view?.toolbar ?? true);
if (!showToolbar) toolbar.style.display = "none";
// Popover anchor: an element to anchor under (ribbon / floating-toolbar button) or a
// bare viewport rect (the right-click menu path, where the shape's on-screen rect is
// the anchor — the menu row is already gone by the time the action runs).
type PopoverAnchor = HTMLElement | { left: number; top: number; width: number; height: number };
// Drawing-shape fill/outline pickers: defined inside the ribbon block (they reuse
// its colour-popover machinery) but also called from the floating shape context
// toolbar registered further below, so their handles are hoisted here. No-ops until
// the ribbon block assigns them (a doc with no ribbon has no popover chrome anyway).
let shapeFillPopover: (anchor: PopoverAnchor) => void = () => {};
let shapeOutlinePopover: (anchor: PopoverAnchor) => void = () => {};
if (toolbar) {
  // ===== Word-style tabbed ribbon ==========================================
  // A tab strip drives a stack of panels (one visible at a time). Each panel
  // holds labeled groups of controls. The Home tab mirrors Word's layout; the
  // app's other real commands live on Insert / Layout / Table / View. Features
  // the engine/renderer can't do yet render as disabled stubs (greyed, tooltip).
  const CARET =
    `<svg class="caret" viewBox="0 0 8 8" fill="none" stroke="currentColor" ` +
    `stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"><path d="M1.5 3 4 5.5 6.5 3"/></svg>`;
  const chevron = (up: boolean): string =>
    `<svg width="12" height="12" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.4" ` +
    `stroke-linecap="round" stroke-linejoin="round"><path d="${up ? "M3 7.5 6 4.5 9 7.5" : "M3 4.5 6 7.5 9 4.5"}"/></svg>`;
  const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string): HTMLElementTagNameMap[K] => {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    return e;
  };

  const tabsBar = el("div", "rib-tabs");
  // Tab buttons live in an inner scroll wrapper so they scroll horizontally on
  // overflow while the right-side cluster (mode select / Review / collapse
  // chevron) stays pinned as direct children of `.rib-tabs`.
  const tabScroll = el("div", "rib-tab-scroll");
  tabsBar.appendChild(tabScroll);
  // Mouse users on a narrow desktop: translate vertical wheel to horizontal scroll
  // (inert when the tabs fit). Auto-removed via the teardown signal.
  tabScroll.addEventListener(
    "wheel",
    (e) => {
      if (e.deltaY === 0 || tabScroll.scrollWidth <= tabScroll.clientWidth) return;
      e.preventDefault();
      tabScroll.scrollLeft += e.deltaY;
    },
    { passive: false, signal: teardown.signal },
  );
  const bodies = el("div", "rib-bodies");
  toolbar.append(tabsBar, bodies);
  const tabButtons = new Map<string, HTMLButtonElement>();
  const tabPanels = new Map<string, HTMLDivElement>();
  // Contextual tabs (Table / Picture / Shape Tools): hidden until their predicate
  // matches the current selection, then revealed — Word's contextual-tab model.
  // Replaces a permanent Table tab + a graveyard of always-disabled shape/image
  // buttons on Insert (critique C5). Evaluated in syncToolbar.
  const contextualTabs: { btn: HTMLButtonElement; pred: (f: CurrentFormat) => boolean }[] = [];
  const showTab = (id: string): void => {
    for (const [tid, b] of tabButtons) b.classList.toggle("active", tid === id);
    for (const [tid, p] of tabPanels) p.classList.toggle("active", tid === id);
    setCollapsed(false); // clicking a tab re-pins a collapsed ribbon (Word)
  };
  const tab = (id: string, label: string, kind?: "file", contextual?: (f: CurrentFormat) => boolean): HTMLDivElement => {
    const t = el("button", "rib-tab" + (kind === "file" ? " file" : "") + (contextual ? " rib-tab-ctx" : ""));
    t.textContent = label;
    t.dataset["ribbonTab"] = id;
    t.addEventListener("click", () => showTab(id));
    tabButtons.set(id, t);
    if (contextual) {
      t.style.display = "none"; // revealed by syncToolbar when the predicate matches
      contextualTabs.push({ btn: t, pred: contextual });
    }
    tabScroll.appendChild(t);
    const p = el("div", "rib-panel");
    p.dataset["tab"] = id;
    p.dataset["ribbonPanel"] = id;
    tabPanels.set(id, p);
    bodies.appendChild(p);
    ribTabsById.set(id, { btn: t, panel: p });
    curTabId = id;
    return p;
  };

  // `controls` is the container the btn/select helpers append into; group() and
  // row() retarget it as the ribbon is built.
  let controls: HTMLElement = el("div");
  // Begin a group: derive its namespaced id, register it, and make it current so
  // subsequent item helpers register under it. `container` is where custom-added
  // buttons land (the group's controls, or the last row for two-row groups).
  const beginGroup = (panel: HTMLElement, label: string, container: HTMLElement, groupEl: HTMLElement): void => {
    const tabId = (panel as HTMLElement).dataset["ribbonPanel"] ?? curTabId;
    curGroupId = ribUniqueId(`${tabId}.${ribSlug(label)}`, ribGroupsById);
    groupEl.dataset["ribbonGroup"] = curGroupId;
    ribGroupsById.set(curGroupId, { tabId, groupEl, container });
  };
  const group = (panel: HTMLElement, label: string): void => {
    const g = el("div", "rib-group");
    controls = el("div", "rib-controls");
    const l = el("div", "rib-label");
    l.textContent = label;
    g.append(controls, l);
    panel.appendChild(g);
    beginGroup(panel, label, controls, g);
  };
  /** A group whose controls stack in two rows (Font, Paragraph). Returns a
   *  `row()` that opens a fresh row and points the helpers at it. */
  const groupRows = (panel: HTMLElement, label: string): (() => void) => {
    const g = el("div", "rib-group");
    const rows = el("div", "rib-rows");
    const l = el("div", "rib-label");
    l.textContent = label;
    g.append(rows, l);
    panel.appendChild(g);
    beginGroup(panel, label, rows, g);
    const gid = curGroupId;
    return () => {
      controls = el("div", "rib-row");
      rows.appendChild(controls);
      const gnode = ribGroupsById.get(gid);
      if (gnode) gnode.container = controls; // custom adds land in the last row
    };
  };
  const sep = (): void => {
    const d = el("div");
    d.style.cssText = "width:1px;height:18px;background:#e1dfdd;margin:0 3px;flex:0 0 auto;";
    controls.appendChild(d);
    ribRegister(d, "sep");
  };

  const btn = (icon: string, title: string, onClick: () => void, caret = false): HTMLButtonElement => {
    const b = el("button", "rib-btn");
    b.innerHTML = icon + (caret ? CARET : "");
    b.title = title;
    b.setAttribute("aria-label", title); // explicit, unique accessible name (C6/A2)
    b.addEventListener("mousedown", (e) => e.preventDefault()); // keep IME-proxy focus
    b.addEventListener("click", onClick);
    controls.appendChild(b);
    ribRegister(b, title);
    return b;
  };
  /** Letter buttons (B, I, U, x²…) — Word renders these as styled text. */
  const txtBtn = (label: string, title: string, onClick: () => void, style = "", caret = false): HTMLButtonElement => {
    const b = el("button", "rib-btn");
    b.title = title;
    // The visible glyph ("A", "AB", "LTR"…) is decorative — the accessible name
    // is the descriptive title, so no two buttons announce as the same "A" (C6).
    b.setAttribute("aria-label", title);
    const s = el("span");
    s.setAttribute("aria-hidden", "true");
    s.textContent = label;
    if (style) s.style.cssText = style;
    b.appendChild(s);
    if (caret) b.insertAdjacentHTML("beforeend", CARET);
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", onClick);
    controls.appendChild(b);
    ribRegister(b, title);
    return b;
  };
  /** Large Paste-style button: icon over a caption (with optional caret). */
  const bigBtn = (icon: string, caption: string, title: string, onClick: () => void, caret = false): HTMLButtonElement => {
    const b = el("button", "rib-btn rib-big");
    b.title = title;
    b.setAttribute("aria-label", title); // unique accessible name (C6/A2)
    b.innerHTML = icon + `<span class="big-cap">${caption}${caret ? CARET : ""}</span>`;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", onClick);
    controls.appendChild(b);
    ribRegister(b, title);
    return b;
  };
  // Formatting-marks toggle — surfaced in both Home (Paragraph) and View, so the
  // buttons share one state and reflect it in lockstep (like Word's ¶ button).
  const marksBtns: HTMLButtonElement[] = [];
  const marksToggleBtn = (): HTMLButtonElement => {
    const b = btn(ICONS.marks, "Show/hide formatting marks (spaces, tabs, paragraph ends, line breaks)", () => {
      showFormattingMarks = !showFormattingMarks;
      editor.setShowFormattingMarks(showFormattingMarks);
      for (const m of marksBtns) m.classList.toggle("active", showFormattingMarks);
    });
    b.classList.toggle("active", showFormattingMarks);
    marksBtns.push(b);
    return b;
  };
  const select = (title: string, width: number): HTMLSelectElement => {
    const s = el("select");
    s.title = title;
    s.style.width = `${width}px`;
    s.addEventListener("mousedown", (e) => e.stopPropagation());
    controls.appendChild(s);
    ribRegister(s, title);
    return s;
  };
  const opt = (s: HTMLSelectElement, value: string, label: string): void => {
    const o = el("option");
    o.value = value;
    o.textContent = label;
    s.appendChild(o);
  };

  // Toggle buttons paint a pressed state from the caret's format (Word). Each
  // registers a predicate; syncToolbar re-evaluates them on every change.
  const toggleButtons: { el: HTMLButtonElement; active: (f: CurrentFormat) => boolean }[] = [];
  const toggle = (b: HTMLButtonElement, active: (f: CurrentFormat) => boolean): HTMLButtonElement => {
    toggleButtons.push({ el: b, active });
    return b;
  };

  // Context-gated buttons: a button is disabled (greyed) when its predicate is
  // false for the caret's context — e.g. image buttons need a selected image, so
  // they don't silently no-op when the caret is in a paragraph. syncToolbar
  // re-evaluates these on every change.
  const enableButtons: { el: HTMLButtonElement; enabled: (f: CurrentFormat) => boolean; title: string; hint?: string }[] = [];
  const enable = (b: HTMLButtonElement, enabled: (f: CurrentFormat) => boolean, hint?: string): HTMLButtonElement => {
    enableButtons.push({ el: b, enabled, title: b.title, ...(hint !== undefined ? { hint } : {}) });
    return b;
  };

  // ---- ribbon customization model (stable ids for customizeRibbon) --------
  // Every built-in tab/group/control registers under an auto-derived, namespaced
  // id ("home", "home.font", "home.font.bold") — slugged from its tooltip/label
  // and de-duped within a group. The `customizeRibbon` hook reorders/removes
  // these by id and adds custom tabs/groups/buttons. `controls`/`curGroupId` are
  // set by group()/row(); item helpers call ribRegister as they append.
  let curTabId = "";
  let curGroupId = "";
  const ribTabsById = new Map<string, { btn: HTMLButtonElement; panel: HTMLDivElement }>();
  const ribGroupsById = new Map<string, { tabId: string; groupEl: HTMLElement; container: HTMLElement }>();
  const ribItemsById = new Map<string, { groupId: string; el: HTMLElement }>();
  const ribSlug = (s: string): string =>
    s.toLowerCase().replace(/\([^)]*\)/g, " ").replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 40) || "item";
  const ribUniqueId = (root: string, taken: { has: (k: string) => boolean }): string => {
    let id = root;
    for (let n = 2; taken.has(id); n++) id = `${root}-${n}`;
    return id;
  };
  const ribRegister = (element: HTMLElement, base: string): HTMLElement => {
    if (!curGroupId) return element; // defensive: only ribbon controls register
    const id = ribUniqueId(`${curGroupId}.${ribSlug(base)}`, ribItemsById);
    element.dataset["ribbonItem"] = id;
    ribItemsById.set(id, { groupId: curGroupId, el: element });
    return element;
  };

  // ---- anchored popovers (palettes, menus, pickers, dialogs) --------------
  let activePop: HTMLElement | null = null;
  const closePop = (): void => {
    if (!activePop) return;
    activePop.remove();
    activePop = null;
    document.removeEventListener("mousedown", onDocDown, true);
    document.removeEventListener("keydown", onPopKey, true);
  };
  const onDocDown = (e: MouseEvent): void => {
    if (activePop && !activePop.contains(e.target as Node)) closePop();
  };
  const onPopKey = (e: KeyboardEvent): void => {
    if (e.key === "Escape") {
      closePop();
      editor.focus();
    }
  };
  /** Open `content` in a popover anchored under `anchor`, clamped on-screen. */
  const openPop = (anchor: PopoverAnchor, content: HTMLElement): HTMLElement => {
    closePop();
    const pop = el("div", "cw-pop");
    pop.addEventListener("mousedown", (e) => e.preventDefault()); // keep editor focus
    pop.appendChild(content);
    document.body.appendChild(pop);
    const b = anchor instanceof HTMLElement ? anchor.getBoundingClientRect() : anchor;
    const aTop = b.top;
    const aBottom = "bottom" in b ? b.bottom : b.top + b.height;
    pop.style.left = `${Math.round(b.left)}px`;
    pop.style.top = `${Math.round(aBottom + 3)}px`;
    const pr = pop.getBoundingClientRect();
    if (pr.right > window.innerWidth - 6) pop.style.left = `${Math.round(window.innerWidth - pr.width - 6)}px`;
    if (pr.bottom > window.innerHeight - 6) pop.style.top = `${Math.round(aTop - pr.height - 3)}px`;
    activePop = pop;
    setTimeout(() => {
      document.addEventListener("mousedown", onDocDown, true);
      document.addEventListener("keydown", onPopKey, true);
    }, 0);
    return pop;
  };
  /** A simple vertical menu of labelled actions; `current` marks the active row. */
  const menu = (
    items: { label: string; sample?: string; current?: boolean; onClick: () => void }[],
  ): HTMLElement => {
    const m = el("div", "cw-menu");
    for (const it of items) {
      const b = el("button");
      const check = el("span", "check");
      check.textContent = it.current ? "✓" : "";
      const lbl = el("span");
      lbl.textContent = it.label;
      if (it.sample) lbl.className = "sample";
      b.append(check, lbl);
      b.addEventListener("click", () => {
        closePop();
        it.onClick();
      });
      m.appendChild(b);
    }
    return m;
  };

  // Word colour palette: greys, then a standard-colour row set.
  const PALETTE = [
    "#000000", "#444444", "#666666", "#999999", "#cccccc", "#ffffff",
    "#c00000", "#ff0000", "#ffc000", "#ffff00", "#92d050", "#00b050",
    "#00b0f0", "#0070c0", "#002060", "#7030a0", "#e84393", "#a0522d",
  ];
  // Theme colours (issue #244, D1): the default Office theme slots
  // (bg1/tx1/bg2/tx2 + accent1..6). NOTE: these are the Office DEFAULT theme, not
  // (yet) the imported document's own a:clrScheme — surfacing the loaded doc's real
  // theme palette needs runtime plumbing and is deferred; this gives Word-parity
  // theme swatches for the common (default-theme) case.
  const THEME_COLORS = [
    "#ffffff", "#000000", "#e7e6e6", "#44546a", "#4472c4", "#ed7d31",
    "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47",
  ];
  // Per-editor "recently used colours", persisted (localStorage) and shared by every
  // colour picker (fill, outline, font colour, highlight) via `colorPopover`.
  const COLOR_RECENT_KEY = "cw-color-recent";
  const recordColor = (c: string): void => { pushRecent(COLOR_RECENT_KEY, c.toLowerCase(), 10); };

  // One shared native colour picker, reused for "More colours…". We apply on
  // `change` only (not `input`) so a drag is one undo step.
  const colorPicker = el("input");
  colorPicker.type = "color";
  colorPicker.style.cssText = "position:fixed;left:-9999px;width:0;height:0;";
  document.body.appendChild(colorPicker);
  detachables.push(colorPicker);
  const pickColor = (initial: string, onPick: (c: string) => void): void => {
    colorPicker.value = initial;
    colorPicker.oninput = null;
    colorPicker.onchange = () => onPick(colorPicker.value);
    colorPicker.click();
  };
  /** A swatch grid over `colors`; each click closes the popover then calls `pick`. */
  const swatchGrid = (colors: string[], pick: (c: string) => void): HTMLElement => {
    const grid = el("div", "cw-swatches");
    for (const c of colors) {
      const s = el("button");
      s.style.background = c;
      s.title = c;
      s.addEventListener("click", () => {
        closePop();
        pick(c);
      });
      grid.appendChild(s);
    }
    return grid;
  };
  /** Append a labelled swatch row (theme / recent colours) to `wrap`. */
  const swatchSection = (wrap: HTMLElement, title: string, colors: string[], pick: (c: string) => void): void => {
    const head = el("div", "cw-grid-cat");
    head.textContent = title;
    wrap.append(head, swatchGrid(colors, pick));
  };
  /** Colour palette popover: main swatch grid + Theme + Recent rows + an optional
   *  clear row + "More…". Every pick is recorded into the shared recent-colours list
   *  (issue #244, D1), so the row grows across all colour surfaces. */
  const colorPopover = (
    anchor: PopoverAnchor,
    last: string,
    clearLabel: string | null,
    onPick: (c: string) => void,
    onClear: (() => void) | null,
  ): void => {
    const pick = (c: string): void => { recordColor(c); onPick(c); };
    const wrap = el("div");
    wrap.appendChild(swatchGrid(PALETTE, pick));
    swatchSection(wrap, "Theme colours", THEME_COLORS, pick);
    const recent = readRecent(COLOR_RECENT_KEY, 10);
    if (recent.length) swatchSection(wrap, "Recent colours", recent, pick);
    if (clearLabel && onClear) {
      const clear = el("button", "pop-action");
      clear.textContent = clearLabel;
      clear.addEventListener("click", () => {
        closePop();
        onClear();
      });
      wrap.appendChild(clear);
    }
    const more = el("button", "pop-action");
    more.textContent = "More colours…";
    more.addEventListener("click", () => {
      closePop();
      pickColor(last, pick);
    });
    wrap.appendChild(more);
    openPop(anchor, wrap);
  };

  // --- drawing-shape fill / outline pickers (both surfaces share these) --------
  const SHAPE_STROKE_WIDTHS = [0.5, 1, 1.5, 2, 3, 4.5, 6];
  const SHAPE_DASHES: { dash: ShapeDash; name: string }[] = [
    { dash: "solid", name: "Solid" },
    { dash: "dash", name: "Dash" },
    { dash: "dot", name: "Dot" },
    { dash: "dashDot", name: "Dash-dot" },
    { dash: "lgDash", name: "Long dash" },
  ];
  const dispatchShape = (id: string, patch: Parameters<typeof setShapeProps>[1]): void => {
    editor.dispatch(setShapeProps(id, patch));
    editor.focus();
  };
  /** Fill picker for the selected shape — the shared colour popover + a "No fill"
   *  row (reuses the exact picker font-colour / highlight use). */
  shapeFillPopover = (anchor: PopoverAnchor): void => {
    const id = editor.getSelectedObject();
    const shape = editor.getSelectedShape();
    if (!id || !shape) return;
    const last = shape.fill && "color" in shape.fill ? shape.fill.color : "#bcd6ef";
    colorPopover(
      anchor,
      last,
      "No fill",
      (c) => dispatchShape(id, { fill: { color: c } }),
      () => dispatchShape(id, { fill: { none: true } }),
    );
  };
  /** Outline picker for the selected shape: colour swatches + width + dash selectors
   *  + a "No outline" row. Each control keeps the shape's other outline facets. */
  shapeOutlinePopover = (anchor: PopoverAnchor): void => {
    const id = editor.getSelectedObject();
    const shape = editor.getSelectedShape();
    if (!id || !shape) return;
    // The concrete outline the controls edit (defaults when the shape has none).
    const cur = shape.stroke && "color" in shape.stroke ? shape.stroke : null;
    const base = (): { color: string; widthPt: number; dash?: ShapeDash } => ({
      color: cur?.color ?? "#41719c",
      widthPt: cur?.widthPt ?? 1,
      ...(cur?.dash ? { dash: cur.dash } : {}),
    });
    const wrap = el("div");
    const pickStroke = (c: string): void => { recordColor(c); dispatchShape(id, { stroke: { ...base(), color: c } }); };
    wrap.appendChild(swatchGrid(PALETTE, pickStroke));
    swatchSection(wrap, "Theme colours", THEME_COLORS, pickStroke);
    const recentOutline = readRecent(COLOR_RECENT_KEY, 10);
    if (recentOutline.length) swatchSection(wrap, "Recent colours", recentOutline, pickStroke);
    // Width + dash rows: a labelled <select> each, applied on change. Styled via the
    // popover's own class system (.cw-pop .pop-row) so they match the swatch grid (D2).
    const mkRow = (labelText: string, sel: HTMLSelectElement): void => {
      const row = el("div", "pop-row");
      const lab = el("span", "pop-row-label");
      lab.textContent = labelText;
      row.append(lab, sel);
      wrap.appendChild(row);
    };
    const widthSel = el("select");
    for (const w of SHAPE_STROKE_WIDTHS) {
      const o = el("option");
      o.value = String(w);
      o.textContent = `${w} pt`;
      if (Math.abs(w - (cur?.widthPt ?? 1)) < 1e-6) o.selected = true;
      widthSel.appendChild(o);
    }
    widthSel.addEventListener("change", () => {
      dispatchShape(id, { stroke: { ...base(), widthPt: Number(widthSel.value) } });
    });
    mkRow("Width", widthSel);
    const dashSel = el("select");
    for (const d of SHAPE_DASHES) {
      const o = el("option");
      o.value = d.dash;
      o.textContent = d.name;
      if ((cur?.dash ?? "solid") === d.dash) o.selected = true;
      dashSel.appendChild(o);
    }
    dashSel.addEventListener("change", () => {
      const b = base();
      const dash = dashSel.value as ShapeDash;
      dispatchShape(id, { stroke: dash === "solid" ? { color: b.color, widthPt: b.widthPt } : { color: b.color, widthPt: b.widthPt, dash } });
    });
    mkRow("Dash", dashSel);
    const none = el("button", "pop-action");
    none.textContent = "No outline";
    none.addEventListener("click", () => {
      closePop();
      dispatchShape(id, { stroke: { none: true } });
    });
    wrap.appendChild(none);
    const more = el("button", "pop-action");
    more.textContent = "More colours…";
    more.addEventListener("click", () => {
      closePop();
      pickColor(base().color, pickStroke);
    });
    wrap.appendChild(more);
    openPop(anchor, wrap);
  };
  /** Word's split colour button: the face applies the last colour, the caret
   *  opens a palette popover. `clearLabel`/`onClear` add a "No Color"/"Automatic"
   *  row; `active` wires the face into the pressed-state sync (highlight toggles). */
  const swatch = (cfg: {
    face: string;
    initial: string;
    title: string;
    apply: (color: string) => void;
    clearLabel?: string;
    onClear?: () => void;
    active?: (f: CurrentFormat) => boolean;
  }): void => {
    let last = cfg.initial;
    const wrap = el("div");
    wrap.style.cssText = "display:flex;align-items:stretch;";
    const main = el("button", "rib-btn rib-swatch");
    main.title = cfg.title;
    main.setAttribute("aria-label", cfg.title); // not the bare "A"/highlight glyph (C6)
    main.innerHTML = `<span class="row" aria-hidden="true">${cfg.face}</span><span class="bar" style="background:${last}"></span>`;
    const bar = main.querySelector(".bar") as HTMLElement;
    main.addEventListener("mousedown", (e) => e.preventDefault());
    main.addEventListener("click", () => {
      cfg.apply(last);
      editor.focus();
    });
    const more = el("button", "rib-btn");
    more.title = `${cfg.title} — choose colour`;
    more.setAttribute("aria-label", more.title);
    more.style.cssText = "min-width:14px;padding:0;";
    more.innerHTML = CARET;
    more.addEventListener("mousedown", (e) => e.preventDefault());
    const pick = (c: string): void => {
      last = c;
      bar.style.background = c;
      cfg.apply(c);
      editor.focus();
    };
    more.addEventListener("click", () =>
      colorPopover(more, last, cfg.clearLabel ?? null, pick, cfg.onClear ? () => {
        cfg.onClear!();
        editor.focus();
      } : null),
    );
    wrap.append(main, more);
    controls.appendChild(wrap);
    ribRegister(wrap, cfg.title);
    if (cfg.active) toggleButtons.push({ el: main, active: cfg.active });
  };

  /** A list button split like Word's: the icon toggles the default list; the
   *  caret opens a style picker (bullet glyphs / number formats). The picker
   *  rows carry a marker preview in their label. */
  const listSplit = (
    icon: string,
    title: string,
    faceCmd: () => Command,
    active: (f: CurrentFormat) => boolean,
    styles: { label: string; cmd: () => Command }[],
  ): void => {
    const wrap = el("div");
    wrap.style.cssText = "display:flex;align-items:stretch;";
    const main = el("button", "rib-btn");
    main.title = title;
    main.setAttribute("aria-label", title);
    main.innerHTML = icon;
    main.addEventListener("mousedown", (e) => e.preventDefault());
    main.addEventListener("click", () => {
      editor.dispatch(faceCmd());
      editor.focus();
    });
    toggleButtons.push({ el: main, active });
    const more = el("button", "rib-btn");
    more.title = `${title} — choose style`;
    more.setAttribute("aria-label", more.title);
    more.style.cssText = "min-width:14px;padding:0;";
    more.innerHTML = CARET;
    more.addEventListener("mousedown", (e) => e.preventDefault());
    more.addEventListener("click", () =>
      openPop(
        more,
        menu(
          styles.map((s) => ({
            label: s.label,
            onClick: () => {
              editor.dispatch(s.cmd());
              editor.focus();
            },
          })),
        ),
      ),
    );
    wrap.append(main, more);
    controls.appendChild(wrap);
    ribRegister(wrap, title);
  };

  /** Word's table-size grid: hover to size, click to insert. */
  const tableGridPopover = (anchor: HTMLElement): void => {
    const ROWS = 8;
    const COLS = 10;
    const wrap = el("div");
    const grid = el("div", "cw-grid");
    grid.style.gridTemplateColumns = `repeat(${COLS}, 15px)`;
    const label = el("div", "cw-grid-label");
    label.textContent = "Insert table";
    const cells: HTMLElement[][] = [];
    const highlight = (R: number, C: number): void => {
      for (let r = 0; r < ROWS; r++)
        for (let c = 0; c < COLS; c++) cells[r]![c]!.classList.toggle("on", r <= R && c <= C);
    };
    for (let r = 0; r < ROWS; r++) {
      const row: HTMLElement[] = [];
      for (let c = 0; c < COLS; c++) {
        const cell = el("div", "cell");
        cell.addEventListener("mouseenter", () => {
          highlight(r, c);
          label.textContent = `${c + 1} × ${r + 1} table`;
        });
        cell.addEventListener("click", () => {
          closePop();
          editor.dispatch(insertTable(r + 1, c + 1));
          editor.focus();
        });
        row.push(cell);
        grid.appendChild(cell);
      }
      cells.push(row);
    }
    wrap.append(grid, label);
    openPop(anchor, wrap);
  };

  /** Shapes gallery: a preset picker that inserts the chosen shape at the caret.
   *  Mirrors the table grid picker's popover pattern. */
  const shapesPopover = (anchor: HTMLElement): void => {
    const wrap = el("div");
    const COLS = 4;
    const label = el("div", "cw-grid-label");
    label.textContent = "Insert shape";
    // Every gallery item keyed by a stable id (preset name, or "textbox" for the
    // special text-box entry). Categories reference these ids, and the "recently
    // used" row rebuilds cells from them (issue #244, A3). The Text Box entry seeds
    // an editable body + enters editing (issue #235); presets select+reveal (A1).
    const items: Record<string, { icon: string; name: string; insert: () => void }> = {
      line: { icon: ICONS.shapeLine, name: "Line", insert: () => editor.insertShape("line") },
      rect: { icon: ICONS.shapeRect, name: "Rectangle", insert: () => editor.insertShape("rect") },
      roundRect: { icon: ICONS.shapeRoundRect, name: "Rounded rectangle", insert: () => editor.insertShape("roundRect") },
      ellipse: { icon: ICONS.shapeEllipse, name: "Ellipse", insert: () => editor.insertShape("ellipse") },
      triangle: { icon: ICONS.shapeTriangle, name: "Triangle", insert: () => editor.insertShape("triangle") },
      diamond: { icon: ICONS.shapeDiamond, name: "Diamond", insert: () => editor.insertShape("diamond") },
      pentagon: { icon: ICONS.shapePentagon, name: "Pentagon", insert: () => editor.insertShape("pentagon") },
      hexagon: { icon: ICONS.shapeHexagon, name: "Hexagon", insert: () => editor.insertShape("hexagon") },
      star5: { icon: ICONS.shapeStar5, name: "5-point star", insert: () => editor.insertShape("star5") },
      parallelogram: { icon: ICONS.shapeParallelogram, name: "Parallelogram", insert: () => editor.insertShape("parallelogram") },
      trapezoid: { icon: ICONS.shapeTrapezoid, name: "Trapezoid", insert: () => editor.insertShape("trapezoid") },
      textbox: { icon: ICONS.shapeTextBox, name: "Text Box", insert: () => editor.insertTextBox() },
      rightArrow: { icon: ICONS.shapeRightArrow, name: "Right arrow", insert: () => editor.insertShape("rightArrow") },
      leftArrow: { icon: ICONS.shapeLeftArrow, name: "Left arrow", insert: () => editor.insertShape("leftArrow") },
      upArrow: { icon: ICONS.shapeUpArrow, name: "Up arrow", insert: () => editor.insertShape("upArrow") },
      downArrow: { icon: ICONS.shapeDownArrow, name: "Down arrow", insert: () => editor.insertShape("downArrow") },
    };
    // Word/PowerPoint-style category grouping so the widened set stays scannable.
    const categories: { title: string; ids: string[] }[] = [
      { title: "Lines", ids: ["line"] },
      { title: "Basic Shapes", ids: ["rect", "roundRect", "ellipse", "triangle", "diamond", "pentagon", "hexagon", "star5", "parallelogram", "trapezoid", "textbox"] },
      { title: "Block Arrows", ids: ["rightArrow", "leftArrow", "upArrow", "downArrow"] },
    ];
    const RECENT_KEY = "cw-shape-recent";
    // Keyboard-operable gallery (E1): every cell across every section is a
    // role="button" in ONE roving-tabindex list — arrows move focus (Home/End jump
    // to the ends), Enter/Space activate — so the whole gallery is a single Tab stop.
    const cellEls: HTMLElement[] = [];
    const focusCell = (i: number): void => {
      const j = Math.max(0, Math.min(cellEls.length - 1, i));
      cellEls.forEach((c, k) => (c.tabIndex = k === j ? 0 : -1));
      cellEls[j]!.focus();
    };
    const buildCell = (id: string): HTMLElement => {
      const it = items[id]!;
      const i = cellEls.length;
      const c = el("div", "cell");
      c.style.width = "28px";
      c.style.height = "28px";
      c.innerHTML = it.icon;
      c.title = it.name;
      c.setAttribute("role", "button");
      c.setAttribute("aria-label", it.name);
      c.tabIndex = i === 0 ? 0 : -1; // roving: only the first cell starts tabbable
      const show = (): void => { label.textContent = it.name; };
      const activate = (): void => { pushRecent(RECENT_KEY, id, 8); closePop(); it.insert(); };
      c.addEventListener("mouseenter", show);
      c.addEventListener("focus", show);
      c.addEventListener("click", activate);
      c.addEventListener("keydown", (ev: KeyboardEvent) => {
        switch (ev.key) {
          case "ArrowRight": focusCell(i + 1); break;
          case "ArrowLeft": focusCell(i - 1); break;
          case "ArrowDown": focusCell(i + COLS); break;
          case "ArrowUp": focusCell(i - COLS); break;
          case "Home": focusCell(0); break;
          case "End": focusCell(cellEls.length - 1); break;
          case "Enter": case " ": activate(); break;
          default: return; // let other keys (Escape, Tab) bubble to the popover
        }
        ev.preventDefault();
      });
      cellEls.push(c);
      return c;
    };
    const section = (title: string, ids: string[]): void => {
      const head = el("div", "cw-grid-cat");
      head.textContent = title;
      const grid = el("div", "cw-grid");
      grid.setAttribute("role", "group");
      grid.setAttribute("aria-label", title);
      grid.style.gridTemplateColumns = `repeat(${COLS}, 28px)`;
      for (const id of ids) grid.appendChild(buildCell(id));
      wrap.append(head, grid);
    };
    // "Recently used" first (Word parity) — only the ids we still know how to build.
    const recent = readRecent(RECENT_KEY, 8).filter((id) => id in items);
    if (recent.length) section("Recently used", recent);
    for (const c of categories) section(c.title, c.ids);
    wrap.append(label);
    openPop(anchor, wrap);
    // Move focus into the gallery so the arrows/Enter drive it straight away (the
    // popover keeps editor focus on mouse-open; a keyboard user lands on cell 0).
    setTimeout(() => cellEls[0]?.focus(), 0);
  };

  /** Hyperlink dialog: a web address OR a place in this document (a bookmark, via
   *  an `#anchor` link — this is what gives bookmarks a consumer, critique B7).
   *  `url` pre-fills the field; a leading `#` opens the document-anchor mode with
   *  that bookmark selected. */
  const linkDialog = (anchor: HTMLElement, url = ""): void => {
    const bookmarks = Object.keys(editor.getDocument().bookmarks ?? {}).sort((a, b) => a.localeCompare(b));
    const startInternal = url.startsWith("#");
    const wrap = el("div", "cw-dialog");

    // Mode toggle — Web address / This document.
    const tabs = el("div", "row");
    tabs.style.cssText = "gap:4px;margin-bottom:8px;";
    const TAB_BASE = "flex:1 1 auto;padding:5px 8px;border:1px solid #d0d4d9;border-radius:6px;cursor:pointer;font-size:12px;";
    const TAB_ACTIVE = TAB_BASE + "border-color:#1a73e8;color:#1a73e8;background:#eef4fe;";
    const TAB_IDLE = TAB_BASE + "color:#3c4043;background:#fff;";
    const mkTab = (label: string): HTMLButtonElement => {
      const b = el("button");
      b.textContent = label;
      b.style.cssText = TAB_IDLE;
      b.addEventListener("mousedown", (e) => e.preventDefault());
      return b;
    };
    const webTab = mkTab("Web address");
    const docTab = mkTab("This document");
    tabs.append(webTab, docTab);

    // Web address field.
    const lab = el("label");
    lab.textContent = "Address";
    const input = el("input");
    input.type = "text";
    input.placeholder = "https://example.com";
    input.value = startInternal ? "" : url;
    input.addEventListener("mousedown", (e) => e.stopPropagation()); // allow focus
    lab.appendChild(input);

    // Bookmark picker (in-document anchor).
    const bmLab = el("label");
    bmLab.textContent = "Bookmark";
    const bmSel = el("select");
    bmSel.style.cssText = "width:100%;height:28px;border:1px solid #d0d4d9;border-radius:6px;padding:0 6px;";
    for (const b of bookmarks) {
      const o = document.createElement("option");
      o.value = b;
      o.textContent = b;
      bmSel.append(o);
    }
    if (startInternal) bmSel.value = url.slice(1);
    bmSel.disabled = bookmarks.length === 0;
    bmSel.addEventListener("mousedown", (e) => e.stopPropagation());
    bmLab.appendChild(bmSel);
    const bmNote = el("div");
    bmNote.style.cssText = "font-size:11px;color:#9aa0a6;margin-top:4px;";
    bmNote.textContent = "No bookmarks yet — add one from Navigator ▸ Marks.";

    let mode: "web" | "doc" = startInternal && bookmarks.length > 0 ? "doc" : "web";
    const applyMode = (): void => {
      const web = mode === "web";
      lab.style.display = web ? "" : "none";
      bmLab.style.display = web ? "none" : "";
      bmNote.style.display = !web && bookmarks.length === 0 ? "" : "none";
      webTab.style.cssText = web ? TAB_ACTIVE : TAB_IDLE;
      docTab.style.cssText = web ? TAB_IDLE : TAB_ACTIVE;
      apply.disabled = !web && bookmarks.length === 0;
    };
    webTab.addEventListener("click", () => { mode = "web"; applyMode(); setTimeout(() => input.focus(), 0); });
    docTab.addEventListener("click", () => { mode = "doc"; applyMode(); });

    const row = el("div", "row");
    const remove = el("button", "danger");
    remove.textContent = "Remove";
    remove.addEventListener("click", () => {
      closePop();
      editor.dispatch(setLinkCmd(null));
      editor.focus();
    });
    const apply = el("button", "primary");
    apply.textContent = "Apply";
    const doApply = (): void => {
      const target = mode === "doc"
        ? (bmSel.value ? `#${bmSel.value}` : null)
        : (input.value.trim() === "" ? null : input.value.trim());
      if (mode === "doc" && bookmarks.length === 0) return;
      closePop();
      editor.dispatch(setLinkCmd(target));
      editor.focus();
    };
    apply.addEventListener("click", doApply);
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") doApply();
      e.stopPropagation();
    });
    row.append(remove, apply);
    wrap.append(tabs, lab, bmLab, bmNote, row);
    applyMode();
    openPop(anchor, wrap);
    setTimeout(() => (mode === "web" ? input : bmSel).focus(), 0);
  };
  // Expose the link dialog to the hyperlink context toolbar (built outside this
  // ribbon closure).
  openLinkDialog = linkDialog;

  /** Open the Font dialog seeded from the caret's effective style; Apply routes
   *  the edited patch through setCharStyle (one undoable edit over the selection,
   *  or the pending typing-style when the caret is collapsed). */
  const openFontDialog = (): void => {
    const f = editor.currentFormat();
    showFontDialog({
      initial: {
        underline: f.underline,
        underlineStyle: f.underlineStyle,
        underlineColor: f.underlineColor,
        caps: f.caps,
        smallCaps: f.smallCaps,
        doubleStrikethrough: f.doubleStrikethrough,
        positionPx: f.positionPx,
        widthScalePct: f.widthScalePct,
        letterSpacingPx: f.letterSpacingPx,
        kerningMinPx: f.kerningMinPx,
        emphasisMark: f.emphasisMark,
        outline: f.outline,
        shadow: f.shadow,
        emboss: f.emboss,
        imprint: f.imprint,
        fitTextPx: f.fitTextPx,
      },
      onApply: (patch) => {
        editor.setCharStyle(patch);
        editor.focus();
      },
    });
  };

  const exportAs = async (format: ExportFormat): Promise<void> => {
    // Rendering (especially PDF) can take a few seconds; show the busy overlay so
    // the app doesn't look hung. Dropped before any onSave dialog / download so it
    // doesn't sit behind a native Save sheet.
    const busy = showBusy(`Exporting ${format.toUpperCase()}…`);
    try {
      const { bytes, warnings } = await exportBaked(format);
      busy.done();
      if (warnings.length > 0) console.warn(`[export-${format}] warnings`, warnings);
      // Linked images blocked by the host's CORS policy embed as gray placeholder
      // boxes — raise a visible notice so it isn't a silent surprise (the editor
      // shows the images fine; only reading their bytes for export is CORS-gated).
      const extImg = warnings.find((w) => w.code === "image-external-unfetchable");
      if (extImg) {
        const n = extImg.count;
        showNotice(
          `${n} linked image${n === 1 ? "" : "s"} couldn't be embedded in the ${format.toUpperCase()} — the image host didn't ` +
            `allow this site to read them (a CORS restriction), so they appear as gray boxes. They still show in the editor.`,
          "warning",
        );
      }
      const blob = blobFor(format, bytes);
      // With an embedder onSave hook, hand off the file instead of downloading;
      // the host routes it to its own pipeline (upload, persist, custom download).
      if (runtime.onSave) {
        await runtime.onSave({ blob, bytes, format, warnings });
        return;
      }
      const url = URL.createObjectURL(blob);
      const a = el("a");
      a.href = url;
      a.download = `${currentDocBaseName()}.${format}`;
      // Attach before clicking: a detached <a download> is navigated-to (not
      // downloaded) in some browsers/headless, which reloads the page.
      a.style.display = "none";
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (err) {
      console.error(`[export-${format}] failed`, err);
    } finally {
      busy.done(); // idempotent — covers the error path
    }
  };
  const stylesheet = (): ReturnType<typeof defaultStylesheet> =>
    editor.getDocument().stylesheet ?? defaultStylesheet();

  // ===== File tab (simplified backstage: open / export / undo) =============
  const fileTab = tab("file", "File", "file");
  group(fileTab, "Open");
  bigBtn(ICONS.open, "Open<br>.docx", "Open a Word document", () => {
    const input = el("input");
    input.type = "file";
    input.accept = `.docx,${DOCX_MIME}`;
    input.addEventListener("change", () => {
      const file = input.files?.[0];
      if (file) void openDocxFile(file);
    });
    input.click();
  });

  // Share (online only): publish the current doc and surface a join link. Also
  // used by openDocx auto-publish via showShareDialog below.
  showShareDialog = (link: string): void => {
    const back = el("div");
    back.style.cssText =
      "position:fixed;inset:0;z-index:60;background:rgba(0,0,0,0.25);display:flex;align-items:center;justify-content:center;";
    const wrap = el("div", "cw-dialog");
    wrap.style.cssText = "background:#fff;border:1px solid #c8c6c4;border-radius:8px;box-shadow:0 8px 30px rgba(0,0,0,0.25);";
    const lab = el("label");
    lab.textContent = "Share this document — anyone with the link can join and edit live";
    const input = el("input");
    input.type = "text";
    input.value = link;
    input.readOnly = true;
    lab.appendChild(input);
    const row = el("div", "row");
    const closeBtn = el("button");
    closeBtn.textContent = "Close";
    closeBtn.addEventListener("click", () => back.remove());
    const copyBtn = el("button", "primary");
    copyBtn.textContent = "Copy link";
    copyBtn.addEventListener("click", () => {
      void navigator.clipboard?.writeText(link);
      copyBtn.textContent = "Copied!";
      setTimeout(() => (copyBtn.textContent = "Copy link"), 1500);
    });
    back.addEventListener("mousedown", (e) => {
      if (e.target === back) back.remove();
    });
    row.append(closeBtn, copyBtn);
    wrap.append(lab, row);
    back.appendChild(wrap);
    document.body.appendChild(back);
    setTimeout(() => input.select(), 0);
  };

  if (online) {
    group(fileTab, "Share");
    bigBtn(ICONS.share, "Share<br>Link", "Publish & get a shareable link", () => {
      void goOnlineWithCurrentDoc().catch((e) => showNotice(`Share failed: ${e instanceof Error ? e.message : String(e)}`, "warning"));
    });
  }

  // Export buttons are independently hideable (view.exportPdf / view.exportDocx)
  // for embedders that route saving through their own pipeline (onSave / the
  // exportDocx()/exportPdf() handle methods). Skip the group entirely if both off.
  const showExportPdf = runtime.view?.exportPdf ?? true;
  const showExportDocx = runtime.view?.exportDocx ?? true;
  if (showExportPdf || showExportDocx) {
    group(fileTab, "Export");
    if (showExportPdf) txtBtn("PDF", "Export to PDF", () => void exportAs("pdf"), "font-size:11px;font-weight:600;");
    if (showExportDocx) txtBtn("DOCX", "Export to .docx", () => void exportAs("docx"), "font-size:11px;font-weight:600;");
  }
  group(fileTab, "Undo");
  btn(ICONS.undo, "Undo (Ctrl+Z)", () => editor.undo());
  btn(ICONS.redo, "Redo (Ctrl+Y)", () => editor.redo());

  // ===== Home tab ==========================================================
  const home = tab("home", "Home");

  // ---- Clipboard ----
  group(home, "Clipboard");
  bigBtn(ICONS.paste, "Paste", "Paste (Ctrl+V)", () => editor.paste(), true);
  {
    const stack = el("div");
    stack.style.cssText = "display:flex;flex-direction:column;gap:1px;";
    const prev = controls;
    controls = stack;
    btn(ICONS.cut, "Cut (Ctrl+X)", () => editor.cut());
    btn(ICONS.copy, "Copy (Ctrl+C)", () => editor.copy());
    btn(ICONS.painter, "Format painter (double-click = sticky)", () => editor.armFormatPainter(false));
    controls = prev;
    controls.appendChild(stack);
  }

  // ---- Font ----
  const fontRow = groupRows(home, "Font");
  fontRow();
  const fontSelect = select("Font family", 150);
  // Labels show the bundled clone we actually render — e.g. "Calibri (Carlito)".
  // Per-instance: built-ins minus disableBuiltin, plus this mount's custom fonts.
  for (const f of toolbarFonts(config.fonts)) opt(fontSelect, f.value, f.label);
  fontSelect.addEventListener("change", () => {
    editor.setCharStyle({ fontFamily: fontSelect.value });
    editor.focus();
  });
  // The model is px-native; Word shows POINTS. Convert on read/write (96dpi:
  // 1pt = 4/3 px). The displayed value snaps to the nearest half-point.
  const pxToPt = (px: number): number => Math.round(sharedPxToPt(px) * 2) / 2;
  const ptToPx = (pt: number): number => sharedPtToPx(pt);
  const sizeInput = el("input");
  sizeInput.type = "number";
  sizeInput.lang = "en"; // dot decimal for 10.5pt regardless of OS locale (V4)
  sizeInput.min = "1";
  sizeInput.max = "72";
  sizeInput.step = "0.5";
  sizeInput.title = "Font size (pt)";
  sizeInput.style.cssText = "width:46px;padding:0 4px;";
  sizeInput.addEventListener("mousedown", (e) => e.stopPropagation());
  /** Apply a size given in POINTS (clamped to the model's 6–96px range). */
  const setSizePt = (pt: number): void => {
    if (!Number.isFinite(pt)) return;
    const px = Math.min(96, Math.max(6, ptToPx(pt)));
    editor.setCharStyle({ fontSizePx: px });
    editor.focus();
  };
  sizeInput.addEventListener("change", () => setSizePt(Number(sizeInput.value)));
  // Editable combo: type any size in the field, or click the caret for a Word
  // preset. (A `<datalist>` dropdown never opens for `type=number` in Chromium,
  // so we render our own menu — which also keeps the typed field + spinner.)
  const SIZE_PRESETS = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];
  const sizeWrap = el("div");
  sizeWrap.style.cssText = "display:inline-flex;align-items:stretch;";
  const sizeCaret = el("button", "rib-btn");
  sizeCaret.title = "Font size presets";
  sizeCaret.setAttribute("aria-label", "Font size presets");
  sizeCaret.innerHTML = CARET;
  sizeCaret.style.cssText = "padding:0 1px;min-width:14px;";
  sizeCaret.addEventListener("mousedown", (e) => e.preventDefault());
  sizeCaret.addEventListener("click", () => {
    const r = sizeWrap.getBoundingClientRect();
    const entries: MenuEntry[] = SIZE_PRESETS.map((p) => ({
      kind: "item",
      label: String(p),
      onClick: () => {
        sizeInput.value = String(p);
        setSizePt(p);
      },
    }));
    showContextMenu(r.left, r.bottom + 2, entries);
  });
  sizeWrap.append(sizeInput, sizeCaret);
  controls.appendChild(sizeWrap);
  ribRegister(sizeWrap, "font-size");
  const curPt = (): number => pxToPt(editor.currentFormat().fontSizePx ?? 16);
  txtBtn("A", "Grow font", () => setSizePt(curPt() + 1), "font-size:15px;font-weight:600;");
  txtBtn("A", "Shrink font", () => setSizePt(curPt() - 1), "font-size:10px;font-weight:600;");
  const CASES: { label: string; mode: CaseMode }[] = [
    { label: "Sentence case", mode: "sentence" },
    { label: "lowercase", mode: "lower" },
    { label: "UPPERCASE", mode: "upper" },
    { label: "Capitalize Each Word", mode: "title" },
    { label: "tOGGLE cASE", mode: "toggle" },
  ];
  const caseBtn = txtBtn("Aa", "Change case", () => {}, "font-weight:600;", true);
  caseBtn.addEventListener("click", () =>
    openPop(
      caseBtn,
      menu(
        CASES.map((c) => ({
          label: c.label,
          onClick: () => {
            editor.dispatch(changeCaseCmd(c.mode));
            editor.focus();
          },
        })),
      ),
    ),
  );
  btn(ICONS.clearFormat, "Clear all formatting", () => {
    editor.dispatch(clearCharFormatting());
    editor.dispatch(applyNamedStyle("Normal"));
    editor.focus();
  });

  fontRow();
  toggle(txtBtn("B", "Bold (Ctrl+B)", () => editor.toggleStyle("bold"), "font-weight:700;"), (f) => f.bold);
  toggle(txtBtn("I", "Italic (Ctrl+I)", () => editor.toggleStyle("italic"), "font-style:italic;font-family:Georgia,serif;"), (f) => f.italic);
  toggle(txtBtn("U", "Underline (Ctrl+U)", () => editor.toggleStyle("underline"), "text-decoration:underline;"), (f) => f.underline);
  toggle(txtBtn("ab", "Strikethrough", () => editor.toggleStyle("strikethrough"), "text-decoration:line-through;"), (f) => f.strikethrough);
  toggle(txtBtn("x²", "Superscript", () => editor.dispatch(toggleVerticalAlign("super"))), (f) => f.superscript);
  toggle(txtBtn("x₂", "Subscript", () => editor.dispatch(toggleVerticalAlign("sub"))), (f) => f.subscript);
  toggle(txtBtn("AB", "All caps", () => editor.toggleStyle("caps"), "font-size:11px;font-weight:600;letter-spacing:.5px;"), (f) => f.caps);
  toggle(txtBtn("Ab", "Small caps", () => editor.toggleStyle("smallCaps"), "font-variant:small-caps;font-size:12px;font-weight:600;"), (f) => f.smallCaps);
  toggle(txtBtn("ab", "Double strikethrough", () => editor.toggleStyle("doubleStrikethrough"), "text-decoration:line-through double;"), (f) => f.doubleStrikethrough);
  sep();
  // Home-tab dialog-launcher for the full Font dialog (caps, underline style +
  // colour, raise/lower, scaling, spacing, kerning, emphasis, text effects).
  btn(`<span style="color:#2b579a;font-weight:700;">A</span>`, "Font — effects, caps, underline style, spacing", () => openFontDialog());
  // Highlight: face toggles the last colour (re-click clears); caret opens the
  // palette, where "No Color" strips the highlight.
  swatch({
    face: ICONS.highlight,
    initial: "#ffeb3b",
    title: "Text highlight colour",
    apply: (c) => editor.dispatch(toggleHighlight(c)),
    clearLabel: "No Color",
    onClear: () => editor.setCharStyle({ highlightColor: undefined }),
    active: (f) => f.highlight,
  });
  // Font colour: face applies the last colour; palette "Automatic" resets it.
  swatch({
    face: `<span style="font-weight:700;">A</span>`,
    initial: "#e00000",
    title: "Font colour",
    apply: (c) => editor.setCharStyle({ color: c }),
    clearLabel: "Automatic",
    onClear: () => editor.setCharStyle({ color: "#202124" }),
  });
  // Hyperlink lives on Insert ▸ Links (the canonical spot, like Word) — not
  // duplicated here in the Font group.

  // ---- Paragraph ----
  const paraRow = groupRows(home, "Paragraph");
  paraRow();
  listSplit(ICONS.bullets, "Bulleted list", () => toggleList("bullet"), (f) => f.listKind === "bullet", [
    { label: "•   Filled circle", cmd: () => toggleList("bullet") },
    { label: "◦   Hollow circle", cmd: () => applyListStyleCmd(bulletListDefinition("bullets-circle", "◦")) },
    { label: "▪   Filled square", cmd: () => applyListStyleCmd(bulletListDefinition("bullets-square", "▪")) },
    { label: "–   Dash", cmd: () => applyListStyleCmd(bulletListDefinition("bullets-dash", "–")) },
    { label: "➤   Arrow", cmd: () => applyListStyleCmd(bulletListDefinition("bullets-arrow", "➤")) },
  ]);
  listSplit(ICONS.numbering, "Numbered list (Tab/Shift+Tab change level)", () => toggleList("decimal"), (f) => f.listKind === "number", [
    { label: "1.   Decimal", cmd: () => toggleList("decimal") },
    { label: "1)   Decimal, parenthesis", cmd: () => applyListStyleCmd(numberListDefinition("numbers-paren", "decimal", ")")) },
    { label: "a.   Lower letter", cmd: () => applyListStyleCmd(numberListDefinition("numbers-lalpha", "lowerLetter", ".")) },
    { label: "A.   Upper letter", cmd: () => applyListStyleCmd(numberListDefinition("numbers-ualpha", "upperLetter", ".")) },
    { label: "i.   Lower roman", cmd: () => applyListStyleCmd(numberListDefinition("numbers-lroman", "lowerRoman", ".")) },
    { label: "I.   Upper roman", cmd: () => applyListStyleCmd(numberListDefinition("numbers-uroman", "upperRoman", ".")) },
  ]);
  btn(ICONS.multilevel, "Multilevel list (1, 1.1, 1.1.1 — Tab / Shift+Tab change level)", () => editor.dispatch(toggleMultilevelList()));
  sep();
  btn(ICONS.indentDecrease, "Decrease indent", () => editor.dispatch(adjustIndentCmd(-config.behavior.indentStepPx)));
  btn(ICONS.indentIncrease, "Increase indent", () => editor.dispatch(adjustIndentCmd(config.behavior.indentStepPx)));
  marksToggleBtn();

  paraRow();
  toggle(btn(ICONS.alignLeft, "Align left", () => editor.align("left")), (f) => f.align === "left");
  toggle(btn(ICONS.alignCenter, "Center", () => editor.align("center")), (f) => f.align === "center");
  toggle(btn(ICONS.alignRight, "Align right", () => editor.align("right")), (f) => f.align === "right");
  toggle(btn(ICONS.alignJustify, "Justify", () => editor.align("justify")), (f) => f.align === "justify");
  sep();
  // Paragraph writing direction (OOXML w:bidi). RTL right-aligns + reorders.
  toggle(txtBtn("LTR", "Left-to-right paragraph", () => editor.dispatch(setDirection("ltr"))), (f) => f.direction !== "rtl");
  toggle(txtBtn("RTL", "Right-to-left paragraph", () => editor.dispatch(setDirection("rtl"))), (f) => f.direction === "rtl");
  sep();
  const SPACINGS = [
    { v: 1, l: "1.0" },
    { v: 1.15, l: "1.15" },
    { v: 1.5, l: "1.5" },
    { v: 2, l: "2.0" },
    { v: 2.5, l: "2.5" },
    { v: 3, l: "3.0" },
  ];
  let curLineHeight: number | null = null;
  // Opens the full Paragraph dialog (borders, shading, line-spacing rule, widow
  // control, vertical alignment, …) seeded from the caret paragraph's style. One OK
  // = one undo step (the dialog commits a single setParaProps patch).
  let paraDlg: ParagraphDialogHandle | null = null;
  teardown.signal.addEventListener("abort", () => paraDlg?.close(), { once: true });
  const openParagraphDialog = (): void => {
    const ps = editor.currentParaStyle();
    if (!ps) return;
    paraDlg?.close();
    paraDlg = showParagraphDialog(
      {
        borders: ps.borders,
        shading: ps.shading ?? null,
        contextualSpacing: ps.contextualSpacing ?? false,
        lineRule: ps.lineRule ?? "auto",
        lineHeight: ps.lineHeight ?? 1,
        lineHeightPx: ps.lineHeightPx ?? 16,
        widowControl: ps.widowControl ?? true, // Word default is ON
        textAlignment: ps.textAlignment ?? "baseline",
        mirrorIndents: ps.mirrorIndents ?? false,
        suppressLineNumbers: ps.suppressLineNumbers ?? false,
        adjustRightInd: ps.adjustRightInd ?? false,
      },
      {
        apply: (patch) => {
          editor.dispatch(setParaProps(patch));
          editor.focus();
        },
      },
    );
    // Let replaceDocument() tear this dialog down if the editor is rebuilt while it's open.
    closeParaDialog = () => { paraDlg?.close(); paraDlg = null; closeParaDialog = () => {}; };
  };
  const spacingBtn = btn(ICONS.lineSpacing, "Line spacing", () => {}, true);
  spacingBtn.addEventListener("click", () =>
    openPop(
      spacingBtn,
      menu([
        ...SPACINGS.map((s) => ({
          label: s.l,
          current: curLineHeight !== null && Math.abs(curLineHeight - s.v) < 1e-6,
          onClick: () => {
            // Picking a multiplier clears any fixed rule (lineRule/lineHeightPx aren't
            // `| undefined` under exactOptionalPropertyTypes; Object.assign sets the
            // explicit-undefined keys setParaStyle uses to remove them).
            const patch: Partial<ParaStyle> = { lineHeight: s.v };
            Object.assign(patch, { lineRule: undefined, lineHeightPx: undefined });
            editor.dispatch(setParaProps(patch));
            editor.focus();
          },
        })),
        { label: "Line Spacing Options…", onClick: openParagraphDialog },
      ]),
    ),
  );
  // One entry point to the Paragraph dialog, which covers both borders and
  // shading (was two separate buttons opening the identical dialog).
  btn(ICONS.borders, "Borders & shading", openParagraphDialog, true);

  // ---- Styles (visual gallery) ----
  group(home, "Styles");
  const styleGallery = el("div", "rib-gallery");
  const styleCards = new Map<string, HTMLButtonElement>();
  // One child document renders every card's swatch — a real, document-styled
  // sample (correct font/weight/size/color) instead of damped CSS. Recreated on
  // rebuild so removed cards' surfaces are torn down.
  let galleryChild: ReturnType<typeof editor.createChild> | null = null;
  let styleMgr: StyleManagerHandle | null = null;
  let devPanel: DevPanelHandle | null = null;
  let pageLayoutDlg: PageLayoutHandle | null = null;
  let sizePosDlg: ShapeSizePositionHandle | null = null;
  let organizeDlg: OrganizePagesHandle | null = null;
  let symbolPicker: SymbolPickerHandle | null = null;

  // Size & Position dialog (B2), shared by the ribbon Shape group and the shape
  // right-click menu (via the onEditShapeSizePosition option hook). Toggles: a
  // second trigger while it's open closes it.
  openShapeSizePosition = (): void => {
    if (sizePosDlg) { sizePosDlg.close(); sizePosDlg = null; return; }
    if (!editor.getSelectedShape()) return;
    sizePosDlg = showShapeSizePosition({ editor, onClose: () => { sizePosDlg = null; } });
  };
  // "Show only styles in use" filter. Since import now keeps every defined style
  // (so authored styles round-trip), a heavy imported doc can crowd the gallery —
  // this collapses it to styles actually applied in the document.
  let hideUnusedStyles = false;
  /** Style ids applied anywhere in the document (paragraph namedStyle + run charStyleId). */
  const usedStyleIds = (): Set<string> => {
    const used = new Set<string>();
    for (const p of paragraphsOf(editor.getDocument())) {
      if (p.style.namedStyle) used.add(p.style.namedStyle);
      for (const r of p.runs) if (r.style.charStyleId) used.add(r.style.charStyleId);
    }
    return used;
  };
  const addStyleCard = (id: string, name: string): HTMLButtonElement => {
    const c = el("button", "style-card");
    const isChar = (() => { const st = styleById(stylesheet(), id); return !!st && styleType(st) === "character"; })();
    c.title = isChar ? `${name} (character style)` : name;
    c.innerHTML = `<span class="preview"></span><span class="name"></span>`;
    (c.querySelector(".name") as HTMLElement).textContent = isChar ? `${name} ⓐ` : name; // ⓐ marks a character style
    const prev = c.querySelector(".preview") as HTMLElement;
    prev.textContent = "AaBbCc"; // plain-text fallback shown until the canvas sample paints (V5)
    galleryChild?.render(prev, { kind: "styleSample", styleId: id, sampleText: "AaBbCc" }, { fit: true, widthPx: 70, maxHeightPx: 24 });
    c.addEventListener("mousedown", (e) => e.preventDefault());
    c.addEventListener("click", () => {
      const st = styleById(stylesheet(), id);
      editor.dispatch(st && styleType(st) === "character" ? applyCharStyle(id) : applyNamedStyle(id));
      editor.focus();
    });
    styleCards.set(id, c);
    styleGallery.appendChild(c);
    return c;
  };
  const rebuildStyleGallery = (): void => {
    galleryChild?.destroy();
    galleryChild = editor.createChild();
    styleGallery.textContent = "";
    styleCards.clear();
    const sheet = stylesheet();
    // When filtering, always keep the default style and the caret's current style
    // visible so the gallery never hides what's active.
    const used = hideUnusedStyles ? usedStyleIds() : null;
    const curId = editor.currentFormat().styleId;
    const keep = (id: string): boolean => !used || used.has(id) || id === sheet.defaultStyleId || id === curId;
    for (const s of sheet.styles) if (keep(s.id)) addStyleCard(s.id, s.name);
  };
  rebuildStyleGallery();
  controls.appendChild(styleGallery);
  ribRegister(styleGallery, "gallery");
  btn(ICONS.stylePencil, "Update current style to match selection", () => {
    const id = editor.currentFormat().styleId;
    if (id) editor.dispatch(updateStyleToSelection(id));
  });
  const openStyleManager = (sel?: { kind: "paragraph" | "character"; id: string }): void => {
    if (styleMgr) { styleMgr.refresh(); return; }
    styleMgr = showStyleManager({
      editor,
      ...(sel ? { initialSelection: sel } : {}),
      onClose: () => { styleMgr = null; },
    });
  };
  btn(ICONS.styleNew, "Manage styles…", () => {
    const id = editor.currentFormat().styleId;
    openStyleManager(id ? { kind: "paragraph", id } : undefined);
  });
  const filterBtn = btn(ICONS.filter, "Show only styles in use", () => {
    hideUnusedStyles = !hideUnusedStyles;
    filterBtn.classList.toggle("active", hideUnusedStyles);
    filterBtn.title = hideUnusedStyles ? "Show all styles" : "Show only styles in use";
    rebuildStyleGallery();
  });

  // ---- Editing ----
  group(home, "Editing");
  {
    const col = el("div");
    col.style.cssText = "display:flex;flex-direction:column;gap:1px;";
    const prev = controls;
    controls = col;
    const wide = "width:100%;justify-content:flex-start;gap:4px;";
    const ed = (icon: string, label: string, title: string, onClick: () => void): void => {
      const b = btn(icon + `<span>${label}</span>`, title, onClick);
      b.style.cssText = wide;
    };
    ed(ICONS.find, "Find", "Find & replace (Ctrl+F)", () => openFind());
    ed(ICONS.replace, "Replace", "Replace (Ctrl+F)", () => openFind());
    ed(ICONS.select, "Select All", "Select all (Ctrl+A)", () => editor.selectAll());
    controls = prev;
    controls.appendChild(col);
  }

  // ===== Insert tab ========================================================
  const insert = tab("insert", "Insert");
  group(insert, "Pages");
  btn(ICONS.pageBreak, "Page break (Ctrl+Enter)", () => {
    editor.dispatch(insertPageBreak());
    editor.focus();
  });
  btn(ICONS.sectionBreak, "Section break — next page", () => {
    editor.dispatch(insertSectionBreak());
    editor.focus();
  });
  group(insert, "Tables");
  const insTableBtn = btn(ICONS.table, "Insert table", () => {}, true);
  insTableBtn.addEventListener("click", () => tableGridPopover(insTableBtn));
  group(insert, "Illustrations");
  // Pick an image from the device, register its bytes in the shared media store
  // (content-addressed, so it persists + exports), then insert it at the caret at
  // its natural size (capped to a sensible width). Both dispatches run — only the
  // one matching the caret context (body vs table cell) applies.
  const pickAndInsertImage = (): void => {
    const input = el("input");
    input.type = "file";
    input.accept = "image/*";
    input.addEventListener("change", () => {
      const file = input.files?.[0];
      if (!file) return;
      void (async () => {
        const bytes = new Uint8Array(await file.arrayBuffer());
        const mime = file.type || "image/png";
        const mediaId = await registerMediaBytes(bytes, mime);
        const src = mediaUrl(mediaId);
        if (!src) return;
        let w = 320;
        let h = 200;
        try {
          const bmp = await createImageBitmap(new Blob([bytes], { type: mime }));
          w = bmp.width;
          h = bmp.height;
          bmp.close();
        } catch {
          /* keep fallback dimensions */
        }
        const MAX_W = 480; // don't overflow the page on a large photo
        if (w > MAX_W) {
          h = Math.round((h * MAX_W) / w);
          w = MAX_W;
        }
        editor.dispatch(insertImage(src, w, h, mediaId)); // top-level caret
        editor.dispatch(insertImageInCell(src, w, h, mediaId)); // table-cell caret
        editor.focus();
      })();
    });
    input.click();
  };
  btn(ICONS.image, "Insert image from your device", () => pickAndInsertImage());
  const insShapeBtn = btn(ICONS.shapes, "Insert a shape", () => {}, true);
  insShapeBtn.addEventListener("click", () => shapesPopover(insShapeBtn));
  // Picture Tools + Shape Tools are contextual tabs — they appear only when an
  // image / shape is selected, replacing the always-disabled image & shape
  // buttons that used to sit greyed on the Insert tab (critique C5). "Insert a
  // shape" stays on Insert (above); acting on a shape moves to Shape Tools.
  const pictureTab = tab("picture", "Picture Tools", undefined, (f) => f.imageSelected);
  const shapeTab = tab("shape", "Shape Tools", undefined, (f) => f.shapeSelected || f.multipleShapesSelected || f.groupSelected);
  group(pictureTab, "Picture"); // acts on the selected image
  enable(
    btn(ICONS.wrapSquare, "Wrap text around image (square)", () => {
      const id = editor.getSelectedObject();
      if (id) editor.dispatch(setImageProps(id, { wrap: "square", align: "left" }));
    }),
    (f) => f.imageSelected,
    "select an image first",
  );
  enable(
    btn(ICONS.wrapInline, "Image in line with text (block)", () => {
      const id = editor.getSelectedObject();
      if (id) editor.dispatch(setImageProps(id, { wrap: "block", align: "center" }));
    }),
    (f) => f.imageSelected,
    "select an image first",
  );
  group(shapeTab, "Shape"); // acts on the selected shape
  const shapeFillBtn = btn(ICONS.shapeFill, "Shape fill", () => shapeFillPopover(shapeFillBtn), true);
  enable(shapeFillBtn, (f) => f.shapeSelected, "select a shape first");
  const shapeOutlineBtn = btn(ICONS.outline, "Shape outline (colour, width, dash)", () => shapeOutlinePopover(shapeOutlineBtn), true);
  enable(shapeOutlineBtn, (f) => f.shapeSelected, "select a shape first");
  // Add / edit the shape's text box body — on the ribbon too, not just the floating
  // toolbar and right-click menu (issue #244 C3). No-ops for a cell-nested shape.
  enable(
    btn(ICONS.shapeTextBox, "Add or edit text in shape", () => {
      editor.addTextToSelectedShape();
    }),
    (f) => f.shapeSelected,
    "select a shape first",
  );
  enable(
    btn(ICONS.wrapSquare, "Wrap text around shape (square)", () => {
      const id = editor.getSelectedObject();
      if (id) { editor.dispatch(setShapeProps(id, { wrap: "square", align: "left", anchor: null })); editor.focus(); }
    }),
    (f) => f.shapeSelected,
    "select a shape first",
  );
  enable(
    btn(ICONS.wrapInline, "Shape in line with text (block)", () => {
      const id = editor.getSelectedObject();
      if (id) { editor.dispatch(setShapeProps(id, { wrap: "block", align: "center", anchor: null })); editor.focus(); }
    }),
    (f) => f.shapeSelected,
    "select a shape first",
  );
  // Layer-vs-text (behind / in front of the body text) — distinct from the z-order
  // buttons below; exposed on the ribbon so it's reachable from every surface, and
  // labelled to disambiguate the two "depth" concepts (issue #244 C2).
  enable(
    btn(ICONS.behindText, "Move shape behind text", () => {
      const id = editor.getSelectedObject();
      if (id) { editor.dispatch(setShapeLayer(id, true)); editor.focus(); }
    }),
    (f) => f.shapeSelected,
    "select a shape first",
  );
  enable(
    btn(ICONS.inFrontText, "Move shape in front of text", () => {
      const id = editor.getSelectedObject();
      if (id) { editor.dispatch(setShapeLayer(id, false)); editor.focus(); }
    }),
    (f) => f.shapeSelected,
    "select a shape first",
  );
  enable(
    btn(ICONS.bringFront, "Bring shape to front (among shapes)", () => {
      const id = editor.getSelectedObject();
      if (id) { editor.dispatch(bringShapeToFront(id)); editor.focus(); }
    }),
    (f) => f.shapeSelected,
    "select a shape first",
  );
  enable(
    btn(ICONS.sendBack, "Send shape to back (among shapes)", () => {
      const id = editor.getSelectedObject();
      if (id) { editor.dispatch(sendShapeToBack(id)); editor.focus(); }
    }),
    (f) => f.shapeSelected,
    "select a shape first",
  );
  enable(
    btn(ICONS.ruler, "Size & position (exact width, height, rotation, offset)", () => openShapeSizePosition()),
    (f) => f.shapeSelected,
    "select a shape first",
  );
  enable(
    btn(ICONS.trash, "Delete shape", () => {
      if (editor.getSelectedObject()) {
        editor.deleteSelectedObject();
        editor.focus();
      }
    }),
    (f) => f.shapeSelected,
    "select a shape first",
  );
  // F5 group authoring (issue #244) — Group ≥2 selected shapes / Ungroup a group.
  enable(
    btn(ICONS.group, "Group shapes", () => { editor.groupShapes(); editor.focus(); }),
    (f) => f.multipleShapesSelected,
    "Shift-click a second shape first",
  );
  enable(
    btn(ICONS.ungroup, "Ungroup shape", () => { editor.ungroupShape(); editor.focus(); }),
    (f) => f.groupSelected,
    "select a group first",
  );
  group(insert, "Equation");
  let equationEditorLoading = false; // debounce double-clicks while the chunk loads
  txtBtn("√x", "Insert equation (MathML)", () => {
    if (equationEditorLoading) return;
    equationEditorLoading = true;
    import("./ui/equationEditor")
      .then(({ showEquationEditor }) => {
        if (teardown.signal.aborted) return;
        showEquationEditor({
          onApply: (eq) => {
            editor.dispatch(eq.display ? insertEquation(eq) : insertInlineEquation(eq));
            editor.focus();
          },
        });
      })
      .catch((e: unknown) => console.warn("[wordcanvas] failed to load the equation editor:", e))
      .finally(() => { equationEditorLoading = false; });
  }, "font-family:Georgia,serif;font-style:italic;");

  group(insert, "Symbols");
  // Single floating picker (toggle on re-click); closed on teardown so its
  // backdrop + document-level Escape listener never leak. Mirrors pageLayoutDlg.
  let symbolPickerLoading = false; // debounce double-clicks while the chunk loads
  btn(ICONS.symbol, "Insert symbol or special character", () => {
    if (symbolPicker) { symbolPicker.close(); return; }
    if (symbolPickerLoading) return;
    symbolPickerLoading = true;
    import("./ui/symbolPicker")
      .then(({ showSymbolPicker }) => {
        if (teardown.signal.aborted || symbolPicker) return;
        symbolPicker = showSymbolPicker({
          onPick: (font, char) => {
            editor.dispatch(insertSymbolCmd(font, char));
            editor.focus();
          },
          onClose: () => { symbolPicker = null; },
        });
      })
      .catch((e: unknown) => console.warn("[wordcanvas] failed to load the symbol picker:", e))
      .finally(() => { symbolPickerLoading = false; });
  });
  teardown.signal.addEventListener("abort", () => symbolPicker?.close(), { once: true });

  group(insert, "Links");
  const insLinkBtn = btn(ICONS.link, "Insert/remove hyperlink", () => {});
  insLinkBtn.addEventListener("click", () => linkDialog(insLinkBtn));
  group(insert, "References");
  btn(ICONS.toc, "Insert / update table of contents (Ctrl+click an entry jumps to it)", () => {
    editor.dispatch(insertTocCmd());
    editor.focus();
  });
  btn(ICONS.tocRefresh, "Recalculate TOC page numbers from the current layout", () => {
    const n = editor.recalculateToc();
    console.log(n > 0 ? `[toc] updated ${n} page number${n === 1 ? "" : "s"}` : "[toc] page numbers already current");
    editor.focus();
  });
  txtBtn("ab¹", "Insert footnote", () => {
    editor.dispatch(insertFootnoteCmd());
    editor.focus();
  }, "font-size:11px;");
  txtBtn("abⁱ", "Insert endnote", () => {
    editor.dispatch(insertEndnoteCmd());
    editor.focus();
  }, "font-size:11px;");
  group(insert, "Controls");
  btn(ICONS.sdtText, "Rich text content control (wraps the selection or selected image)", () => {
    // A selected image is an object selection (no text caret), so route it to the
    // image-wrapping command; otherwise wrap the text selection/caret as before.
    const imgId = editor.getSelectedObject();
    if (imgId) {
      editor.dispatch(wrapImageInContentControl(imgId, "richText", { alias: "Text" }));
      // Keep the image selected so the new control frame is visible — don't refocus
      // the doc (which would drop the object selection).
    } else {
      editor.dispatch(insertContentControl("richText", { alias: "Text" }));
      editor.focus();
    }
  });
  btn(ICONS.sdtCheckbox, "Check box content control", () => {
    editor.dispatch(insertContentControl("checkbox", { alias: "Check Box" }));
    editor.focus();
  });
  btn(ICONS.sdtDropdown, "Drop-down list content control", () => {
    showInputDialog({
      title: "Drop-down list",
      label: "List items (comma-separated)",
      initial: "Yes, No, N/A",
      okLabel: "Insert",
      validate: (v) => (v.split(",").some((s) => s.trim().length > 0) ? null : "Enter at least one item."),
      onSubmit: (raw) => {
        const listItems = raw
          .split(",")
          .map((s) => s.trim())
          .filter((s) => s.length > 0)
          .map((s) => ({ display: s, value: s }));
        editor.dispatch(insertContentControl("dropDown", { alias: "Drop-Down List", listItems }));
        editor.focus();
      },
    });
  });
  btn(ICONS.sdtDate, "Date picker content control", () => {
    editor.dispatch(insertContentControl("date", { alias: "Date", dateFormat: "M/d/yyyy" }));
    editor.focus();
  });
  enable(
    btn(ICONS.sdtProps, "Content control properties & content (inspect the control at the caret)", () => {
      if (!editor.inspectContentControl()) showNotice("Place the caret inside a content control first.", "info");
    }),
    (f) => f.inContentControl,
    "place the caret in a content control",
  );
  enable(
    btn(ICONS.sdtRemove, "Remove the content control at the caret or around the selected image (keeps its content)", () => {
      const id = editor.activeContentControlId();
      if (id) editor.dispatch(removeContentControl(id, false));
      // Don't force focus to the doc — that would drop an active image selection.
    }),
    (f) => f.inContentControl,
    "place the caret in a content control",
  );

  // ===== Layout tab ========================================================
  const layout = tab("layout", "Layout");
  group(layout, "Page Setup");
  btn(ICONS.pageSetup, "Page layout (size, orientation, margins, columns, header/footer distance, page color & borders — applies to the caret's section)", () => {
    if (pageLayoutDlg) { pageLayoutDlg.close(); pageLayoutDlg = null; return; }
    pageLayoutDlg = showPageLayout({ editor, onClose: () => { pageLayoutDlg = null; } });
  });
  if (config.organizePages) {
    group(layout, "Pages");
    btn(ICONS.organizePages, "Organize pages — drag page thumbnails to reorder whole sections (never splits content)", () => {
      if (organizeDlg) { organizeDlg.close(); organizeDlg = null; return; }
      organizeDlg = showOrganizePages({
        editor,
        fontRegistry,
        theme: config.theme,
        onClose: () => { organizeDlg = null; },
      });
    });
  }

  // ===== Table Tools — contextual tab (only while the caret is in a table) ==
  const tableTab = tab("table", "Table Tools", undefined, (f) => f.inTable);
  group(tableTab, "Rows & Columns");
  btn(ICONS.rowAbove, "Insert row above", () => editor.dispatch(insertTableRowCmd("above")));
  btn(ICONS.rowBelow, "Insert row below", () => editor.dispatch(insertTableRowCmd("below")));
  btn(ICONS.colLeft, "Insert column left", () => editor.dispatch(insertTableColumnCmd("left")));
  btn(ICONS.colRight, "Insert column right", () => editor.dispatch(insertTableColumnCmd("right")));
  btn(ICONS.deleteRow, "Delete row", () => editor.dispatch(deleteTableRowCmd()));
  btn(ICONS.deleteCol, "Delete column", () => editor.dispatch(deleteTableColumnCmd()));
  btn(ICONS.deleteTable, "Delete table", () => editor.dispatch(deleteTableCmd()));
  group(tableTab, "Merge");
  btn(ICONS.mergeCells, "Merge cells (select across cells in one row)", () => editor.dispatch(mergeCellsCmd()));
  btn(ICONS.unmergeCells, "Unmerge cell", () => editor.dispatch(unmergeCellCmd()));
  group(tableTab, "Size");
  const autofitBtn = txtBtn("AutoFit", "AutoFit columns to contents or window, or use fixed widths", () => {}, "", true);
  autofitBtn.addEventListener("click", () => {
    // Tick the active mode (resolved against the table the caret is in right now).
    const tbl = tableAtSelection({ doc: editor.getDocument(), selection: editor.getSelection() });
    const cur = tbl?.widthMode ?? "fixed";
    const pref = tbl?.preferredWidth;
    const curAlign = tbl?.align ?? "left";
    const fit = (text: string, mode: "fixed" | "autofitContents" | "autofitWindow") => ({
      label: cur === mode ? `${text}  ✓` : text,
      onClick: () => { editor.dispatch(setTableWidthModeAtSelectionCmd(mode)); editor.focus(); },
    });
    const widthItem = (text: string, value: number | null) => ({
      label:
        (value === null ? !pref : pref?.type === "pct" && Math.round(pref.value) === value) ? `${text}  ✓` : text,
      onClick: () => {
        editor.dispatch(setTablePreferredWidthAtSelectionCmd(value === null ? null : { type: "pct", value }));
        editor.focus();
      },
    });
    const alignItem = (text: string, a: "left" | "center" | "right") => ({
      label: curAlign === a ? `${text}  ✓` : text,
      onClick: () => { editor.dispatch(setTableAlignAtSelectionCmd(a)); editor.focus(); },
    });
    openPop(
      autofitBtn,
      menu([
        fit("AutoFit to Contents", "autofitContents"),
        fit("AutoFit to Window", "autofitWindow"),
        fit("Fixed Column Width", "fixed"),
        widthItem("Width: 25%", 25),
        widthItem("Width: 50%", 50),
        widthItem("Width: 75%", 75),
        widthItem("Width: Full page", null),
        alignItem("Align Left", "left"),
        alignItem("Align Center", "center"),
        alignItem("Align Right", "right"),
      ]),
    );
  });

  // ===== View tab ==========================================================
  const view = tab("view", "View");
  group(view, "Navigator");
  let outlineToggle = (): void => {};
  // One Navigator (row 17): the rail switches Headings/Pages/Objects/Marks. These
  // two are just shortcuts to open it (on Headings, or on the Marks/bookmarks tab).
  const outlineBtn = btn(ICONS.outline, "Navigator — headings, pages, objects & marks", () => outlineToggle());
  btn(ICONS.bookmark, "Bookmarks (Navigator ▸ Marks)", () => toggleBookmarks());
  const rulerBtn = btn(ICONS.ruler, "Horizontal ruler", () => rulerBtn.classList.toggle("active", toggleRuler()));
  rulerBtn.classList.toggle("active", showHRuler);
  const vrulerBtn = btn(ICONS.rulerV, "Vertical ruler", () => vrulerBtn.classList.toggle("active", toggleVRuler()));
  vrulerBtn.classList.toggle("active", showVRuler);
  const gridBtn = btn(ICONS.grid, "Show grid — a light mesh for aligning objects", () => {
    gridView.show = !gridView.show;
    editor.setShowGrid(gridView.show);
    gridBtn.classList.toggle("active", gridView.show);
  });
  gridBtn.classList.toggle("active", gridView.show);
  const snapBtn = btn(ICONS.snap, "Snap to grid — anchored objects snap to grid lines while dragging", () => {
    gridView.snap = !gridView.snap;
    editor.setSnapToGrid(gridView.snap);
    snapBtn.classList.toggle("active", gridView.snap);
  });
  snapBtn.classList.toggle("active", gridView.snap);
  const gridSpacingSel = select("Grid spacing", 92);
  for (const [px, label] of [
    [12, 'Grid: 1/8"'],
    [24, 'Grid: 1/4"'],
    [48, 'Grid: 1/2"'],
  ] as [number, string][]) {
    opt(gridSpacingSel, String(px), label);
  }
  gridSpacingSel.value = String(gridView.spacingPx);
  gridSpacingSel.addEventListener("change", () => {
    gridView.spacingPx = Number(gridSpacingSel.value);
    editor.setGridSpacing(gridView.spacingPx);
  });
  marksToggleBtn();
  if (online) btn(ICONS.activity, "Activity — who created/edited this document and when", () => toggleActivity());

  // ===== Developer tab (develop mode only) =================================
  // Gated on the `develop` config flag: the tab only EXISTS when the embedder
  // opts in, and even then nothing dev-related runs until the developer clicks
  // "Inspect document tree" to open the floating Document-tree inspector.
  if (config.develop) {
    const dev = tab("developer", "Developer");
    group(dev, "Inspect");
    const toggleDevPanel = (): void => {
      if (devPanel) { devPanel.close(); return; } // toggle: a second click closes it
      devPanel = showDevPanel({
        editor,
        setDocument: (next) => replaceDocument(next),
        onClose: () => {
          devPanel = null;
          refreshDevPanel = () => {};
          inspectorHoverSink = () => {};
          inspectorProbeSink = () => {};
          closeDevPanel = () => {};
          devBtn.classList.remove("active");
        },
      });
      refreshDevPanel = () => devPanel?.refresh();
      inspectorHoverSink = (blockId) => devPanel?.highlightNode(blockId);
      inspectorProbeSink = (probe) => devPanel?.setProbe(probe);
      closeDevPanel = () => devPanel?.close();
      devBtn.classList.add("active");
    };
    const devBtn = btn(
      ICONS.devtools + "<span>Inspect document tree</span>",
      "Open the Document-tree inspector — browse the parsed model, highlight nodes on the page, and read each node's JSON",
      toggleDevPanel,
    );
    devBtn.style.cssText = "width:100%;justify-content:flex-start;gap:4px;";

    // Reverse builder: emit DocumentBuilder TypeScript that reconstructs the
    // current document, so a visually-authored doc becomes an editable code
    // template. Develop-mode only — a developer-facing tool.
    const codeBtn = btn(
      ICONS.devtools + "<span>Export builder code</span>",
      "Generate DocumentBuilder TypeScript that reconstructs this document and download it as a .ts template",
      () => {
        try {
          const { code, uncovered } = emitBuilderCode(editor.getDocument());
          const url = URL.createObjectURL(new Blob([code], { type: "text/plain;charset=utf-8" }));
          const a = el("a");
          a.href = url;
          a.download = "template.ts";
          a.click();
          setTimeout(() => URL.revokeObjectURL(url), 1000);
          if (uncovered.length > 0) {
            console.warn("[export-builder-code] fields not expressible via the fluent API:", uncovered);
            showNotice(`Builder code exported. ${uncovered.length} field(s) couldn't be reproduced through the builder API — see the console for the list.`, "warning");
          } else {
            showNotice("Builder code exported to template.ts.", "info");
          }
        } catch (err) {
          console.error("[export-builder-code] failed", err);
          showNotice("Couldn't export builder code — see the console.", "warning");
        }
      },
    );
    codeBtn.style.cssText = "width:100%;justify-content:flex-start;gap:4px;";
  }

  // ---- Review controls live in the ribbon HEADER (right of the tab strip,
  //      by the collapse button) so the mode switch + pane toggle are reachable
  //      from any tab — not buried inside the View tab.
  const MODE_LABELS: Array<[EditMode, string]> = [["edit", "Editing"], ["suggest", "Suggesting"], ["view", "Viewing"]];
  const allowedModesUi = runtime.allowedModes;
  const headerReview = el("div", "cw-header-review");
  const modeSel = el("select", "cw-mode-select");
  modeSel.title = "Editing mode";
  for (const [v, l] of MODE_LABELS) {
    if (allowedModesUi && !allowedModesUi.includes(v)) continue;
    const o = el("option");
    o.value = v;
    o.textContent = l;
    modeSel.appendChild(o);
  }
  modeSel.value = editor.getMode();
  modeSel.addEventListener("change", () => {
    if (!editor.setMode(modeSel.value as EditMode)) modeSel.value = editor.getMode(); // rejected → revert
    editor.focus();
  });
  syncMode = (): void => {
    modeSel.value = editor.getMode();
  };
  const inspectorBtn = el("button", "cw-header-btn");
  inspectorBtn.textContent = "Inspector";
  inspectorBtn.title = "Inspector — edit the selection's font & paragraph properties live";
  inspectorBtn.addEventListener("mousedown", (e) => e.preventDefault());
  inspectorBtn.addEventListener("click", () => toggleInspector());
  const reviewBtn = el("button", "cw-header-btn");
  reviewBtn.textContent = "Review";
  reviewBtn.title = "Suggestions & comments — review, accept, reject";
  reviewBtn.addEventListener("mousedown", (e) => e.preventDefault());
  reviewBtn.addEventListener("click", () => toggleReview());
  headerReview.append(modeSel, inspectorBtn, reviewBtn);
  // Tab-strip overflow (critique R3): when the tabs don't fit, this button appears
  // and opens a menu of them, so Table/View/Developer never silently vanish.
  const tabOverflowBtn = el("button", "rib-tab-overflow");
  tabOverflowBtn.textContent = "⋯";
  tabOverflowBtn.title = "More tabs";
  tabOverflowBtn.style.display = "none";
  tabOverflowBtn.addEventListener("mousedown", (e) => e.preventDefault());
  tabOverflowBtn.addEventListener("click", () => {
    const items: { label: string; onClick: () => void }[] = [];
    for (const [id, b] of tabButtons) {
      if (b.style.display === "none") continue; // skip hidden contextual tabs
      items.push({ label: b.textContent ?? id, onClick: () => showTab(id) });
    }
    openPop(tabOverflowBtn, menu(items));
  });
  tabsBar.appendChild(tabOverflowBtn);
  tabsBar.appendChild(headerReview);

  // ---- Document identity + live save state (pinned top-left of the ribbon) ---
  // Answers "what is this file, and is my work safe?" without opening a menu.
  // The model has no title field, so the title is presentational: an explicit
  // `documentTitle`, else the opened .docx filename, else "Untitled document".
  const docIdent = el("div", "cw-doc-ident");
  const titleEl = el("span", "cw-doc-title");
  const saveWrap = el("span", "cw-save");
  const saveDot = el("span", "cw-save-dot");
  const saveText = el("span", "cw-save-text");
  saveWrap.append(saveDot, saveText);
  docIdent.append(titleEl, saveWrap);
  tabsBar.insertBefore(docIdent, tabScroll); // leftmost, before the scrolling tabs

  // ---- Import fidelity badge (Move 3): the moat made visible — but only its
  // WARNING half. "⚠ N notes" appears when a .docx import preserved-but-adapted
  // something, expanding to a plain-language list. A clean import (or no import at
  // all) shows NOTHING: the badge is not in the DOM, so it doesn't even occupy
  // header width. The old permanent "✓ Word-faithful" success state was dropped by
  // design — a standing success indicator asserts the default assumption still
  // holds, which trains the eye to ignore this region and makes the warning state
  // less noticeable, not more (and it was unearned for never-imported documents).
  const fidelityBtn = el("button", "cw-fidelity");
  fidelityBtn.type = "button";
  fidelityBtn.setAttribute("aria-haspopup", "true");
  fidelityBtn.setAttribute("aria-expanded", "false");
  // NOT appended here — refreshFidelity attaches it only in the warning state.
  let fidelitySurface: SurfaceHandle | null = null;
  let fidelityPanelEl: HTMLElement | null = null;
  let fidelityOutside: ((e: MouseEvent) => void) | null = null;
  const closeFidelityPanel = (): void => {
    if (fidelityOutside) { window.removeEventListener("mousedown", fidelityOutside, true); fidelityOutside = null; }
    fidelitySurface?.release();
    fidelitySurface = null;
    fidelityPanelEl?.remove();
    fidelityPanelEl = null;
    fidelityBtn.setAttribute("aria-expanded", "false");
  };
  const openFidelityPanel = (): void => {
    // The badge only exists in the warning state, so the panel is only ever opened
    // with warnings present — the old "✓ Word-faithful" clean-state branch is gone.
    const warnings = importWarnings ?? [];
    const n = warnings.length;
    if (n === 0) return;
    const panel = el("div", "cw-fidelity-panel");
    const head = el("div", "cw-fidelity-head");
    head.textContent = `⚠ ${n} feature${n === 1 ? "" : "s"} preserved but adapted`;
    const sub = el("div", "cw-fidelity-sub");
    sub.textContent =
      "These were kept when the document opened, but adapted or simplified for editing here. Everything else round-trips faithfully.";
    panel.append(head, sub);
    const ul = el("ul", "cw-fidelity-list");
    for (const w of warnings) {
      const li = el("li");
      li.textContent = w.message;
      ul.append(li);
    }
    panel.append(ul);
    const r = fidelityBtn.getBoundingClientRect();
    panel.style.left = `${Math.round(Math.min(r.left, window.innerWidth - 340))}px`;
    panel.style.top = `${Math.round(r.bottom + 5)}px`;
    document.body.append(panel);
    fidelityPanelEl = panel;
    fidelityBtn.setAttribute("aria-expanded", "true");
    fidelitySurface = openSurface({ el: panel, close: closeFidelityPanel, kind: "overlay" });
    // Deferred a tick so the opening click doesn't immediately self-close.
    fidelityOutside = (e: MouseEvent): void => {
      const t = e.target as Node;
      if (panel.contains(t) || fidelityBtn.contains(t)) return;
      closeFidelityPanel();
    };
    setTimeout(() => { if (fidelityOutside) window.addEventListener("mousedown", fidelityOutside, true); }, 0);
  };
  fidelityBtn.addEventListener("click", () => {
    if (fidelityPanelEl) closeFidelityPanel();
    else openFidelityPanel();
  });
  refreshFidelity = (): void => {
    const n = importWarnings ? importWarnings.length : 0;
    // Clean import or no import at all → render nothing (detach the element so it
    // occupies no header width). Only the warning state is ever visible.
    if (n === 0) {
      fidelityBtn.remove();
      if (fidelityPanelEl) closeFidelityPanel();
      return;
    }
    if (!fidelityBtn.isConnected) docIdent.append(fidelityBtn);
    fidelityBtn.className = "cw-fidelity notes";
    fidelityBtn.textContent = `⚠ ${n} note${n === 1 ? "" : "s"}`;
    fidelityBtn.title = `${n} feature${n === 1 ? " was" : "s were"} preserved but adapted on import. Click for details.`;
    if (fidelityPanelEl) { closeFidelityPanel(); } // stale content — reopen on demand
  };
  refreshFidelity(); // initial state: no import yet → nothing rendered

  // Where does a save actually go? Host pipeline > online autosync > nowhere.
  const saveTarget: "host" | "online" | "none" = runtime.onSave ? "host" : online ? "online" : "none";
  let docTitle = runtime.documentTitle?.trim() || "Untitled document";
  let savedHead = editor.getChangeHead(); // "clean" baseline (changes recorded since load)
  let lastSavedAt: number | null = null;
  let lastDownloadAt: number | null = null;
  let saving = false;

  const stripExt = (n: string): string => n.replace(/\.[^.]+$/, "").trim() || n;
  const fmtTime = (ts: number): string => new Date(ts).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
  const isDirty = (): boolean => editor.getChangeHead() !== savedHead;
  const applyTitle = (): void => {
    titleEl.textContent = docTitle;
    titleEl.title = docTitle;
  };

  refreshSaveState = (): void => {
    const dirty = isDirty();
    let cls: string;
    let text: string;
    let tip: string;
    if (saving) {
      cls = "saving"; text = "Saving…"; tip = "Saving your document…";
    } else if (saveTarget === "online") {
      // Online continuously streams every edit to the server — genuinely saved.
      cls = "saved"; text = "Saved"; tip = "Changes are saved automatically to the server.";
    } else if (saveTarget === "host") {
      if (dirty) { cls = "dirty"; text = "Unsaved changes"; tip = "You have unsaved changes. Press Ctrl+S to save."; }
      else if (lastSavedAt != null) { cls = "saved"; text = `Saved ${fmtTime(lastSavedAt)}`; tip = "All changes saved."; }
      else { cls = "clean"; text = "No unsaved changes"; tip = "Press Ctrl+S to save."; }
    } else {
      // "none": nothing persists this session — be honest, never claim "Saved".
      if (dirty) { cls = "dirty"; text = "Unsaved changes"; tip = "This document isn't connected to storage. Press Ctrl+S to download a .docx copy."; }
      else if (lastDownloadAt != null) { cls = "clean"; text = `Copy downloaded ${fmtTime(lastDownloadAt)}`; tip = "A .docx copy was downloaded. The live document still isn't connected to storage."; }
      else { cls = "clean"; text = "Not saved"; tip = "This document isn't connected to storage. Press Ctrl+S to download a .docx copy."; }
    }
    saveDot.className = "cw-save-dot " + cls;
    saveText.textContent = text;
    saveWrap.title = tip;
  };
  resetSaveBaseline = (): void => {
    savedHead = editor.getChangeHead();
    lastSavedAt = null;
    lastDownloadAt = null;
    refreshSaveState();
  };
  setDocTitleFromFile = (name: string | null): void => {
    if (name) { docTitle = stripExt(name); applyTitle(); }
  };
  // Downloaded exports are named after the document title (filesystem-safe).
  currentDocBaseName = (): string => docTitle.replace(/[\\/:*?"<>|]+/g, "_").trim() || "document";

  // Ctrl+S / Cmd+S — bound so the browser's "Save page" dialog never fires. With
  // a host onSave it saves through that pipeline; offline it downloads a .docx
  // copy; online it's already autosynced, so it just re-affirms the state.
  const triggerSave = async (): Promise<void> => {
    if (saving) return;
    if (saveTarget === "online") { refreshSaveState(); return; }
    saving = true;
    refreshSaveState();
    try {
      await exportAs("docx"); // honors runtime.onSave; else triggers the .docx download
      savedHead = editor.getChangeHead();
      if (saveTarget === "host") lastSavedAt = Date.now();
      else lastDownloadAt = Date.now();
    } catch (e) {
      console.error("[save]", e);
    } finally {
      saving = false;
      refreshSaveState();
    }
  };
  window.addEventListener(
    "keydown",
    (e) => {
      if (!(e.ctrlKey || e.metaKey) || e.altKey || e.key.toLowerCase() !== "s") return;
      e.preventDefault(); // suppress the browser Save-Page dialog in every mode
      void triggerSave();
    },
    { signal: teardown.signal },
  );
  applyTitle();
  refreshSaveState();

  // ---- Quick-access toolbar: undo / redo / save, always visible on any tab ---
  // Undo & Redo are the two most-used commands in any editor; burying them in
  // the File tab (critique C1) put them two clicks deep and invisible from the
  // default tab. Surface them (plus Save) right after the document title.
  const qat = el("div", "cw-qat");
  const qatBtn = (icon: string, title: string, onClick: () => void): HTMLButtonElement => {
    const b = el("button", "cw-qat-btn");
    b.innerHTML = icon;
    b.title = title;
    b.addEventListener("mousedown", (e) => e.preventDefault()); // keep editor focus
    b.addEventListener("click", () => {
      onClick();
      editor.focus();
    });
    qat.appendChild(b);
    return b;
  };
  const undoQat = qatBtn(ICONS.undo, "Undo (Ctrl+Z)", () => editor.undo());
  const redoQat = qatBtn(ICONS.redo, "Redo (Ctrl+Y)", () => editor.redo());
  qatBtn(ICONS.save, "Save (Ctrl+S)", () => void triggerSave());
  tabsBar.insertBefore(qat, tabScroll); // between the title and the scrolling tabs
  // Grey undo/redo when the stack is empty (kept live via syncQat, below).
  syncQat = (): void => {
    undoQat.disabled = !editor.canUndo();
    redoQat.disabled = !editor.canRedo();
  };

  group(view, "Zoom");
  txtBtn("−", "Zoom out", () => editor.setZoom(editor.getZoom() / config.behavior.zoomStep), "font-size:15px;");
  const zoomSel = select("Zoom level", 66);
  for (const z of [0.5, 0.75, 1, 1.25, 1.5, 2, 3]) opt(zoomSel, String(z), `${Math.round(z * 100)}%`);
  zoomSel.value = "1";
  zoomSel.addEventListener("change", () => editor.setZoom(parseFloat(zoomSel.value)));
  txtBtn("+", "Zoom in", () => editor.setZoom(editor.getZoom() * config.behavior.zoomStep), "font-size:15px;");

  // ---- Activity panel (who created/edited, when) — online only ------------
  if (online) {
    const panel = el("div");
    panel.className = "cw-float-drawer"; // positioning/scoping lives in the class
    panel.style.display = "none";
    const ahead = el("div");
    ahead.style.cssText =
      "flex:0 0 auto;display:flex;align-items:center;justify-content:space-between;padding:10px 12px;border-bottom:1px solid #e1dfdd;font-weight:600;color:#323130;";
    const atitle = el("span");
    atitle.textContent = "Activity";
    const aclose = el("button");
    aclose.textContent = "×";
    aclose.style.cssText = "border:none;background:transparent;font-size:18px;cursor:pointer;color:#605e5c;";
    ahead.append(atitle, aclose);
    const ameta = el("div");
    ameta.style.cssText = "flex:0 0 auto;padding:10px 12px;color:#605e5c;border-bottom:1px solid #f0eeee;";
    const alist = el("div");
    alist.style.cssText = "flex:1 1 auto;overflow-y:auto;padding:2px 0;";
    panel.append(ahead, ameta, alist);
    shell.workarea.appendChild(panel);

    const nameOf = (f: string, l: string): string => `${f} ${l}`.trim() || "Unknown";
    const kindLabel = (origin: string): string =>
      origin === "typing" ? "typed"
      : origin === "paste" ? "pasted"
      : origin === "undo" ? "undo"
      : origin === "redo" ? "redo"
      : "edited";
    const timeAgo = (ts: number): string => {
      const s = Math.max(0, (Date.now() - ts) / 1000);
      if (s < 60) return "just now";
      if (s < 3600) return `${Math.floor(s / 60)}m ago`;
      if (s < 86400) return `${Math.floor(s / 3600)}h ago`;
      return new Date(ts).toLocaleDateString();
    };

    interface ActivityResp {
      createdBy: string | null;
      creatorFirstName: string;
      creatorLastName: string;
      createdAt: number;
      entries: Array<{ firstName: string; lastName: string; ts: number; origin: string; opCount: number }>;
    }
    const render = (a: ActivityResp): void => {
      ameta.textContent = a.createdBy
        ? `Created by ${nameOf(a.creatorFirstName, a.creatorLastName)} · ${timeAgo(a.createdAt)}`
        : "Created locally";
      alist.textContent = "";
      if (a.entries.length === 0) {
        const e = el("div");
        e.style.cssText = "padding:12px;color:#80868b;";
        e.textContent = "No edits yet.";
        alist.appendChild(e);
        return;
      }
      for (const en of a.entries) {
        const row = el("div");
        row.style.cssText = "padding:7px 12px;border-bottom:1px solid #f6f5f5;display:flex;justify-content:space-between;gap:8px;";
        const who = el("span");
        who.style.color = "#323130";
        who.textContent = `${nameOf(en.firstName, en.lastName)} · ${kindLabel(en.origin)}`;
        const when = el("span");
        when.style.cssText = "color:#80868b;white-space:nowrap;";
        when.textContent = timeAgo(en.ts);
        row.append(who, when);
        alist.appendChild(row);
      }
    };

    const fetchActivity = async (): Promise<void> => {
      if (!collabId || !BACKEND_HTTP) {
        ameta.textContent = "Share the document (or open one online) to track activity.";
        alist.textContent = "";
        return;
      }
      try {
        const r = await fetch(`${BACKEND_HTTP}/docs/${encodeURIComponent(collabId)}/activity`);
        if (r.ok) render((await r.json()) as ActivityResp);
      } catch {
        /* ignore */
      }
    };

    // Live refresh: debounced on remote changes (the author's own edits surface
    // via the open-panel poll, since the server doesn't echo them back as remote).
    let refreshTimer: ReturnType<typeof setTimeout> | null = null;
    remoteChangeListeners.push(() => {
      if (refreshTimer) return;
      refreshTimer = setTimeout(() => {
        refreshTimer = null;
        void fetchActivity();
      }, 400);
    });

    let pollTimer: ReturnType<typeof setInterval> | null = null;
    let open = false;
    toggleActivity = (): void => {
      open = !open;
      panel.style.display = open ? "flex" : "none";
      if (open) {
        void fetchActivity();
        pollTimer = setInterval(() => void fetchActivity(), 2500);
      } else if (pollTimer) {
        clearInterval(pollTimer);
        pollTimer = null;
      }
    };
    aclose.addEventListener("click", () => toggleActivity());
    if (runtime.view?.activity) toggleActivity(); // closed by default; embedder can open it
  }

  // ---- Outline drawer (left navigation pane) ------------------------------
  // Lists every Heading-styled paragraph; clicking one moves the caret there
  // and scrolls it into view. The DOM list is rebuilt only when the set of
  // headings actually changes (so typing doesn't reset the drawer's scroll);
  // the active highlight follows the caret on every change.
  const outlineEl = shell.outline;
  if (outlineEl) {
    // ===== Navigator: a 48px icon rail + swappable tab panels (critique S1) ===
    // One panel unifying the Outline, Bookmarks, page organizer and object lists,
    // in one paradigm: Headings · Pages · Objects · Marks.
    outlineEl.classList.add("cw-navigator");
    const rail = el("div", "cw-nav-rail");
    const content = el("div", "cw-nav-content");
    const head = el("div", "outline-head");
    const title = el("span");
    const closeBtn = el("button");
    closeBtn.innerHTML = "×";
    closeBtn.title = "Close navigator";
    closeBtn.setAttribute("aria-label", "Close navigator");
    head.append(title, closeBtn);
    const panels = el("div", "cw-nav-panels");
    content.append(head, panels);
    outlineEl.append(rail, content);
    const mkPanel = (id: string): HTMLElement => {
      const p = el("div", "cw-nav-panel");
      p.dataset["nav"] = id;
      panels.appendChild(p);
      return p;
    };
    const headingsPanel = mkPanel("headings");
    const pagesPanel = mkPanel("pages");
    const objectsPanel = mkPanel("objects");
    const stylesPanel = mkPanel("styles");
    const marksPanel = mkPanel("marks");
    navMarksPanel = marksPanel;
    // Headings tab content: filter + the outline list (rows 9/16).
    const filterWrap = el("div", "cw-outline-filter");
    const filterInput = el("input");
    filterInput.type = "text";
    filterInput.placeholder = "Filter headings…";
    filterInput.setAttribute("aria-label", "Filter headings");
    filterWrap.appendChild(filterInput);
    const list = el("div", "cw-outline-list");
    headingsPanel.append(filterWrap, list);
    // Rail tabs.
    const NAV_LABELS: Record<string, string> = { headings: "Headings", pages: "Pages", objects: "Objects", styles: "Styles", marks: "Marks" };
    const railBtns = new Map<string, HTMLButtonElement>();
    const setNav = (id: string): void => {
      if (id !== "styles") endStylesPreview(); // leaving the Styles tab drops any live preview
      title.textContent = NAV_LABELS[id] ?? "";
      for (const [k, b] of railBtns) b.classList.toggle("active", k === id);
      for (const p of Array.from(panels.children)) (p as HTMLElement).classList.toggle("active", (p as HTMLElement).dataset["nav"] === id);
      if (id === "objects") buildObjects();
      if (id === "pages") buildPages();
      if (id === "styles") buildStyles();
      if (id === "marks") refreshBookmarks();
    };
    navShow = setNav;
    const railBtn = (id: string, icon: string, label: string): void => {
      const b = el("button", "cw-nav-tab");
      b.innerHTML = icon;
      b.title = label;
      b.setAttribute("aria-label", label);
      b.addEventListener("mousedown", (e) => e.preventDefault());
      b.addEventListener("click", () => setNav(id));
      rail.appendChild(b);
      railBtns.set(id, b);
    };
    railBtn("headings", ICONS.outline, "Headings");
    railBtn("pages", ICONS.organizePages, "Pages");
    railBtn("objects", ICONS.shapes, "Objects");
    railBtn("styles", ICONS.styleNew, "Styles");
    railBtn("marks", ICONS.bookmark, "Marks");

    const openOrganize = (): void => {
      if (organizeDlg) { organizeDlg.close(); organizeDlg = null; return; }
      organizeDlg = showOrganizePages({ editor, fontRegistry, theme: config.theme, onClose: () => { organizeDlg = null; } });
    };

    // Objects tab: images / shapes / tables / equations, click to reveal.
    const buildObjects = (): void => {
      objectsPanel.textContent = "";
      const blocks = editor.getDocument().blocks;
      const label = (b: (typeof blocks)[number]): string | null =>
        b.kind === "image" ? "Image" : b.kind === "shape" ? "Shape" : b.kind === "table" ? "Table" : b.kind === "equation" ? "Equation" : null;
      let n = 0;
      const counts: Record<string, number> = {};
      for (const b of blocks) {
        const l = label(b);
        if (!l) continue;
        counts[l] = (counts[l] ?? 0) + 1;
        const row = el("button", "outline-item");
        const lbl = el("span", "outline-label");
        lbl.textContent = `${l} ${counts[l]}`;
        row.appendChild(lbl);
        row.addEventListener("mousedown", (e) => e.preventDefault());
        row.addEventListener("click", () => editor.revealBlock(b.id));
        objectsPanel.appendChild(row);
        n++;
      }
      if (n === 0) {
        const empty = el("div", "outline-empty");
        empty.textContent = "No images, shapes, tables or equations yet.";
        objectsPanel.appendChild(empty);
      }
    };

    // Pages tab: a page list (jump to a page) + the reorder overlay.
    const buildPages = (): void => {
      pagesPanel.textContent = "";
      if (config.organizePages) {
        const act = el("button", "cw-nav-action");
        act.textContent = "Reorder pages…";
        act.addEventListener("mousedown", (e) => e.preventDefault());
        act.addEventListener("click", () => openOrganize());
        pagesPanel.appendChild(act);
      }
      const blocks = editor.getDocument().blocks;
      const lastPage = blocks.length ? editor.getBlockPage(blocks[blocks.length - 1]!.id) ?? 1 : 1;
      const firstBlockOnPage = new Map<number, string>();
      for (const b of blocks) {
        const p = editor.getBlockPage(b.id);
        if (p != null && !firstBlockOnPage.has(p)) firstBlockOnPage.set(p, b.id);
      }
      for (let p = 1; p <= lastPage; p++) {
        const row = el("button", "outline-item");
        const lbl = el("span", "outline-label");
        lbl.textContent = `Page ${p}`;
        row.appendChild(lbl);
        const target = firstBlockOnPage.get(p);
        row.addEventListener("mousedown", (e) => e.preventDefault());
        if (target) row.addEventListener("click", () => editor.revealBlock(target));
        pagesPanel.appendChild(row);
      }
    };
    // Styles tab (S2): a live styles panel — hover to preview, click to apply,
    // right-click for select-all-instances / update-to-match / rename-everywhere.
    // A chip surfaces the caret's direct-formatting overrides. Preview uses the
    // "transient" apply + restoreParagraphsCmd primitives (no undo pollution).
    let stylesChild: ReturnType<typeof editor.createChild> | null = null;
    let stylesPreviewSnap: ParaSnapshot[] | null = null;
    const endStylesPreview = (): void => {
      if (stylesPreviewSnap) { editor.dispatch(restoreParagraphsCmd(stylesPreviewSnap)); stylesPreviewSnap = null; }
    };
    const buildStyles = (): void => {
      endStylesPreview();
      stylesPanel.textContent = "";
      stylesChild?.destroy();
      stylesChild = editor.createChild();
      const child = stylesChild;
      const doc = editor.getDocument();
      const sheet = doc.stylesheet ?? defaultStylesheet();
      const caret = editor.getSelection()?.focus ?? null;

      if (caret && hasDirectFormattingAt(doc, caret)) {
        const chip = el("div", "cw-styles-chip");
        const txt = el("span");
        txt.textContent = "Direct formatting on this text";
        const clearBtn = el("button", "cw-styles-chip-clear");
        clearBtn.textContent = "Clear";
        clearBtn.title = "Reset this text to its paragraph style";
        clearBtn.addEventListener("mousedown", (e) => e.preventDefault());
        clearBtn.addEventListener("click", () => {
          const s = editor.getSelection();
          if (!s) return;
          const collapsed = s.anchor.blockId === s.focus.blockId && s.anchor.offset === s.focus.offset;
          const orig = s.focus;
          // clearCharFormatting no-ops on a collapsed caret — expand to the whole
          // paragraph so "Clear" always does something, then re-bake its named style
          // (clear resets to raw defaults; the style re-application restores it).
          if (collapsed) {
            const b = editor.getDocument().blocks.find((x) => x.id === s.focus.blockId);
            const len = b && b.kind === "paragraph" ? b.runs.reduce((nn, r) => nn + r.text.length, 0) : 0;
            editor.setSelection({ anchor: { blockId: s.focus.blockId, offset: 0 }, focus: { blockId: s.focus.blockId, offset: len } });
          }
          editor.dispatch(clearCharFormatting());
          const sid = editor.currentFormat().styleId;
          if (sid) editor.dispatch(applyNamedStyle(sid));
          editor.setSelection({ anchor: orig, focus: orig });
          editor.focus();
          buildStyles();
        });
        chip.append(txt, clearBtn);
        stylesPanel.appendChild(chip);
      }

      const listEl = el("div", "cw-styles-list");
      stylesPanel.appendChild(listEl);
      const currentId = editor.currentFormat().styleId;

      // Paragraphs the apply would touch — snapshotted before a transient preview.
      const affectedParas = (): ParaSnapshot[] => {
        const s = editor.getSelection();
        if (!s) return [];
        const blocks = editor.getDocument().blocks;
        const ai = blocks.findIndex((b) => b.id === s.anchor.blockId);
        const fi = blocks.findIndex((b) => b.id === s.focus.blockId);
        if (ai < 0 || fi < 0) return []; // not top-level (e.g. a table cell) — skip preview
        const [lo, hi] = ai <= fi ? [ai, fi] : [fi, ai];
        const out: ParaSnapshot[] = [];
        for (const b of blocks.slice(lo, hi + 1)) if (b.kind === "paragraph") out.push({ id: b.id, style: b.style, runs: b.runs });
        return out;
      };

      const selectAllInstances = (styleId: string): void => {
        const ranges = styleInstanceRanges(editor.getDocument(), styleId);
        if (!ranges.length) return;
        const first = ranges[0]!;
        editor.setSelection({ anchor: { blockId: first.blockId, offset: first.start }, focus: { blockId: first.blockId, offset: first.end } });
        editor.revealBlock(first.blockId);
        editor.setDecorations(ranges.map((r) => ({
          type: "highlight" as const, color: "#1a73e8", opacity: 0.26,
          range: { anchor: { blockId: r.blockId, offset: r.start }, focus: { blockId: r.blockId, offset: r.end } },
        })));
        window.setTimeout(() => editor.setDecorations([]), 4000); // flash all instances, don't linger
        editor.focus();
      };

      const renameStyle = (styleId: string, currentName: string): void => {
        showInputDialog({
          title: "Rename style", label: "Style name", initial: currentName, okLabel: "Rename",
          validate: (v) => (v.trim() === "" ? "Enter a name." : null),
          onSubmit: (next) => { if (next.trim() !== currentName) editor.dispatch(renameNamedStyle(styleId, next.trim())); editor.focus(); buildStyles(); },
        });
      };

      for (const st of sheet.styles) {
        const isChar = styleType(st) === "character";
        const count = styleInstanceRanges(doc, st.id).length;
        const card = el("button", "cw-style-card");
        if (st.id === currentId) card.classList.add("current");
        const swatch = el("span", "cw-style-swatch");
        const name = el("span", "cw-style-name");
        name.textContent = st.name + (isChar ? " ⓐ" : "");
        const cnt = el("span", "cw-style-count");
        cnt.textContent = count > 0 ? String(count) : "";
        card.append(swatch, name, cnt);
        card.title = `Apply "${st.name}"${count ? ` · ${count} in use` : ""}`;
        child.render(swatch, { kind: "styleSample", styleId: st.id, sampleText: "AaBbCc" }, { fit: true, widthPx: 88, maxHeightPx: 24 });
        card.addEventListener("mousedown", (e) => e.preventDefault());
        card.addEventListener("mouseenter", () => {
          endStylesPreview();
          const snap = affectedParas();
          if (!snap.length) return;
          stylesPreviewSnap = snap;
          editor.dispatch(isChar ? applyCharStyle(st.id, "transient") : applyNamedStyle(st.id, "transient"));
        });
        card.addEventListener("mouseleave", endStylesPreview);
        card.addEventListener("click", () => {
          endStylesPreview();
          editor.dispatch(isChar ? applyCharStyle(st.id) : applyNamedStyle(st.id));
          editor.focus();
          buildStyles();
        });
        card.addEventListener("contextmenu", (e) => {
          e.preventDefault();
          endStylesPreview();
          const entries: MenuEntry[] = [
            { kind: "item", label: "Select all instances", disabled: count === 0, onClick: () => selectAllInstances(st.id) },
            { kind: "item", label: "Update to match selection", onClick: () => { editor.dispatch(updateStyleToSelection(st.id)); editor.focus(); buildStyles(); } },
            { kind: "item", label: "Rename everywhere…", onClick: () => renameStyle(st.id, st.name) },
          ];
          showContextMenu(e.clientX, e.clientY, entries);
        });
        listEl.appendChild(card);
      }
      if (sheet.styles.length === 0) {
        const empty = el("div", "outline-empty");
        empty.textContent = "No styles defined.";
        stylesPanel.appendChild(empty);
      }
    };
    setNav("headings"); // default tab
    // Drag-to-resize handle on the pane's right edge (O5).
    const resizer = el("div", "cw-outline-resize");
    resizer.title = "Drag to resize";
    outlineEl.appendChild(resizer);
    resizer.addEventListener("mousedown", (e) => {
      e.preventDefault();
      const startX = e.clientX;
      const startW = outlineEl.getBoundingClientRect().width;
      const onMove = (ev: MouseEvent): void => {
        const w = Math.max(180, Math.min(520, startW + (ev.clientX - startX)));
        outlineEl.style.flex = `0 0 ${w}px`;
        outlineEl.style.width = `${w}px`;
      };
      const onUp = (): void => {
        document.removeEventListener("mousemove", onMove);
        document.removeEventListener("mouseup", onUp);
      };
      document.addEventListener("mousemove", onMove);
      document.addEventListener("mouseup", onUp);
    });

    const collapsed = new Set<string>(); // heading ids whose sub-tree is collapsed
    let filterText = "";

    // Detect a heading + its outline level. Real .docx files name their styles
    // "Heading 1" but keep an OPAQUE styleId (e.g. "Style27"), so we resolve the
    // level through the stylesheet's name/basedOn chain — mirroring the
    // importer's isHeading(). Built-in ids like "Heading1"/"Title" still match
    // directly (the sample document has no separate stylesheet entry for them).
    const labelLevel = (label: string | undefined): number | null => {
      if (!label) return null;
      const m = /(?:^|\s)heading\s*([1-9])/i.exec(label);
      if (m) return Number(m[1]);
      if (/^\s*title\s*$/i.test(label)) return 0;
      if (/(?:^|\s)heading(?:\s|$)/i.test(label)) return 1; // "Heading" with no number
      return null;
    };
    const headingLevel = (styleId: string | undefined): number | null => {
      if (!styleId) return null;
      const sheet = editor.getDocument().stylesheet ?? defaultStylesheet();
      let level = labelLevel(styleId);
      const seen = new Set<string>();
      for (
        let cur = styleById(sheet, styleId);
        cur && level === null && !seen.has(cur.id);
        cur = cur.basedOn ? styleById(sheet, cur.basedOn) : undefined
      ) {
        seen.add(cur.id);
        level = labelLevel(cur.name) ?? labelLevel(cur.id);
      }
      return level;
    };
    type Entry = { id: string; index: number; level: number; text: string };
    let entries: Entry[] = [];
    const rows = new Map<string, HTMLElement>();
    const pageSpans = new Map<string, HTMLElement>();
    let lastSig = "\0"; // sentinel — guarantees the first render runs

    const updateActive = (): void => {
      const caret = editor.getSelection()?.focus.blockId ?? null;
      const ci = caret ? editor.getDocument().blocks.findIndex((b) => b.id === caret) : -1;
      let activeId: string | null = null;
      for (const e of entries) if (ci >= 0 && e.index <= ci) activeId = e.id; // nearest heading at/above caret
      for (const [id, r] of rows) r.classList.toggle("active", id === activeId);
    };
    const updatePages = (): void => {
      for (const [id, span] of pageSpans) {
        const p = editor.getBlockPage(id);
        span.textContent = p != null ? String(p) : "";
      }
    };
    const hasChildren = (i: number): boolean => i + 1 < entries.length && entries[i + 1]!.level > entries[i]!.level;

    // Drag-to-reorder (O1): dragging a heading moves its WHOLE section — the
    // heading plus every block up to the next same-or-shallower heading —
    // reusing the block-range move that Organize Pages uses (never splits).
    let dragFromId: string | null = null;
    const dropLine = el("div", "cw-outline-drop");
    dropLine.style.display = "none";
    outlineEl.appendChild(dropLine);
    const sectionRange = (headingId: string): { firstId: string; lastId: string } | null => {
      const blocks = editor.getDocument().blocks;
      const i = entries.findIndex((e) => e.id === headingId);
      const startIdx = blocks.findIndex((b) => b.id === headingId);
      if (i < 0 || startIdx < 0) return null;
      let endIdx = blocks.length;
      for (let j = i + 1; j < entries.length; j++) {
        if (entries[j]!.level <= entries[i]!.level) { endIdx = blocks.findIndex((b) => b.id === entries[j]!.id); break; }
      }
      return { firstId: blocks[startIdx]!.id, lastId: blocks[endIdx - 1]!.id };
    };
    const afterAnchor = (headingId: string): string | null => {
      const i = entries.findIndex((e) => e.id === headingId);
      for (let j = i + 1; j < entries.length; j++) if (entries[j]!.level <= entries[i]!.level) return entries[j]!.id;
      return null;
    };
    const performDrop = (targetId: string, after: boolean): void => {
      if (!dragFromId || dragFromId === targetId) return;
      const range = sectionRange(dragFromId);
      if (!range) return;
      const anchorId = after ? afterAnchor(targetId) : targetId;
      editor.dispatch(reorderPageGroupCmd(range.firstId, range.lastId, anchorId));
      editor.focus();
    };

    // Render the visible rows (O2 collapse + O3 filter + O4 level styling + O6
    // page numbers). Structure-only; page numbers/active state patch separately.
    const render = (): void => {
      list.textContent = "";
      rows.clear();
      pageSpans.clear();
      if (entries.length === 0) {
        const empty = el("div", "outline-empty");
        empty.textContent = "No headings yet. Apply a Heading style (Heading 1–9) to build an outline.";
        list.appendChild(empty);
        return;
      }
      const q = filterText.trim().toLowerCase();
      let collapseLevel: number | null = null; // while set, hide entries deeper than it
      entries.forEach((e, i) => {
        let visible: boolean;
        if (q) visible = e.text.toLowerCase().includes(q); // filter is flat, ignores collapse
        else if (collapseLevel !== null && e.level > collapseLevel) visible = false;
        else {
          collapseLevel = null;
          visible = true;
          if (collapsed.has(e.id) && hasChildren(i)) collapseLevel = e.level;
        }
        if (!visible) return;
        const row = el("div", `outline-item lvl-${Math.min(e.level, 3)}`);
        row.style.paddingLeft = `${6 + e.level * 12}px`;
        const chev = el("button", "outline-chevron");
        chev.setAttribute("aria-hidden", "true");
        if (hasChildren(i) && !q) {
          const isCol = collapsed.has(e.id);
          chev.textContent = isCol ? "▸" : "▾";
          chev.title = isCol ? "Expand section" : "Collapse section";
          chev.addEventListener("mousedown", (ev) => ev.preventDefault());
          chev.addEventListener("click", (ev) => {
            ev.stopPropagation();
            if (collapsed.has(e.id)) collapsed.delete(e.id);
            else collapsed.add(e.id);
            build();
          });
        } else {
          chev.classList.add("leaf");
        }
        const label = el("span", "outline-label");
        label.textContent = e.text;
        label.title = e.text; // full text on hover — truncation is never unrecoverable (O5)
        const page = el("span", "outline-page");
        row.append(chev, label, page);
        row.addEventListener("mousedown", (ev) => ev.preventDefault());
        row.addEventListener("click", () => editor.revealBlock(e.id));
        // Drag-to-reorder this heading's section.
        row.draggable = true;
        row.addEventListener("dragstart", (ev) => {
          dragFromId = e.id;
          if (ev.dataTransfer) { ev.dataTransfer.setData("text/plain", e.id); ev.dataTransfer.effectAllowed = "move"; }
          row.classList.add("cw-outline-dragging");
        });
        row.addEventListener("dragend", () => { row.classList.remove("cw-outline-dragging"); dropLine.style.display = "none"; dragFromId = null; });
        row.addEventListener("dragover", (ev) => {
          if (!dragFromId || dragFromId === e.id) return;
          ev.preventDefault();
          if (ev.dataTransfer) ev.dataTransfer.dropEffect = "move";
          const r = row.getBoundingClientRect();
          const after = ev.clientY > r.top + r.height / 2;
          row.dataset["dropAfter"] = after ? "1" : "0";
          dropLine.style.display = "block";
          dropLine.style.top = `${(after ? r.bottom : r.top) - outlineEl.getBoundingClientRect().top}px`;
        });
        row.addEventListener("drop", (ev) => {
          ev.preventDefault();
          performDrop(e.id, row.dataset["dropAfter"] === "1");
          dropLine.style.display = "none";
        });
        rows.set(e.id, row);
        pageSpans.set(e.id, page);
        list.appendChild(row);
      });
    };

    const build = (): void => {
      const collected: Entry[] = [];
      editor.getDocument().blocks.forEach((b, index) => {
        if (b.kind !== "paragraph") return;
        const level = headingLevel(b.style.namedStyle);
        if (level === null) return;
        const text = b.runs.map((r) => r.text).join("").trim();
        if (text === "") return; // skip empty heading-styled paragraphs (structural artifacts)
        collected.push({ id: b.id, index, level, text });
      });
      entries = collected;
      // Re-render only when the heading set / collapse / filter changed, so
      // typing (which fires refreshOutline) doesn't reset the pane's scroll.
      const sig =
        collected.map((c) => `${c.id}|${c.level}|${c.text}`).join("\n") +
        "" + [...collapsed].sort().join(",") +
        "" + filterText.trim().toLowerCase();
      if (sig !== lastSig) {
        lastSig = sig;
        render();
      }
      updatePages();
      updateActive();
    };

    filterInput.addEventListener("input", () => {
      filterText = filterInput.value;
      build();
    });

    let open = false;
    const setOpen = (v: boolean): void => {
      open = v;
      outlineEl.classList.toggle("open", v);
      outlineBtn.classList.toggle("active", v);
      if (v) build();
    };
    closeBtn.addEventListener("click", () => { setOpen(false); autoCollapsedOutline = false; }); // manual close cancels auto-reopen
    outlineToggle = (): void => { setOpen(!open); autoCollapsedOutline = false; };
    refreshOutline = (): void => {
      if (open) build();
    };
    setOutlineOpen = setOpen; // for the responsive auto-collapse (row 12)
    isOutlineOpen = (): boolean => open;
    setOpen(runtime.view?.outline ?? true); // visible by default; embedder can start it closed
  }

  // ---- Bookmarks panel (list + Go To + add/rename/delete) -----------------
  {
    // Bookmarks render into the Navigator's Marks tab (row 17), not a drawer.
    const addBtn = el("button", "cw-nav-action");
    addBtn.textContent = "＋ Add bookmark";
    addBtn.title = "Add a bookmark for the current selection";
    const list = el("div", "cw-outline-list");
    if (navMarksPanel) navMarksPanel.append(addBtn, list);

    // OOXML bookmark-name rules — a native prompt() could neither enforce nor
    // explain these (critique B5). Returns an error message, or null when valid.
    const bookmarkNameError = (raw: string, self?: string): string | null => {
      const n = raw.trim();
      if (n === "") return "Enter a name.";
      if (n.length > 40) return "Use 40 characters or fewer.";
      if (/\s/.test(n)) return "No spaces — use letters, digits, or underscores.";
      if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(n)) return "Start with a letter or _, then letters, digits, or _.";
      if (n !== self && Object.prototype.hasOwnProperty.call(editor.getDocument().bookmarks ?? {}, n)) return `"${n}" already exists.`;
      return null;
    };

    const build = (): void => {
      const names = Object.keys(editor.getDocument().bookmarks ?? {}).sort((a, b) => a.localeCompare(b));
      list.textContent = "";
      if (names.length === 0) {
        const empty = el("div", "outline-empty");
        empty.textContent = "No bookmarks. Select text (or place the caret) and press + to add one.";
        list.appendChild(empty);
        return;
      }
      for (const name of names) {
        const row = el("div");
        row.style.cssText = "display:flex;align-items:center;gap:2px;padding:0 6px 0 0;";
        const go = el("button", "outline-item");
        go.style.cssText = "flex:1 1 auto;";
        go.textContent = name;
        go.title = `Go to "${name}"`;
        go.addEventListener("mousedown", (e) => e.preventDefault());
        go.addEventListener("click", () => editor.revealBookmark(name));
        const iconBtn = (glyph: string, tip: string, onClick: () => void): HTMLButtonElement => {
          const b = el("button");
          b.textContent = glyph;
          b.title = tip;
          b.style.cssText = "border:none;background:transparent;cursor:pointer;color:#605e5c;width:24px;height:24px;border-radius:4px;font-size:13px;";
          b.addEventListener("mousedown", (e) => e.preventDefault());
          b.addEventListener("click", onClick);
          return b;
        };
        const renameB = iconBtn("✎", "Rename bookmark", () => {
          showInputDialog({
            title: "Rename bookmark",
            label: "Bookmark name",
            initial: name,
            okLabel: "Rename",
            validate: (v) => bookmarkNameError(v, name),
            onSubmit: (next) => {
              if (next.trim() !== name) editor.dispatch(renameBookmarkCmd(name, next.trim()));
              editor.focus();
            },
          });
        });
        const delB = iconBtn("🗑", "Delete bookmark", () => {
          editor.dispatch(removeBookmarkCmd(name));
          editor.focus();
        });
        row.append(go, renameB, delB);
        list.appendChild(row);
      }
    };

    addBtn.addEventListener("click", () => {
      showInputDialog({
        title: "Add bookmark",
        label: "Bookmark name",
        placeholder: "e.g. Introduction",
        okLabel: "Add",
        validate: (v) => bookmarkNameError(v),
        onSubmit: (name) => {
          editor.dispatch(addBookmarkCmd(name.trim()));
          editor.focus();
          build();
        },
      });
    });
    toggleBookmarks = (): void => { setOutlineOpen(true); navShow("marks"); };
    refreshBookmarks = (): void => { if (navMarksPanel) build(); };
    build();
    if (runtime.view?.bookmarks) { setOutlineOpen(true); navShow("marks"); }
  }

  // ---- Review pane (track changes + comments) — docked right, in-shell ----
  const COMMENT_STYLE = { ...DEFAULT_CHAR_STYLE };
  const commentFragment = (text: string): Fragment => [{ text, style: COMMENT_STYLE }];
  {
    const reviewEl = shell.review;
    const authorName = (a: { firstName: string; lastName: string }): string => `${a.firstName} ${a.lastName}`.trim() || "Anonymous";
    const initials = (a: { firstName: string; lastName: string }): string => ((a.firstName[0] ?? "?") + (a.lastName[0] ?? "")).toUpperCase();
    const timeAgo = (ms: number): string => {
      const s = Math.max(0, Math.round((Date.now() - ms) / 1000));
      if (s < 60) return "just now";
      const m = Math.round(s / 60);
      if (m < 60) return `${m}m ago`;
      const h = Math.round(m / 60);
      if (h < 24) return `${h}h ago`;
      return `${Math.round(h / 24)}d ago`;
    };
    const mkBtn = (cls: string, label: string, onClick: () => void): HTMLButtonElement => {
      const b = el("button") as HTMLButtonElement;
      b.className = cls;
      b.textContent = label;
      b.addEventListener("mousedown", (e) => e.preventDefault());
      b.addEventListener("click", onClick);
      return b;
    };
    const avatar = (a: { id: string; firstName: string; lastName: string }): HTMLElement => {
      const av = el("div", "cw-avatar");
      av.style.background = colorForId(a.id);
      av.textContent = initials(a);
      return av;
    };

    // header
    const head = el("div", "cw-review-head");
    const title = el("span", "cw-review-title");
    title.textContent = "Review";
    const closeBtn = mkBtn("cw-review-close", "×", () => setOpen(false));
    closeBtn.title = "Close";
    head.append(title, closeBtn);

    // tabs
    const tabs = el("div", "cw-review-tabs");
    const tabSug = el("button", "cw-review-tab") as HTMLButtonElement;
    const tabCom = el("button", "cw-review-tab") as HTMLButtonElement;
    const sugPill = el("span", "cw-pill");
    const comPill = el("span", "cw-pill");
    tabSug.append(document.createTextNode("Suggestions "), sugPill);
    tabCom.append(document.createTextNode("Comments "), comPill);
    tabs.append(tabSug, tabCom);

    const body = el("div", "cw-review-body");
    reviewEl.append(head, tabs, body);

    let activeTab: "suggestions" | "comments" = "suggestions";

    const emptyState = (ico: string, text: string): HTMLElement => {
      const e = el("div", "cw-review-empty");
      const i = el("div", "cw-review-empty-ico");
      i.textContent = ico;
      const t = el("div");
      t.textContent = text;
      e.append(i, t);
      return e;
    };

    const renderSuggestions = (review: ReturnType<typeof editor.getReview>): void => {
      if (review.suggestions.length === 0) {
        body.append(emptyState("✦", "No suggestions yet. Switch the mode to Suggesting and edit — your changes become tracked proposals here."));
        return;
      }
      const bar = el("div", "cw-review-actions");
      bar.append(
        mkBtn("cw-btn cw-btn-accept cw-btn-sm", "✓ Accept all", () => { editor.acceptAllSuggestions(); editor.focus(); }),
        mkBtn("cw-btn cw-btn-reject cw-btn-sm", "✗ Reject all", () => { editor.rejectAllSuggestions(); editor.focus(); }),
      );
      body.append(bar);
      for (const s of review.suggestions) {
        const card = el("div", "cw-sug");
        card.style.setProperty("--cw-author", colorForId(s.author.id));
        const top = el("div", "cw-sug-top");
        const kind = el("div", "cw-sug-kind");
        const dot = el("span", "cw-dot");
        const verb = s.kind === "insert" ? "Insertion" : s.kind === "delete" ? "Deletion" : "Formatting";
        kind.append(dot, document.createTextNode(verb));
        const meta = el("div", "cw-sug-meta");
        meta.textContent = `${authorName(s.author)} · ${timeAgo(s.createdAt)}`;
        top.append(kind, meta);
        const actions = el("div", "cw-sug-actions");
        actions.append(
          mkBtn("cw-btn cw-btn-accept cw-btn-sm", "✓ Accept", () => { editor.acceptSuggestion(s.id); editor.focus(); }),
          mkBtn("cw-btn cw-btn-reject cw-btn-sm", "✗ Reject", () => { editor.rejectSuggestion(s.id); editor.focus(); }),
        );
        card.append(top, actions);
        card.addEventListener("click", (e) => {
          if (!(e.target as HTMLElement).closest("button")) editor.revealReview(s.id);
        });
        body.append(card);
      }
    };

    // Render a comment body as text + highlighted @mention chips (matched against
    // the comment's resolved mentions; longest display first to avoid overlaps).
    const renderCommentBody = (host: HTMLElement, text: string, mentions?: { firstName: string; lastName: string }[]): void => {
      const tokens = (mentions ?? []).map((u) => `@${authorName(u)}`).sort((a, b) => b.length - a.length);
      if (tokens.length === 0) {
        host.textContent = text;
        return;
      }
      let rest = text;
      while (rest.length > 0) {
        let at = -1;
        let tok = "";
        for (const n of tokens) {
          const i = rest.indexOf(n);
          if (i >= 0 && (at < 0 || i < at)) { at = i; tok = n; }
        }
        if (at < 0) { host.appendChild(document.createTextNode(rest)); break; }
        if (at > 0) host.appendChild(document.createTextNode(rest.slice(0, at)));
        const chip = el("span", "cw-mention");
        chip.textContent = tok;
        host.appendChild(chip);
        rest = rest.slice(at + tok.length);
      }
    };

    const renderComments = (review: ReturnType<typeof editor.getReview>): void => {
      if (review.threads.length === 0) {
        body.append(emptyState("💬", "No comments yet. Select text in the document and click 💬 in the toolbar to start a discussion."));
        return;
      }
      for (const t of review.threads) {
        const card = el("div", t.status === "resolved" ? "cw-thread resolved" : "cw-thread");
        for (const c of t.comments) {
          const cm = el("div", "cw-comment");
          const main = el("div", "cw-comment-main");
          const who = el("div", "cw-comment-who");
          who.textContent = authorName(c.author);
          const when = el("span", "cw-comment-when");
          when.textContent = timeAgo(c.createdAt);
          who.append(when);
          const text = el("div", "cw-comment-body");
          renderCommentBody(text, c.body.map((r) => r.text).join(""), c.mentions);
          main.append(who, text);
          cm.append(avatar(c.author), main);
          card.append(cm);
        }
        // inline reply editor (Google-Docs-style, no prompt)
        const replyBox = el("div", "cw-reply-box");
        const replyTa = el("textarea") as HTMLTextAreaElement;
        replyTa.placeholder = "Reply…";
        replyTa.style.cssText = "flex:1 1 auto;resize:none;min-height:34px;border:1px solid #dadce0;border-radius:6px;padding:6px 8px;font:13px/1.4 inherit;outline:none;";
        const replyMentions = attachMentionAutocomplete(replyTa, () => editor.getKnownUsers());
        const replySend = mkBtn("cw-btn cw-btn-primary cw-btn-sm", "Reply", () => {
          const v = replyTa.value.trim();
          if (!v) return;
          editor.replyToComment(t.id, commentFragment(v), replyMentions.getMentions());
          editor.focus();
        });
        replyBox.append(replyTa, replySend);
        const actions = el("div", "cw-thread-actions");
        if (t.status === "resolved") {
          const tag = el("span", "cw-resolved-tag");
          tag.textContent = "✓ Resolved";
          actions.append(tag);
          actions.append(mkBtn("cw-btn cw-btn-ghost cw-btn-sm", "Reopen", () => { editor.resolveThread(t.id, false); editor.focus(); }));
        } else {
          actions.append(
            mkBtn("cw-btn cw-btn-ghost cw-btn-sm", "Reply", () => {
              replyBox.classList.add("open");
              replyTa.focus();
            }),
            mkBtn("cw-btn cw-btn-ghost cw-btn-sm", "Resolve", () => { editor.resolveThread(t.id, true); editor.focus(); }),
          );
        }
        card.append(replyBox, actions);
        card.addEventListener("click", (e) => {
          if (!(e.target as HTMLElement).closest("button, textarea, .cw-reply-box")) editor.revealReview(t.id);
        });
        body.append(card);
      }
    };

    const build = (): void => {
      const review = editor.getReview();
      sugPill.textContent = String(review.suggestions.length);
      comPill.textContent = String(review.threads.filter((t) => t.status === "open").length);
      tabSug.classList.toggle("active", activeTab === "suggestions");
      tabCom.classList.toggle("active", activeTab === "comments");
      body.textContent = "";
      if (activeTab === "suggestions") renderSuggestions(review);
      else renderComments(review);
    };
    tabSug.addEventListener("click", () => { activeTab = "suggestions"; build(); });
    tabCom.addEventListener("click", () => { activeTab = "comments"; build(); });

    let open = false;
    const setOpen = (v: boolean): void => {
      open = v;
      reviewEl.classList.toggle("open", v);
      reviewBtn.classList.toggle("active", v);
      if (v) build();
    };
    toggleReview = (): void => setOpen(!open);
    refreshReview = (): void => {
      if (open) build();
    };
    if (runtime.view?.reviewPane) setOpen(true); // closed by default; embedder can open it
  }

  // ---- Inspector (Move 2): right-docked, selection-aware property sheet ------
  // Text + Paragraph + Table + Page + Object (image/shape) sections; every control
  // applies LIVE as one undoable edit (no Apply button). Each section covers the
  // common levers; the matching dialog stays as the advanced fallback (colour,
  // effects, borders, page background, anchor offsets) per the parity rule, so no
  // dialog is deleted.
  //
  // Information architecture: a breadcrumb scope trail (Page › Table › Cell ›
  // Paragraph › Text) heads the panel; below it the sections render outer→inner and
  // each is a collapsible disclosure whose collapsed header keeps a live one-line
  // value summary (collapsing costs the controls, never the information). Default
  // expansion follows the selection's tightest scope; a manual toggle wins and
  // persists. See the machinery block below.
  {
    const inspectorEl = shell.inspector;
    const head = el("div", "cw-insp-head");
    const title = el("span", "cw-insp-title");
    title.textContent = "Inspector";
    const closeBtn = el("button", "cw-insp-close");
    closeBtn.innerHTML = "×";
    closeBtn.title = "Close inspector";
    closeBtn.setAttribute("aria-label", "Close inspector");
    closeBtn.addEventListener("click", () => toggleInspector());
    head.append(title, closeBtn);
    const bodyEl = el("div", "cw-insp-body");
    inspectorEl.append(head, bodyEl);

    const toPt = (px: number): number => Math.round(sharedPxToPt(px) * 2) / 2;
    const fromPt = (pt: number): number => sharedPtToPx(pt);

    // ---- Information architecture (Move 2 restructure) ----------------------
    // Every section is a collapsible disclosure. A COLLAPSED header still shows a
    // live one-line value summary of its own state (Text → "Calibri · 12 · Bold"),
    // so collapsing costs the user the CONTROLS but never the INFORMATION. Default
    // expansion follows the selection's tightest scope (so "Page collapsed by
    // default" falls out as a consequence, not a special case); a manual toggle
    // wins over that rule and PERSISTS (localStorage — the same convention as the
    // colour/shape recents) so the panel never fights the user.
    const OVERRIDE_KEY = "cw:inspector:collapse";
    const loadOverrides = (): Record<string, boolean> => {
      try {
        const raw = localStorage.getItem(OVERRIDE_KEY);
        const o: unknown = raw ? JSON.parse(raw) : {};
        return o && typeof o === "object" ? (o as Record<string, boolean>) : {};
      } catch {
        return {};
      }
    };
    const sectionOverrides = loadOverrides();
    const persistOverrides = (): void => {
      try { localStorage.setItem(OVERRIDE_KEY, JSON.stringify(sectionOverrides)); } catch { /* storage optional */ }
    };
    // The section whose scope the selection most tightly addresses — set per
    // rebuild. Its section is expanded by default; ancestors collapse.
    let tightestSection = "Page";
    const isExpanded = (name: string): boolean =>
      name in sectionOverrides ? !!sectionOverrides[name] : name === tightestSection;
    const toggleSection = (name: string): void => {
      sectionOverrides[name] = !isExpanded(name);
      persistOverrides();
      rebuild();
    };
    // A collapsible section. `summary` shows in the header while collapsed; the
    // returned element is the body the caller appends rows into. Callers early-out
    // (skip building rows) when `expanded` is false — the summary already carries
    // the information.
    const insSection = (name: string, summary: string, expanded: boolean): HTMLElement => {
      const s = el("div", "cw-insp-section" + (expanded ? "" : " collapsed"));
      const h = el("button", "cw-insp-sec-head");
      (h as HTMLButtonElement).type = "button";
      h.setAttribute("aria-expanded", String(expanded));
      const chev = el("span", "cw-insp-chevron");
      chev.innerHTML = chevron(false);
      const t = el("span", "cw-insp-sec-title");
      t.textContent = name;
      const sum = el("span", "cw-insp-sec-summary");
      sum.textContent = summary;
      h.append(chev, t, sum);
      h.addEventListener("click", () => toggleSection(name));
      const body = el("div", "cw-insp-sec-body");
      s.append(h, body);
      bodyEl.append(s);
      return body;
    };
    const cap = (s: string): string => s.charAt(0).toUpperCase() + s.slice(1);
    const insRow = (parent: HTMLElement, label: string | null, ...controls: HTMLElement[]): void => {
      const r = el("div", "cw-insp-row");
      if (label !== null) { const l = el("span", "cw-insp-label"); l.textContent = label; r.append(l); }
      for (const c of controls) r.append(c);
      parent.append(r);
    };
    const insSelect = (opts: { value: string; label: string }[], value: string, onChange: (v: string) => void): HTMLSelectElement => {
      const s = document.createElement("select");
      for (const o of opts) { const op = document.createElement("option"); op.value = o.value; op.textContent = o.label; s.append(op); }
      s.value = value;
      s.addEventListener("change", () => { onChange(s.value); editor.focus(); });
      return s;
    };
    const insNumber = (value: number, min: number, max: number, onCommit: (n: number) => void): HTMLInputElement => {
      const i = document.createElement("input");
      i.type = "number"; i.min = String(min); i.max = String(max); i.step = "1";
      i.value = String(value);
      i.addEventListener("change", () => { const n = Number(i.value); if (Number.isFinite(n)) onCommit(n); editor.focus(); });
      return i;
    };
    // A decimal-stepped numeric input (page margins in inches etc.).
    const insDecimal = (value: number, min: number, max: number, step: number, onCommit: (n: number) => void): HTMLInputElement => {
      const i = document.createElement("input");
      i.type = "number"; i.min = String(min); i.max = String(max); i.step = String(step);
      i.value = String(value);
      i.addEventListener("change", () => { const n = Number(i.value); if (Number.isFinite(n)) onCommit(n); editor.focus(); });
      return i;
    };
    // Two controls side-by-side in one row cell (e.g. Top/Bottom margins).
    const insPair = (a: HTMLElement, b: HTMLElement): HTMLElement => {
      const w = el("div", "cw-insp-pair");
      w.append(a, b);
      return w;
    };
    const insToggles = (defs: { label: string; title: string; on: boolean; onClick: () => void }[]): HTMLElement => {
      const wrap = el("div", "cw-insp-btns");
      for (const d of defs) {
        const b = el("button", "cw-insp-tgl");
        b.textContent = d.label;
        b.title = d.title;
        if (d.on) b.classList.add("on");
        b.addEventListener("mousedown", (e) => e.preventDefault());
        b.addEventListener("click", () => { d.onClick(); editor.focus(); });
        wrap.append(b);
      }
      return wrap;
    };

    const textSummary = (f: ReturnType<typeof editor.currentFormat>): string => {
      const fam = (f.fontFamily ?? "").split(",")[0]!.trim() || "Default";
      const flags: string[] = [];
      if (f.bold) flags.push("Bold");
      if (f.italic) flags.push("Italic");
      if (f.underline) flags.push("Underline");
      if (f.strikethrough) flags.push("Strike");
      return `${fam} · ${toPt(f.fontSizePx ?? 16)} · ${flags.length ? flags.join(" ") : "Regular"}`;
    };
    const buildTextSection = (f: ReturnType<typeof editor.currentFormat>): void => {
      const expanded = isExpanded("Text");
      const sec = insSection("Text", textSummary(f), expanded);
      if (!expanded) return;
      const fams = toolbarFonts(config.fonts).map((x) => ({ value: x.value, label: x.label }));
      const famVal = fams.find((x) => x.value.toLowerCase() === (f.fontFamily ?? "").toLowerCase())?.value ?? f.fontFamily ?? fams[0]?.value ?? "";
      insRow(sec, "Font", insSelect(fams, famVal, (v) => editor.setCharStyle({ fontFamily: v })));
      insRow(sec, "Size (pt)", insNumber(toPt(f.fontSizePx ?? 16), 6, 72, (pt) => editor.setCharStyle({ fontSizePx: Math.min(96, Math.max(6, fromPt(pt))) })));
      insRow(sec, "Style", insToggles([
        { label: "B", title: "Bold", on: f.bold, onClick: () => editor.toggleStyle("bold") },
        { label: "I", title: "Italic", on: f.italic, onClick: () => editor.toggleStyle("italic") },
        { label: "U", title: "Underline", on: f.underline, onClick: () => editor.toggleStyle("underline") },
        { label: "S", title: "Strikethrough", on: f.strikethrough, onClick: () => editor.toggleStyle("strikethrough") },
      ]));
      insRow(sec, "Caps", insToggles([
        { label: "AB", title: "All caps", on: f.caps, onClick: () => editor.setCharStyle({ caps: !f.caps }) },
        { label: "Ab", title: "Small caps", on: f.smallCaps, onClick: () => editor.setCharStyle({ smallCaps: !f.smallCaps }) },
      ]));
    };

    const lineLabel = (lh: number): string =>
      lh === 1 ? "Single" : lh === 1.15 ? "1.15" : lh === 1.5 ? "1.5" : lh === 2 ? "Double" : String(lh);
    const paragraphSummary = (f: ReturnType<typeof editor.currentFormat>, ps: ParaStyle | null): string => {
      const lh = ps?.lineHeight ?? 1;
      return `${cap(f.align ?? "left")} · ${lineLabel(lh)} · ${toPt(ps?.spaceBeforePx ?? 0)}/${toPt(ps?.spaceAfterPx ?? 0)} pt`;
    };
    const buildParagraphSection = (f: ReturnType<typeof editor.currentFormat>): void => {
      const ps = editor.currentParaStyle();
      const expanded = isExpanded("Paragraph");
      const sec = insSection("Paragraph", paragraphSummary(f, ps), expanded);
      if (!expanded) return;
      const aligns = (["left", "center", "right", "justify"] as const);
      insRow(sec, "Align", insToggles(aligns.map((a) => ({
        label: a.charAt(0).toUpperCase(), title: a.charAt(0).toUpperCase() + a.slice(1), on: f.align === a, onClick: () => editor.align(a),
      }))));
      const lh = ps?.lineHeight ?? 1;
      insRow(sec, "Line", insSelect(
        [{ value: "1", label: "Single" }, { value: "1.15", label: "1.15" }, { value: "1.5", label: "1.5" }, { value: "2", label: "Double" }],
        String(lh),
        (v) => {
          // Clearing lineRule/lineHeightPx via Object.assign (exactOptionalPropertyTypes).
          const patch: Partial<ParaStyle> = { lineHeight: Number(v) };
          Object.assign(patch, { lineRule: undefined, lineHeightPx: undefined });
          editor.dispatch(setParaProps(patch));
        },
      ));
      insRow(sec, "Before (pt)", insNumber(toPt(ps?.spaceBeforePx ?? 0), 0, 200, (pt) => editor.dispatch(setParaProps({ spaceBeforePx: fromPt(Math.max(0, pt)) }))));
      insRow(sec, "After (pt)", insNumber(toPt(ps?.spaceAfterPx ?? 0), 0, 200, (pt) => editor.dispatch(setParaProps({ spaceAfterPx: fromPt(Math.max(0, pt)) }))));
    };

    // ---- Page / Section (live page-setup levers) ----------------------------
    // The common levers, applied live (no Apply); size/orientation/margins/columns.
    // The Page Layout dialog stays the advanced fallback (borders, page colour,
    // header/footer distance, line numbering) per the parity rule.
    const PL_SIZES: Record<string, { w: number; h: number }> = {
      Letter: { w: 816, h: 1056 },
      A4: { w: 794, h: 1123 },
      Legal: { w: 816, h: 1344 },
    };
    const PL_MARGINS: Record<string, number> = { Normal: 96, Narrow: 48, Wide: 144 };
    const applyPage = (mut: (g: ReturnType<typeof pageSetupAt>) => void): void => {
      const base = pageSetupAt({ doc: editor.getDocument(), selection: editor.getSelection() });
      mut(base);
      editor.dispatch(applyPageSetup(base));
    };
    const pageSizeName = (geo: ReturnType<typeof pageSetupAt>): string =>
      Object.keys(PL_SIZES).find((k) => {
        const s = PL_SIZES[k]!;
        return (s.w === geo.pageWidthPx && s.h === geo.pageHeightPx) || (s.h === geo.pageWidthPx && s.w === geo.pageHeightPx);
      }) ?? "Custom";
    const pageSummary = (geo: ReturnType<typeof pageSetupAt>): string => {
      const orient = geo.pageWidthPx > geo.pageHeightPx ? "Landscape" : "Portrait";
      const cm = Math.round((geo.marginPx.top / PX_PER_INCH) * 2.54 * 10) / 10;
      return `${pageSizeName(geo)} · ${orient} · ${cm} cm`;
    };
    const buildPageSection = (): void => {
      const geo = pageSetupAt({ doc: editor.getDocument(), selection: editor.getSelection() });
      const expanded = isExpanded("Page");
      const sec = insSection("Page", pageSummary(geo), expanded);
      if (!expanded) return;
      const landscape = geo.pageWidthPx > geo.pageHeightPx;
      const sizeName = pageSizeName(geo);
      insRow(sec, "Size", insSelect(
        [...Object.keys(PL_SIZES).map((k) => ({ value: k, label: k })), { value: "Custom", label: "Custom" }],
        sizeName,
        (v) => {
          const s = PL_SIZES[v];
          if (!s) return;
          applyPage((g) => { g.pageWidthPx = landscape ? s.h : s.w; g.pageHeightPx = landscape ? s.w : s.h; });
        },
      ));
      insRow(sec, "Orient", insToggles([
        { label: "Portrait", title: "Portrait", on: !landscape, onClick: () => { if (landscape) applyPage((g) => { [g.pageWidthPx, g.pageHeightPx] = [g.pageHeightPx, g.pageWidthPx]; }); } },
        { label: "Landscape", title: "Landscape", on: landscape, onClick: () => { if (!landscape) applyPage((g) => { [g.pageWidthPx, g.pageHeightPx] = [g.pageHeightPx, g.pageWidthPx]; }); } },
      ]));
      insRow(sec, "Margins", insToggles(Object.keys(PL_MARGINS).map((k) => ({
        label: k, title: `${k} margins`, on: false,
        onClick: () => { const m = PL_MARGINS[k]!; applyPage((g) => { g.marginPx = { top: m, right: m, bottom: m, left: m }; }); },
      }))));
      const inch = (px: number): number => Math.round((px / PX_PER_INCH) * 100) / 100;
      insRow(sec, "Top / Btm", insPair(
        insDecimal(inch(geo.marginPx.top), 0, 6, 0.1, (v) => applyPage((g) => { g.marginPx = { ...g.marginPx, top: Math.round(v * PX_PER_INCH) }; })),
        insDecimal(inch(geo.marginPx.bottom), 0, 6, 0.1, (v) => applyPage((g) => { g.marginPx = { ...g.marginPx, bottom: Math.round(v * PX_PER_INCH) }; })),
      ));
      insRow(sec, "Left / Rt", insPair(
        insDecimal(inch(geo.marginPx.left), 0, 6, 0.1, (v) => applyPage((g) => { g.marginPx = { ...g.marginPx, left: Math.round(v * PX_PER_INCH) }; })),
        insDecimal(inch(geo.marginPx.right), 0, 6, 0.1, (v) => applyPage((g) => { g.marginPx = { ...g.marginPx, right: Math.round(v * PX_PER_INCH) }; })),
      ));
      const colCount = geo.columns?.count ?? 1;
      insRow(sec, "Columns", insSelect(
        [1, 2, 3].map((n) => ({ value: String(n), label: String(n) })),
        String(Math.min(3, colCount)),
        (v) => {
          const n = Number(v);
          applyPage((g) => { g.columns = n <= 1 ? null : { count: n, gapPx: g.columns?.gapPx ?? 36, sep: g.columns?.sep ?? false }; });
        },
      ));
    };

    // ---- Table (live cell / row / table levers) -----------------------------
    // Shown when the caret is in a table cell: cell vertical alignment + text
    // direction, row height + flags, and table width / alignment / indent — each
    // applied live to the caret's cell or its table. Borders & Shading stays in the
    // dialog (advanced fallback) per the parity rule.
    const buildTableSection = (): void => {
      const doc = editor.getDocument();
      const sel = editor.getSelection();
      const loc = sel ? locateParagraph(doc, sel.focus.blockId) : null;
      if (!loc || loc.kind !== "cell") return;
      const table = containerListOf(doc, loc.where)[loc.bi] as TableBlock | undefined;
      if (!table || table.kind !== "table") return;
      const row: { cells: TableCell[]; props?: RowProps } | undefined = table.rows[loc.ri];
      const cell: TableCell | undefined = row?.cells[loc.ci];
      const rowCount = table.rows.length;
      const colCount = table.rows.reduce((m, r) => Math.max(m, r.cells.length), 0);
      const hasBorders = !!table.defaultBorders ||
        table.rows.some((r) => r.cells.some((c) => { const b = c.borders; return !!b && !!(b.top || b.bottom || b.left || b.right); }));
      const expanded = isExpanded("Table");
      const sec = insSection("Table", `${rowCount}×${colCount} · ${hasBorders ? "Grid" : "No borders"}`, expanded);
      if (!expanded) return;

      // Cell vertical alignment (top = clear, since it's the default).
      const vAlign = cell?.vAlign ?? "top";
      insRow(sec, "Cell", insToggles((["top", "center", "bottom"] as const).map((a) => ({
        label: a.charAt(0).toUpperCase() + a.slice(1), title: `Vertical align ${a}`, on: vAlign === a,
        onClick: () => editor.dispatch(setCellVAlignCmd(a === "top" ? null : a)),
      }))));

      // Cell text direction (horizontal = clear).
      const dirRaw = cell?.textDirection;
      const dir = dirRaw === "tbRl" || dirRaw === "btLr" ? dirRaw : "lrTb";
      insRow(sec, "Direction", insSelect(
        [{ value: "lrTb", label: "Horizontal" }, { value: "tbRl", label: "Rotate 90°" }, { value: "btLr", label: "Rotate 270°" }],
        dir,
        (v) => editor.dispatch(setCellTextDirectionCmd(v === "lrTb" ? null : (v as "tbRl" | "btLr"))),
      ));

      // Row height rule + value.
      const h = row?.props?.height ?? null;
      insRow(sec, "Row height", insSelect(
        [{ value: "none", label: "Auto" }, { value: "atLeast", label: "At least" }, { value: "exact", label: "Exactly" }],
        h ? h.rule : "none",
        (v) => {
          if (v === "none") editor.dispatch(setRowHeightAtSelectionCmd(null));
          else editor.dispatch(setRowHeightAtSelectionCmd({ value: h && h.value > 0 ? h.value : 40, rule: v as "atLeast" | "exact" }));
        },
      ));
      if (h) insRow(sec, "Height (px)", insNumber(Math.round(h.value), 1, 2000, (n) => {
        if (n > 0) editor.dispatch(setRowHeightAtSelectionCmd({ value: n, rule: h.rule }));
      }));

      // Row flags.
      insRow(sec, "Row", insToggles([
        { label: "Keep", title: "Keep row together (don't split across pages)", on: !!row?.props?.cantSplit, onClick: () => editor.dispatch(setRowPropsCmd({ cantSplit: !row?.props?.cantSplit })) },
        { label: "Header", title: "Repeat as header row on every page", on: !!row?.props?.repeatHeader, onClick: () => editor.dispatch(setRowPropsCmd({ repeatHeader: !row?.props?.repeatHeader })) },
      ]));

      // Table preferred width.
      const pw = table.preferredWidth ?? null;
      const wUnit = pw ? (pw.type === "pct" ? "pct" : "in") : "full";
      insRow(sec, "Width", insSelect(
        [{ value: "full", label: "Full width" }, { value: "pct", label: "% of page" }, { value: "in", label: "Inches" }],
        wUnit,
        (v) => {
          if (v === "full") editor.dispatch(setTablePreferredWidthAtSelectionCmd(null));
          else if (v === "pct") editor.dispatch(setTablePreferredWidthAtSelectionCmd({ type: "pct", value: pw?.type === "pct" ? pw.value : 50 }));
          else editor.dispatch(setTablePreferredWidthAtSelectionCmd({ type: "px", value: pw?.type === "px" ? pw.value : PX_PER_INCH * 4 }));
        },
      ));
      if (pw?.type === "pct") insRow(sec, "Width %", insNumber(Math.round(pw.value), 5, 100, (n) => editor.dispatch(setTablePreferredWidthAtSelectionCmd({ type: "pct", value: Math.min(100, Math.max(5, n)) }))));
      else if (pw?.type === "px") insRow(sec, "Width (in)", insDecimal(Math.round((pw.value / PX_PER_INCH) * 100) / 100, 0.5, 20, 0.1, (n) => editor.dispatch(setTablePreferredWidthAtSelectionCmd({ type: "px", value: Math.round(n * PX_PER_INCH) }))));

      // Table alignment.
      const tAlign = table.align ?? "left";
      insRow(sec, "Align", insToggles((["left", "center", "right"] as const).map((a) => ({
        label: a.charAt(0).toUpperCase(), title: `Table ${a}`, on: tAlign === a,
        onClick: () => editor.dispatch(setTableAlignAtSelectionCmd(a)),
      }))));

      // Table indent from the leading edge.
      insRow(sec, "Indent (px)", insNumber(Math.round(table.indentPx ?? 0), 0, 400, (n) => editor.dispatch(setTablePropsAtSelectionCmd({ indentPx: n > 0 ? n : null }))));
    };

    // ---- Object (selected image / shape: size, wrap, align, rotation) --------
    // Live size/wrap/align/rotation for the object-selected image or drawing shape,
    // routed to setImageProps / setShapeProps by kind. Fill/stroke/geometry stay on
    // the object's floating toolbar; the Shape Size & Position dialog remains the
    // advanced fallback for exact anchor offsets.
    const applyObject = (patch: { widthPx?: number; heightPx?: number; align?: "left" | "center" | "right"; wrap?: "block" | "square"; rotation?: number }): void => {
      const id = editor.getSelectedObject();
      const props = editor.getSelectedObjectProps();
      if (!id || !props) return;
      editor.dispatch(props.kind === "shape" ? setShapeProps(id, patch) : setImageProps(id, patch));
    };
    const buildObjectSection = (kind: "image" | "shape"): void => {
      const p = editor.getSelectedObjectProps();
      if (!p) {
        insEmpty("Use the object's floating toolbar or right-click menu to work with this selection.");
        return;
      }
      const name = kind === "shape" ? "Shape" : "Image";
      const summary = `${Math.round(p.widthPx)}×${Math.round(p.heightPx)} · ${p.wrap === "square" ? "Square" : "Inline"} wrap`;
      const expanded = isExpanded(name);
      const sec = insSection(name, summary, expanded);
      if (!expanded) return;
      insRow(sec, "Size (px)", insPair(
        insNumber(Math.round(p.widthPx), 8, 4000, (n) => applyObject({ widthPx: Math.max(8, n) })),
        insNumber(Math.round(p.heightPx), 8, 4000, (n) => applyObject({ heightPx: Math.max(8, n) })),
      ));
      insRow(sec, "Wrap", insToggles([
        { label: "Inline", title: "In line with text", on: p.wrap === "block", onClick: () => applyObject({ wrap: "block" }) },
        { label: "Float", title: "Float — text wraps around (square)", on: p.wrap === "square", onClick: () => applyObject({ wrap: "square" }) },
      ]));
      insRow(sec, "Align", insToggles((["left", "center", "right"] as const).map((a) => ({
        label: a.charAt(0).toUpperCase(), title: `Align ${a}`, on: p.align === a, onClick: () => applyObject({ align: a }),
      }))));
      insRow(sec, "Rotation°", insNumber(Math.round(p.rotation), 0, 359, (n) => applyObject({ rotation: ((Math.round(n) % 360) + 360) % 360 })));
    };

    const insEmpty = (text: string): void => {
      const e = el("div", "cw-insp-empty");
      e.textContent = text;
      bodyEl.append(e);
    };

    // --- breadcrumb scope trail + scope-changing selection helpers -----------
    // Clicking a crumb CHANGES the selection to that scope (via the existing
    // setSelection primitive), which re-derives the tightest-scope expansion — so
    // "click Table → whole table selected → Table section expands" falls out.
    const runsLen = (b: { kind?: string; runs?: { text?: string }[] } | null | undefined): number =>
      b && b.kind === "paragraph" && b.runs ? b.runs.reduce((n, r) => n + (r.text ? r.text.length : 0), 0) : 0;
    const firstParaOf = (cell: TableCell | undefined): { id: string } | null =>
      (cell?.blocks.find((b) => b.kind === "paragraph") as { id: string } | undefined) ?? null;
    const lastParaOf = (cell: TableCell | undefined): { id: string; kind?: string; runs?: { text?: string }[] } | null => {
      const ps = cell?.blocks.filter((b) => b.kind === "paragraph") as { id: string; kind?: string; runs?: { text?: string }[] }[] | undefined;
      return ps && ps.length ? ps[ps.length - 1]! : null;
    };
    const selectTable = (table: TableBlock | undefined): void => {
      if (!table) return;
      const first = firstParaOf(table.rows[0]?.cells[0]);
      const lastRow = table.rows[table.rows.length - 1];
      const last = lastParaOf(lastRow?.cells[(lastRow?.cells.length ?? 1) - 1]);
      if (!first || !last) return;
      editor.setSelection({ anchor: { blockId: first.id, offset: 0 }, focus: { blockId: last.id, offset: runsLen(last) } });
    };
    const selectCell = (cell: TableCell | undefined): void => {
      const p = firstParaOf(cell);
      if (p) editor.setSelection({ anchor: { blockId: p.id, offset: 0 }, focus: { blockId: p.id, offset: 0 } });
    };
    const buildBreadcrumb = (
      scope: "object" | "table" | "text" | "page",
      f: ReturnType<typeof editor.currentFormat>,
      doc: Document,
      loc: ReturnType<typeof locateParagraph> | null,
    ): void => {
      const focus = editor.getSelection()?.focus ?? null;
      const paraBlock = focus ? (paragraphsOf(doc).find((p) => p.id === focus.blockId) ?? null) : null;
      const crumbs: { label: string; action: (() => void) | null }[] = [];
      if (scope === "object") {
        crumbs.push({ label: f.shapeSelected ? "Shape" : "Image", action: null });
      } else {
        crumbs.push({ label: "Page", action: () => editor.setSelection(null) });
        if (scope === "table" && loc?.kind === "cell") {
          const table = containerListOf(doc, loc.where)[loc.bi] as TableBlock | undefined;
          crumbs.push({ label: "Table", action: () => selectTable(table) });
          crumbs.push({ label: "Cell", action: () => selectCell(table?.rows[loc.ri]?.cells[loc.ci]) });
        }
        if (paraBlock) {
          crumbs.push({ label: "Paragraph", action: () => editor.setSelection({ anchor: { blockId: paraBlock.id, offset: 0 }, focus: { blockId: paraBlock.id, offset: runsLen(paraBlock) } }) });
        }
        if (focus) crumbs.push({ label: "Text", action: () => editor.setSelection({ anchor: { ...focus }, focus: { ...focus } }) });
      }
      const rowEl = el("div", "cw-insp-crumbs");
      crumbs.forEach((c, i) => {
        if (i > 0) { const s = el("span", "cw-insp-crumb-sep"); s.textContent = "›"; rowEl.append(s); }
        const b = el("button", "cw-insp-crumb" + (i === crumbs.length - 1 ? " current" : ""));
        (b as HTMLButtonElement).type = "button";
        b.textContent = c.label;
        if (c.action) {
          b.addEventListener("mousedown", (e) => e.preventDefault());
          b.addEventListener("click", () => { c.action?.(); editor.focus(); });
        }
        rowEl.append(b);
      });
      bodyEl.append(rowEl);
    };

    let inspOpen = false;
    const rebuild = (): void => {
      bodyEl.textContent = "";
      const f = editor.currentFormat();
      const doc = editor.getDocument();
      const sel = editor.getSelection();
      const loc = sel ? locateParagraph(doc, sel.focus.blockId) : null;
      // Tightest scope the selection addresses (drives default expansion + the
      // breadcrumb). Table wins over text for a caret in a cell, per the IA spec.
      const scope: "object" | "table" | "text" | "page" =
        f.imageSelected || f.shapeSelected ? "object" : f.inTable ? "table" : sel ? "text" : "page";
      tightestSection = scope === "object" ? (f.shapeSelected ? "Shape" : "Image")
        : scope === "table" ? "Table" : scope === "text" ? "Text" : "Page";
      title.textContent = scope === "object" ? (f.shapeSelected ? "Shape" : "Image") : "Inspector";
      buildBreadcrumb(scope, f, doc, loc);
      if (scope === "object") { buildObjectSection(f.shapeSelected ? "shape" : "image"); return; }
      // Sections outer → inner: Page, Table, Paragraph, Text (only applicable ones).
      buildPageSection();
      if (scope === "table") buildTableSection();
      if (scope === "text" || scope === "table") { buildParagraphSection(f); buildTextSection(f); }
    };
    refreshInspector = (): void => {
      if (!inspOpen) return;
      // Don't clobber a field the user is mid-edit in; toggle buttons rebuild fine.
      const ae = document.activeElement;
      if (ae && inspectorEl.contains(ae) && (ae.tagName === "INPUT" || ae.tagName === "SELECT")) return;
      rebuild();
    };
    const setInspectorOpen = (v: boolean): void => {
      inspOpen = v;
      inspectorEl.classList.toggle("open", v);
      inspectorBtn.classList.toggle("active", v);
      if (v) rebuild();
    };
    toggleInspector = (): void => setInspectorOpen(!inspOpen);
    // Closed by default; opened via the Inspector header button.
  }

  // Collapse / expand the ribbon body (keeps the tab strip). A pinned chevron
  // on the right of the tab strip toggles it; clicking any tab re-pins.
  const collapseBtn = el("button", "rib-tab");
  // headerReview carries the single margin-left:auto that pushes the right
  // cluster to the corner; the collapse chevron just sits after it with a gap.
  collapseBtn.style.marginLeft = "6px";
  collapseBtn.innerHTML = chevron(true);
  const setCollapsed = (v: boolean): void => {
    toolbar.classList.toggle("collapsed", v);
    collapseBtn.innerHTML = chevron(!v);
    collapseBtn.title = v ? "Pin the ribbon" : "Collapse the ribbon (Ctrl+F1)";
  };
  collapseBtn.addEventListener("click", () => setCollapsed(!toolbar.classList.contains("collapsed")));
  window.addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === "F1") {
      e.preventDefault();
      setCollapsed(!toolbar.classList.contains("collapsed"));
    }
  }, { signal: teardown.signal });
  setCollapsed(false);
  tabsBar.appendChild(collapseBtn);

  // ===== customizeRibbon: reorder/remove built-ins + add custom tabs/buttons ===
  if (runtime.customizeRibbon) {
    const warn = (what: string, id: string): void =>
      console.warn(`[canvas-word] customizeRibbon: ${what} "${id}" not found`);
    // DOM is the source of truth for order; these read current ids in order.
    const tabOrder = (): string[] =>
      [...tabScroll.children].map((c) => (c as HTMLElement).dataset?.["ribbonTab"]).filter((x): x is string => !!x);
    const groupOrder = (tabId: string): string[] => {
      const panel = ribTabsById.get(tabId)?.panel;
      return panel
        ? [...panel.querySelectorAll(":scope > .rib-group")]
            .map((g) => (g as HTMLElement).dataset["ribbonGroup"])
            .filter((x): x is string => !!x)
        : [];
    };
    const itemOrder = (groupId: string): string[] => {
      const g = ribGroupsById.get(groupId)?.groupEl;
      return g
        ? [...g.querySelectorAll("[data-ribbon-item]")]
            .map((e) => (e as HTMLElement).dataset["ribbonItem"])
            .filter((x): x is string => !!x)
        : [];
    };
    const dropSync = (target: HTMLElement): void => {
      for (let i = toggleButtons.length - 1; i >= 0; i--) if (toggleButtons[i]!.el === target) toggleButtons.splice(i, 1);
      for (let i = enableButtons.length - 1; i >= 0; i--) if (enableButtons[i]!.el === target) enableButtons.splice(i, 1);
    };
    const buildCustomButton = (spec: RibbonButtonSpec): HTMLButtonElement => {
      const b = el("button", "rib-btn");
      if (spec.icon !== undefined) b.innerHTML = spec.icon;
      else {
        const s = el("span");
        s.textContent = spec.label ?? "";
        b.appendChild(s);
      }
      if (spec.tooltip) b.title = spec.tooltip;
      b.setAttribute("aria-label", spec.tooltip || spec.label || spec.id); // accessible name for custom buttons
      b.dataset["ribbonItem"] = spec.id;
      b.addEventListener("mousedown", (e) => e.preventDefault());
      b.addEventListener("click", () => {
        try {
          if (ribbonCtx) spec.onClick(ribbonCtx);
        } catch (err) {
          console.error("[canvas-word] custom ribbon button onClick threw", err);
        }
      });
      if (spec.active) toggleButtons.push({ el: b, active: spec.active });
      if (spec.enabled) enableButtons.push({ el: b, enabled: spec.enabled, title: spec.tooltip ?? "" });
      return b;
    };
    const api: RibbonApi = {
      tabs: () => tabOrder(),
      groups: (tabId) => groupOrder(tabId),
      items: (groupId) => itemOrder(groupId),
      addTab: (spec) => {
        if (ribTabsById.has(spec.id)) return warn("tab (already exists)", spec.id);
        const beforeId = anchorBeforeId(tabOrder(), spec, spec.id);
        const ref = (beforeId ? ribTabsById.get(beforeId)?.btn : null) ?? null;
        const panel = tab(spec.id, spec.label);
        tabScroll.insertBefore(ribTabsById.get(spec.id)!.btn, ref);
        void panel;
      },
      removeTab: (id) => {
        const node = ribTabsById.get(id);
        if (!node) return warn("tab", id);
        node.btn.remove();
        node.panel.remove();
        ribTabsById.delete(id);
        tabButtons.delete(id);
        tabPanels.delete(id);
      },
      moveTab: (id, pos) => {
        const node = ribTabsById.get(id);
        if (!node) return warn("tab", id);
        const beforeId = anchorBeforeId(tabOrder(), pos, id);
        const ref = (beforeId ? ribTabsById.get(beforeId)?.btn : null) ?? null;
        tabScroll.insertBefore(node.btn, ref);
      },
      addGroup: (tabId, spec) => {
        const tabNode = ribTabsById.get(tabId);
        if (!tabNode) return warn("tab", tabId);
        if (ribGroupsById.has(spec.id)) return warn("group (already exists)", spec.id);
        const g = el("div", "rib-group");
        const c = el("div", "rib-controls");
        const l = el("div", "rib-label");
        l.textContent = spec.label;
        g.append(c, l);
        g.dataset["ribbonGroup"] = spec.id;
        ribGroupsById.set(spec.id, { tabId, groupEl: g, container: c });
        const beforeId = anchorBeforeId(groupOrder(tabId), spec, spec.id);
        const ref = beforeId ? ribGroupsById.get(beforeId)?.groupEl ?? null : null;
        tabNode.panel.insertBefore(g, ref);
      },
      removeGroup: (id) => {
        const node = ribGroupsById.get(id);
        if (!node) return warn("group", id);
        for (const itemId of itemOrder(id)) {
          const it = ribItemsById.get(itemId);
          if (it) dropSync(it.el as HTMLElement);
          ribItemsById.delete(itemId);
        }
        node.groupEl.remove();
        ribGroupsById.delete(id);
      },
      moveGroup: (id, pos) => {
        const node = ribGroupsById.get(id);
        if (!node) return warn("group", id);
        const beforeId = anchorBeforeId(groupOrder(node.tabId), pos, id);
        const ref = beforeId ? ribGroupsById.get(beforeId)?.groupEl ?? null : null;
        node.groupEl.parentElement?.insertBefore(node.groupEl, ref);
      },
      addButton: (groupId, spec) => {
        const gnode = ribGroupsById.get(groupId);
        if (!gnode) return warn("group", groupId);
        if (ribItemsById.has(spec.id)) return warn("button (id already exists)", spec.id);
        const b = buildCustomButton(spec);
        ribItemsById.set(spec.id, { groupId, el: b });
        const beforeId = anchorBeforeId(itemOrder(groupId), spec, spec.id);
        const ref = beforeId ? ribItemsById.get(beforeId)?.el ?? null : null;
        if (ref && ref.parentElement) ref.parentElement.insertBefore(b, ref);
        else gnode.container.appendChild(b);
      },
      removeItem: (id) => {
        const node = ribItemsById.get(id);
        if (!node) return warn("item", id);
        dropSync(node.el as HTMLElement);
        node.el.remove();
        ribItemsById.delete(id);
      },
      moveItem: (id, pos) => {
        const node = ribItemsById.get(id);
        if (!node) return warn("item", id);
        const targetGroup = pos.toGroup ?? node.groupId;
        const gnode = ribGroupsById.get(targetGroup);
        if (!gnode) return warn("group", targetGroup);
        const beforeId = anchorBeforeId(itemOrder(targetGroup), pos, id);
        const ref = beforeId ? ribItemsById.get(beforeId)?.el ?? null : null;
        if (ref && ref.parentElement) ref.parentElement.insertBefore(node.el, ref);
        else gnode.container.appendChild(node.el);
        node.groupId = targetGroup;
      },
    };
    try {
      runtime.customizeRibbon(api);
    } catch (err) {
      console.error("[canvas-word] customizeRibbon threw", err);
    }
    // Custom buttons' active/enabled predicates sync on the next syncToolbar()
    // (wired to editor onChange + the initial sync below).
  }

  // Fit-to-width zoom: the largest zoom (≤100%) at which the page fits the scroll
  // area's width. Only shrinks — a page that already fits stays at 100% (R2).
  const fitWidthZoom = (): number => {
    const pageW = editor.getDocument().section.pageWidthPx;
    const avail = app.clientWidth - 40; // small gutter + scrollbar allowance
    if (pageW <= 0 || avail <= 0) return editor.getZoom();
    return Math.max(0.3, Math.min(1, avail / pageW));
  };
  const applyFitWidth = (): void => {
    const z = fitWidthZoom();
    if (Math.abs(z - editor.getZoom()) < 0.005) return;
    applyingFit = true;
    editor.setZoom(z);
    applyingFit = false;
  };

  // Container-width-driven compact ribbon: when the EDITOR itself is narrow (not
  // just the viewport — so it also works embedded in a narrow pane on a wide page),
  // switch the body to a dense, single horizontally-scrollable row. Mirrors the
  // `.collapsed` class pattern. Hysteresis (enter <720, exit ≥760) avoids flapping.
  if (typeof ResizeObserver !== "undefined") {
    const COMPACT_ON = 720;
    const COMPACT_OFF = 760;
    // Below this the outline pane + a 100%-zoom page no longer fit side by side.
    const COLLAPSE_ON = 1000;
    const COLLAPSE_OFF = 1120;
    const applyResponsive = (w: number): void => {
      // Compact ribbon (existing hysteresis).
      const on = toolbar.classList.contains("compact");
      if (!on && w < COMPACT_ON) toolbar.classList.add("compact");
      else if (on && w >= COMPACT_OFF) toolbar.classList.remove("compact");
      // Auto-collapse the outline when it and the page can't share the width
      // (R1/R4). Remembered, so it reopens once there's room — unless the user
      // toggled it manually in between (which clears the flag).
      if (w < COLLAPSE_ON && isOutlineOpen() && !autoCollapsedOutline) {
        autoCollapsedOutline = true;
        setOutlineOpen(false);
      } else if (w >= COLLAPSE_OFF && autoCollapsedOutline) {
        autoCollapsedOutline = false;
        setOutlineOpen(true);
      }
      // Keep the page fit-to-width while in auto-fit mode (R2).
      if (autoFit) applyFitWidth();
      // Tab-strip overflow indicator (R3).
      tabOverflowBtn.style.display = tabScroll.scrollWidth > tabScroll.clientWidth + 2 ? "" : "none";
    };
    compactRO = new ResizeObserver((entries) => {
      for (const e of entries) applyResponsive(e.contentRect.width);
    });
    compactRO.observe(shell.root);
    // Initial fit-width: default the zoom to fit the page when the viewport can't
    // hold it at 100% and the embedder didn't pin a zoom (R2).
    if (runtime.view?.zoom == null && fitWidthZoom() < 1) {
      autoFit = true;
      applyFitWidth();
    }
    applyResponsive(shell.root.getBoundingClientRect().width);
  }

  showTab("home");
  // After showTab (which re-pins), honor the embedder's initial collapsed choice.
  if (runtime.view?.ribbonCollapsed) setCollapsed(true);

  // ---- Command palette (Ctrl/Cmd+K) — every ribbon command reachable by name.
  // Gathers the live ribbon controls (enabled buttons only) plus any embedder
  // commands, so all ~145 actions are one keystroke away (critique Move 1).
  const gatherPaletteCommands = (): PaletteCommand[] => {
    const out: PaletteCommand[] = [];
    const seen = new Set<string>();
    for (const [id, item] of ribItemsById) {
      const el = item.el;
      if (!(el instanceof HTMLButtonElement) || el.disabled) continue;
      const label = (el.title || el.textContent || "").trim();
      if (!label) continue;
      if (seen.has(label)) continue; // the same command often sits on two tabs
      seen.add(label);
      const tabId = ribGroupsById.get(item.groupId)?.tabId;
      const hint = (tabId ? ribTabsById.get(tabId)?.btn.textContent : "") ?? "";
      out.push({ id, label, hint, run: () => el.click() });
    }
    for (const [id, cmd] of commandRegistry) out.push({ id, label: cmd.label ?? id, hint: "Command", run: () => runRegisteredCommand(id) });
    out.sort((a, b) => a.label.localeCompare(b.label));
    return out;
  };
  window.addEventListener(
    "keydown",
    (e) => {
      if (!(e.ctrlKey || e.metaKey) || e.altKey || e.key.toLowerCase() !== "k") return;
      e.preventDefault();
      showCommandPalette(gatherPaletteCommands());
    },
    { signal: teardown.signal },
  );

  // ---- Minimal chrome preset (Move 1) -------------------------------------
  // `chrome: 'minimal'` demotes the ribbon to a quiet ~44px command bar: hide the
  // ribbon body + tab strip + collapse chevron (the identity/save, quick-access and
  // Inspector/Review header clusters stay), and drop a compact command cluster into
  // the freed space — style picker, the six core formatting commands, an insert ＋
  // menu, and a ⋯ overflow that opens the command palette. Everything else reaches
  // the user through the contextual bar, the Inspector and the palette. The full
  // ribbon is still built (so all command wiring, popovers and syncToolbar state
  // remain live); only its body/tabs are hidden, so the preset is a pure skin swap.
  if (config.chrome === "minimal") {
    toolbar.classList.add("cw-chrome-minimal");
    bodies.style.display = "none";
    tabScroll.style.display = "none";
    tabOverflowBtn.style.display = "none";
    collapseBtn.style.display = "none";

    const bar = el("div", "cw-minibar");

    const styleSel = document.createElement("select");
    styleSel.className = "cw-minibar-style";
    styleSel.title = "Paragraph / character style";
    styleSel.setAttribute("aria-label", "Style");
    const fillStyleOptions = (): void => {
      const cur = styleSel.value;
      styleSel.textContent = "";
      for (const s of stylesheet().styles) {
        const o = document.createElement("option");
        o.value = s.id;
        o.textContent = s.name;
        styleSel.appendChild(o);
      }
      if (cur) styleSel.value = cur;
    };
    fillStyleOptions();
    styleSel.addEventListener("mousedown", () => fillStyleOptions()); // stay fresh after an import swaps the stylesheet
    styleSel.addEventListener("change", () => {
      const id = styleSel.value;
      const st = styleById(stylesheet(), id);
      editor.dispatch(st && styleType(st) === "character" ? applyCharStyle(id) : applyNamedStyle(id));
      editor.focus();
    });
    bar.appendChild(styleSel);
    bar.appendChild(el("span", "cw-minibar-sep"));

    const glyphBtn = (html: string, title: string, onClick: () => void): HTMLButtonElement => {
      const b = el("button", "cw-minibar-btn");
      b.innerHTML = html;
      b.title = title;
      b.setAttribute("aria-label", title);
      b.addEventListener("mousedown", (ev) => ev.preventDefault());
      b.addEventListener("click", () => { onClick(); editor.focus(); refreshMinimalBar(); });
      bar.appendChild(b);
      return b;
    };
    const bBold = glyphBtn(`<span style="font-weight:800">B</span>`, "Bold (Ctrl+B)", () => editor.toggleStyle("bold"));
    const bItalic = glyphBtn(`<span style="font-style:italic;font-weight:600">I</span>`, "Italic (Ctrl+I)", () => editor.toggleStyle("italic"));
    const bUnder = glyphBtn(`<span style="text-decoration:underline;font-weight:600">U</span>`, "Underline (Ctrl+U)", () => editor.toggleStyle("underline"));
    const bStrike = glyphBtn(`<span style="text-decoration:line-through;font-weight:600">S</span>`, "Strikethrough", () => editor.toggleStyle("strikethrough"));
    const bBullet = glyphBtn(ICONS.bullets, "Bulleted list", () => editor.dispatch(toggleList("bullet")));
    const bNumber = glyphBtn(ICONS.numbering, "Numbered list", () => editor.dispatch(toggleList("decimal")));
    bar.appendChild(el("span", "cw-minibar-sep"));

    const insertBtn = el("button", "cw-minibar-btn cw-minibar-insert");
    insertBtn.innerHTML = `<span style="font-size:17px;line-height:1">＋</span>`;
    insertBtn.title = "Insert";
    insertBtn.setAttribute("aria-label", "Insert");
    insertBtn.addEventListener("mousedown", (ev) => ev.preventDefault());
    insertBtn.addEventListener("click", () => {
      const r = insertBtn.getBoundingClientRect();
      const entries: MenuEntry[] = [
        { kind: "item", label: "Table", onClick: () => editor.dispatch(insertTable(3, 3)) },
        { kind: "item", label: "Picture…", onClick: () => pickAndInsertImage() },
        { kind: "item", label: "Page break", onClick: () => editor.dispatch(insertPageBreak()) },
        { kind: "item", label: "Table of contents", onClick: () => editor.dispatch(insertTocCmd()) },
        { kind: "item", label: "Footnote", onClick: () => editor.dispatch(insertFootnoteCmd()) },
        { kind: "sep" },
        { kind: "item", label: "More commands…", onClick: () => showCommandPalette(gatherPaletteCommands()) },
      ];
      showContextMenu(r.left, r.bottom + 2, entries);
    });
    bar.appendChild(insertBtn);

    const moreBtn = el("button", "cw-minibar-btn");
    moreBtn.innerHTML = `<span style="font-size:18px;line-height:1">⋯</span>`;
    moreBtn.title = "More commands (Ctrl+K)";
    moreBtn.setAttribute("aria-label", "More commands");
    moreBtn.addEventListener("mousedown", (ev) => ev.preventDefault());
    moreBtn.addEventListener("click", () => showCommandPalette(gatherPaletteCommands()));
    bar.appendChild(moreBtn);

    tabsBar.insertBefore(bar, tabScroll);

    refreshMinimalBar = (): void => {
      const f = editor.currentFormat();
      bBold.classList.toggle("on", f.bold);
      bItalic.classList.toggle("on", f.italic);
      bUnder.classList.toggle("on", f.underline);
      bStrike.classList.toggle("on", f.strikethrough);
      bBullet.classList.toggle("on", f.listKind === "bullet");
      bNumber.classList.toggle("on", f.listKind === "number");
      if (f.styleId && Array.from(styleSel.options).some((o) => o.value === f.styleId)) styleSel.value = f.styleId;
    };
    refreshMinimalBar();
  }

  // ---- toolbar controls mirror the caret formatting -----------------------
  let lastStylesheet = editor.getDocument().stylesheet ?? null;
  let lastUsedSig = ""; // signature of the in-use style set (for the gallery filter)
  syncToolbar = (): void => {
    // A .docx import swaps the whole stylesheet — rebuild the gallery so
    // imported style ids (Heading 1, …) resolve instead of sticking on Normal.
    const sheet = editor.getDocument().stylesheet ?? null;
    if (sheet !== lastStylesheet) {
      lastStylesheet = sheet;
      rebuildStyleGallery();
      styleMgr?.refresh();
    } else if (hideUnusedStyles) {
      // While the "in use" filter is on, keep it live as usage changes (applying
      // or clearing a style). Cheap signature check — only runs while filtering.
      const sig = [...usedStyleIds()].sort().join(",");
      if (sig !== lastUsedSig) { lastUsedSig = sig; rebuildStyleGallery(); }
    }
    const f = editor.currentFormat();
    if (f.styleId) {
      // Paragraph references a style the gallery doesn't list (importer kept the
      // ref but the sheet lacks it) — surface it rather than lying.
      if (!styleCards.has(f.styleId)) addStyleCard(f.styleId, f.styleId);
      for (const [id, c] of styleCards) c.classList.toggle("active", id === f.styleId);
    }
    if (f.fontFamily) {
      const match = toolbarFonts(config.fonts).find((x) => x.value.toLowerCase() === f.fontFamily!.toLowerCase());
      if (match) fontSelect.value = match.value;
    }
    if (f.fontSizePx !== null && document.activeElement !== sizeInput) {
      sizeInput.value = String(pxToPt(f.fontSizePx));
    }
    curLineHeight = f.lineHeight; // read by the line-spacing menu when opened
    // Pressed state for B/I/U/S, highlight, sub/super, lists, alignment.
    for (const t of toggleButtons) t.el.classList.toggle("active", t.active(f));
    // Enabled state for context-dependent buttons (image, content control, …).
    for (const e of enableButtons) {
      const on = e.enabled(f);
      e.el.disabled = !on;
      const t = on || !e.hint ? e.title : `${e.title} — ${e.hint}`;
      e.el.title = t;
      e.el.setAttribute("aria-label", t); // keep the accessible name in step with the tooltip
    }
    syncQat(); // grey the quick-access undo/redo to match the stack
    // Reveal/hide contextual tabs by selection; if the active tab just hid, fall
    // back to Home so the user is never stranded on an invisible tab.
    for (const ct of contextualTabs) {
      const show = ct.pred(f);
      ct.btn.style.display = show ? "" : "none";
      if (!show && ct.btn.classList.contains("active")) showTab("home");
    }
  };

  // Reflect the live zoom (also driven by Ctrl+wheel): snap the select to the
  // nearest preset, or insert a one-off "NN%" option for in-between values.
  syncZoom = (z: number): void => {
    const pct = `${Math.round(z * 100)}%`;
    const preset = [...zoomSel.options].find((o) => o.textContent === pct);
    [...zoomSel.options].filter((o) => o.dataset["custom"]).forEach((o) => o.remove());
    if (preset) {
      zoomSel.value = preset.value;
    } else {
      const o = el("option");
      o.value = String(z);
      o.textContent = pct;
      o.dataset["custom"] = "1";
      zoomSel.appendChild(o);
      zoomSel.value = String(z);
    }
  };

  syncToolbar(); // set initial pressed/enabled state before the first edit
}

// ---- status bar (page count, word/character count, zoom slider) -------------
{
  const sb = shell.statusbar;
  if (runtime.view?.statusBar === false) sb.style.display = "none";
  if (sb) {
    const left = document.createElement("div");
    left.className = "sb-left";
    const right = document.createElement("div");
    right.className = "sb-right";
    sb.append(left, right);
    const item = (parent: HTMLElement): HTMLSpanElement => {
      const s = document.createElement("span");
      s.className = "sb-item";
      parent.appendChild(s);
      return s;
    };
    const pageItem = item(left);
    const wordItem = item(left);

    const zBtn = (label: string, title: string, onClick: () => void): void => {
      const b = document.createElement("button");
      b.className = "sb-btn";
      b.textContent = label;
      b.title = title;
      b.addEventListener("click", onClick);
      right.appendChild(b);
    };
    zBtn("−", "Zoom out", () => editor.setZoom(editor.getZoom() / config.behavior.zoomStep));
    const slider = document.createElement("input");
    slider.type = "range";
    slider.min = "50";
    slider.max = "300";
    slider.step = "10";
    slider.title = "Zoom";
    slider.addEventListener("input", () => editor.setZoom(Number(slider.value) / 100));
    right.appendChild(slider);
    zBtn("+", "Zoom in", () => editor.setZoom(editor.getZoom() * config.behavior.zoomStep));
    const zLabel = document.createElement("span");
    zLabel.className = "sb-item sb-zoom";
    right.appendChild(zLabel);

    // Page shown is the one the user is LOOKING AT (scroll), not where the caret
    // is: the deepest page whose top has scrolled above a line in the upper view.
    let pageCount = 1;
    const setPageItem = (): void => {
      const phs = app.querySelectorAll<HTMLElement>("[data-page]");
      let current = 1;
      if (phs.length) {
        const appRect = app.getBoundingClientRect();
        const refLine = appRect.top + appRect.height * 0.3;
        phs.forEach((ph) => {
          if (ph.getBoundingClientRect().top <= refLine) current = Number(ph.dataset["page"]) + 1;
        });
      }
      pageItem.textContent = `Page ${Math.min(current, pageCount)} of ${pageCount}`;
    };

    refreshStatus = (): void => {
      pageCount = editor.getLayoutInfo().pageCount;
      let chars = 0;
      let words = 0;
      for (const p of paragraphsOf(editor.getDocument())) {
        const t = textOfRuns(p.runs);
        chars += t.length;
        const m = t.match(/\S+/g);
        if (m) words += m.length;
      }
      wordItem.textContent = `${words} word${words === 1 ? "" : "s"} · ${chars} character${chars === 1 ? "" : "s"}`;
      const pct = Math.round(editor.getZoom() * 100);
      slider.value = String(Math.min(300, Math.max(50, pct)));
      zLabel.textContent = `${pct}%`;
      setPageItem();
    };
    // Scrolling only re-evaluates the visible page (cheap) — not the word count.
    app.addEventListener("scroll", setPageItem, { passive: true });
    window.addEventListener("resize", setPageItem, { signal: teardown.signal });
    refreshStatus();
  }
}

// ---- horizontal ruler (inch ticks, margin shading, draggable indents) -------
{
  // View-only: the ruler's only interactions are draggable indent markers, so
  // hide it (refreshRuler also early-returns while hidden).
  const rulerRow = shell.ruler.parentElement as HTMLElement; // .cw-ruler-row
  const rulerCorner = rulerRow.firstElementChild as HTMLElement; // .cw-ruler-corner (spacer over the vruler)
  const ruler = readonly ? null : shell.ruler;
  // Apply initial visibility: hide the whole ruler-row when the horizontal ruler
  // is off (otherwise it leaves an empty 22px bar), and hide the corner spacer
  // when the vertical ruler is off so the horizontal ruler aligns with .cw-app.
  if (!showHRuler) {
    rulerRow.classList.add("hidden");
    shell.ruler.classList.add("hidden");
  }
  if (!showVRuler) {
    shell.vruler.classList.add("hidden");
    rulerCorner.classList.add("hidden");
  }
  if (ruler) {
    const canvas = document.createElement("canvas");
    const leftMarker = document.createElement("div");
    leftMarker.className = "ruler-marker ruler-left";
    leftMarker.title = "Left indent — drag to set";
    const firstMarker = document.createElement("div");
    firstMarker.className = "ruler-marker ruler-first";
    firstMarker.title = "First-line indent — drag to set";
    ruler.append(canvas, leftMarker, firstMarker);

    // Geometry captured each refresh, read by the drag handlers.
    let geom: { contentLeft: number; zoom: number; paraId: string | null; li: number; fi: number } | null = null;

    const draw = (W: number, H: number, pageLeft: number, pageW: number, cL: number, cR: number, zoom: number): void => {
      const ctx = canvas.getContext("2d");
      if (!ctx) return;
      const dpr = window.devicePixelRatio || 1;
      canvas.width = Math.round(W * dpr);
      canvas.height = Math.round(H * dpr);
      canvas.style.width = `${W}px`;
      canvas.style.height = `${H}px`;
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      ctx.clearRect(0, 0, W, H);
      // page band (margins greyed, content white)
      ctx.fillStyle = config.theme.ruler.bg;
      ctx.fillRect(pageLeft, 4, pageW, H - 8);
      ctx.fillStyle = config.theme.ruler.content;
      ctx.fillRect(cL, 4, cR - cL, H - 8);
      // inch ticks + numbers, measured from the left content edge (Word's 0)
      // Ticks hang from the top of the ruler band; numbers sit in the lower half
      // so they never overlap the tick lines.
      const inch = PX_PER_INCH * zoom;
      const tickTop = RULER_TICK_START; // top of the page band
      const tickBot = RULER_TICK_LEN;   // bottom of the major tick (6 px, upper half only)
      const halfTickBot = RULER_HALF_TICK_LEN; // bottom of the half-inch tick (4 px)
      const numY = H - 4;       // number baseline anchored to the bottom of the band
      ctx.strokeStyle = config.theme.ruler.line;
      ctx.fillStyle = config.theme.ruler.label;
      ctx.font = config.theme.ruler.font;
      ctx.textAlign = "center";
      ctx.textBaseline = "bottom";
      ctx.lineWidth = 1;
      for (let i = 1; cL + i * inch < cR - 2; i++) {
        const x = Math.round(cL + i * inch) + 0.5;
        ctx.beginPath();
        ctx.moveTo(x, tickTop);
        ctx.lineTo(x, tickBot);
        ctx.stroke();
        ctx.fillText(String(i), x, numY);
        // mid tick at the half-inch
        const hx = Math.round(cL + (i - 0.5) * inch) + 0.5;
        ctx.beginPath();
        ctx.moveTo(hx, tickTop);
        ctx.lineTo(hx, halfTickBot);
        ctx.stroke();
      }
    };

    refreshRuler = (): void => {
      if (ruler.classList.contains("hidden")) return;
      const ph = app.querySelector<HTMLElement>("[data-page]");
      if (!ph) {
        leftMarker.style.display = firstMarker.style.display = "none";
        return;
      }
      const appRect = app.getBoundingClientRect();
      const phRect = ph.getBoundingClientRect();
      const zoom = editor.getZoom();
      const sec = editor.getDocument().section;
      const pageLeft = phRect.left - appRect.left;
      const pageW = phRect.width;
      const cL = pageLeft + sec.marginPx.left * zoom;
      const cR = pageLeft + pageW - sec.marginPx.right * zoom;
      draw(ruler.clientWidth, ruler.clientHeight, pageLeft, pageW, cL, cR, zoom);
      // indent markers for the caret paragraph
      const sel = editor.getSelection();
      const para = sel ? paragraphsOf(editor.getDocument()).find((p) => p.id === sel.focus.blockId) : null;
      if (para) {
        const li = para.style.indentLeftPx ?? 0;
        const fi = para.style.indentFirstLinePx ?? 0;
        leftMarker.style.left = `${cL + li * zoom - 5}px`;
        firstMarker.style.left = `${cL + (li + fi) * zoom - 5}px`;
        leftMarker.style.display = firstMarker.style.display = "block";
        geom = { contentLeft: cL, zoom, paraId: para.id, li, fi };
      } else {
        leftMarker.style.display = firstMarker.style.display = "none";
        geom = null;
      }
    };

    const startDrag = (which: "left" | "first") => (e: PointerEvent): void => {
      const g = geom;
      if (!g || !g.paraId) return;
      e.preventDefault();
      let liNext = g.li;
      let fiNext = g.fi;
      const onMove = (ev: PointerEvent): void => {
        const x = ev.clientX - app.getBoundingClientRect().left;
        const px = Math.max(0, (x - g.contentLeft) / g.zoom);
        if (which === "left") {
          liNext = px;
          leftMarker.style.left = `${g.contentLeft + liNext * g.zoom - 5}px`;
          firstMarker.style.left = `${g.contentLeft + (liNext + fiNext) * g.zoom - 5}px`;
        } else {
          fiNext = px - liNext; // first line sits at left-indent + first-line-indent
          firstMarker.style.left = `${g.contentLeft + px * g.zoom - 5}px`;
        }
      };
      const onUp = (): void => {
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
        editor.dispatch(
          setParaProps(which === "left" ? { indentLeftPx: Math.round(liNext) } : { indentFirstLinePx: Math.round(fiNext) }),
        );
        editor.focus();
      };
      window.addEventListener("pointermove", onMove);
      window.addEventListener("pointerup", onUp);
    };
    leftMarker.addEventListener("pointerdown", startDrag("left"));
    firstMarker.addEventListener("pointerdown", startDrag("first"));

    app.addEventListener("scroll", () => refreshRuler());
    window.addEventListener("resize", () => refreshRuler(), { signal: teardown.signal });
    // expose a toggle for the View-tab button — toggles the whole ruler-row (so
    // the corner spacer doesn't linger as an empty bar) and the ruler in lockstep.
    toggleRuler = (): boolean => {
      const hidden = rulerRow.classList.toggle("hidden");
      ruler.classList.toggle("hidden", hidden);
      if (!hidden) refreshRuler();
      return !hidden;
    };
    refreshRuler();
  }

  // ---- vertical ruler (inch ticks + draggable top/bottom margin markers) ----
  // A mirror of the horizontal ruler down the left of the scroll area, with
  // draggable handles at the page's top/bottom margin boundaries (Word's vertical
  // ruler resize). Tracks the topmost visible page so it stays meaningful while
  // scrolling.
  if (!readonly) {
    const vruler = shell.vruler;
    const vcanvas = document.createElement("canvas");
    const topMarker = document.createElement("div");
    topMarker.className = "vruler-marker vruler-top";
    topMarker.title = "Top margin — drag to set";
    const bottomMarker = document.createElement("div");
    bottomMarker.className = "vruler-marker vruler-bottom";
    bottomMarker.title = "Bottom margin — drag to set";
    vruler.append(vcanvas, topMarker, bottomMarker);

    // Geometry captured each refresh, read by the margin drag handlers.
    let vgeom: { pageTop: number; pageH: number; zoom: number; topPx: number; bottomPx: number } | null = null;

    const drawV = (W: number, H: number, pageTop: number, pageH: number, cT: number, cB: number, zoom: number): void => {
      const ctx = vcanvas.getContext("2d");
      if (!ctx) return;
      const dpr = window.devicePixelRatio || 1;
      vcanvas.width = Math.round(W * dpr);
      vcanvas.height = Math.round(H * dpr);
      vcanvas.style.width = `${W}px`;
      vcanvas.style.height = `${H}px`;
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      ctx.clearRect(0, 0, W, H);
      // page band (margins greyed, content white) — mirror of the horizontal draw
      ctx.fillStyle = config.theme.ruler.bg;
      ctx.fillRect(4, pageTop, W - 8, pageH);
      ctx.fillStyle = config.theme.ruler.content;
      ctx.fillRect(4, cT, W - 8, cB - cT);
      // inch ticks + numbers, measured from the top content edge (Word's 0)
      const inch = PX_PER_INCH * zoom;
      const tickLeft = RULER_TICK_START;
      const tickRight = RULER_TICK_LEN; // major tick (6 px)
      const halfTickRight = RULER_HALF_TICK_LEN; // half-inch tick (4 px)
      const numX = W - 4;
      ctx.strokeStyle = config.theme.ruler.line;
      ctx.fillStyle = config.theme.ruler.label;
      ctx.font = config.theme.ruler.font;
      ctx.textAlign = "right";
      ctx.textBaseline = "middle";
      ctx.lineWidth = 1;
      for (let i = 1; cT + i * inch < cB - 2; i++) {
        const y = Math.round(cT + i * inch) + 0.5;
        ctx.beginPath();
        ctx.moveTo(tickLeft, y);
        ctx.lineTo(tickRight, y);
        ctx.stroke();
        ctx.fillText(String(i), numX, y);
        const hy = Math.round(cT + (i - 0.5) * inch) + 0.5;
        ctx.beginPath();
        ctx.moveTo(tickLeft, hy);
        ctx.lineTo(halfTickRight, hy);
        ctx.stroke();
      }
    };

    refreshVRuler = (): void => {
      if (vruler.classList.contains("hidden")) return;
      const appRect = app.getBoundingClientRect();
      // Topmost page still in view (its bottom past the scroll area's top), so the
      // ruler tracks the page you're reading rather than always page 1.
      let ph: HTMLElement | null = null;
      for (const el of app.querySelectorAll<HTMLElement>("[data-page]")) {
        if (el.getBoundingClientRect().bottom > appRect.top) {
          ph = el;
          break;
        }
      }
      if (!ph) {
        topMarker.style.display = bottomMarker.style.display = "none";
        vgeom = null;
        return;
      }
      const vRect = vruler.getBoundingClientRect();
      const phRect = ph.getBoundingClientRect();
      const zoom = editor.getZoom();
      const sec = editor.getDocument().section;
      const pageTop = phRect.top - vRect.top;
      const pageH = phRect.height;
      const cT = pageTop + sec.marginPx.top * zoom;
      const cB = pageTop + pageH - sec.marginPx.bottom * zoom;
      drawV(vruler.clientWidth, vruler.clientHeight, pageTop, pageH, cT, cB, zoom);
      // Position the margin handles at the content boundaries (centered on them).
      topMarker.style.top = `${cT - 5}px`;
      bottomMarker.style.top = `${cB - 5}px`;
      topMarker.style.display = bottomMarker.style.display = "block";
      vgeom = { pageTop, pageH, zoom, topPx: sec.marginPx.top, bottomPx: sec.marginPx.bottom };
    };

    const startVDrag = (which: "top" | "bottom") => (e: PointerEvent): void => {
      const g = vgeom;
      if (!g) return;
      e.preventDefault();
      let topNext = g.topPx;
      let bottomNext = g.bottomPx;
      const pageHDoc = g.pageH / g.zoom; // page height in document px
      const MIN_CONTENT = 48; // keep at least this much white between margins
      const onMove = (ev: PointerEvent): void => {
        const y = ev.clientY - vruler.getBoundingClientRect().top; // ruler-local screen-y
        if (which === "top") {
          const maxTop = pageHDoc - bottomNext - MIN_CONTENT;
          topNext = Math.max(0, Math.min(maxTop, (y - g.pageTop) / g.zoom));
          topMarker.style.top = `${g.pageTop + topNext * g.zoom - 5}px`;
        } else {
          const pageBottom = g.pageTop + g.pageH;
          const maxBottom = pageHDoc - topNext - MIN_CONTENT;
          bottomNext = Math.max(0, Math.min(maxBottom, (pageBottom - y) / g.zoom));
          bottomMarker.style.top = `${pageBottom - bottomNext * g.zoom - 5}px`;
        }
      };
      const onUp = (): void => {
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
        const base = pageSetupAt({ doc: editor.getDocument(), selection: editor.getSelection() });
        editor.dispatch(
          applyPageSetup({
            ...base,
            marginPx: { ...base.marginPx, top: Math.round(topNext), bottom: Math.round(bottomNext) },
          }),
        );
        editor.focus();
      };
      window.addEventListener("pointermove", onMove);
      window.addEventListener("pointerup", onUp);
    };
    topMarker.addEventListener("pointerdown", startVDrag("top"));
    bottomMarker.addEventListener("pointerdown", startVDrag("bottom"));

    app.addEventListener("scroll", () => refreshVRuler());
    window.addEventListener("resize", () => refreshVRuler(), { signal: teardown.signal });
    toggleVRuler = (): boolean => {
      const hidden = vruler.classList.toggle("hidden");
      rulerCorner.classList.toggle("hidden", hidden); // corner spacer tracks the vruler
      refreshRuler(); // the horizontal ruler shifts as the corner appears/disappears
      if (!hidden) refreshVRuler();
      return !hidden;
    };
    refreshVRuler();
  }
}

// ---- contextual floating toolbars ------------------------------------------
// One priority-based manager shows the single most-relevant floating bar above the
// caret/selection: the selected-image bar, the hyperlink bar, the text format bar,
// and any embedder-registered `contextToolbars`. Edit-only (view mode has nothing
// to format). The manager owns the scroll/resize/Escape lifecycle and the mutual
// exclusion (highest priority wins) that used to be hand-wired per bar.
if (!readonly) {
  const isRangeSelection = (): boolean => {
    const sel = editor.getSelection();
    if (!sel) return false;
    return sel.anchor.blockId !== sel.focus.blockId || sel.anchor.offset !== sel.focus.offset;
  };
  const clampSizePx = (pt: number): number => Math.min(96, Math.max(6, sharedPtToPx(pt)));
  // A custom button gets the same context a custom ribbon button gets; a throwing
  // handler is caught rather than surfaced (mirrors the ribbon path).
  const runCustomButton = (spec: FloatingToolbarButtonSpec): void => {
    try {
      if (ribbonCtx) spec.onClick(ribbonCtx);
    } catch (err) {
      console.error("[canvas-word] custom context-toolbar button onClick threw", err);
    }
  };

  const manager = createContextToolbarManager({ scrollEl: app, signal: teardown.signal });

  // Image bar (priority 30) — the highest, so a selected image always wins.
  // Run `fn` with the selected object's id (the image or equation the bar acts on).
  const withSelectedObject = (fn: (id: string) => void): void => {
    const id = editor.getSelectedObject();
    if (id) fn(id);
  };
  manager.register(
    createImageContextToolbar({
      anchorRect: () => editor.getSelectedObjectRect(),
      actions: {
        wrapInline: () => withSelectedObject((id) => editor.dispatch(setImageProps(id, { wrap: "block", align: "center" }))),
        wrapSquare: () => withSelectedObject((id) => editor.dispatch(setImageProps(id, { wrap: "square", align: "left" }))),
        alignLeft: () => editor.align("left"),
        alignCenter: () => editor.align("center"),
        alignRight: () => editor.align("right"),
        remove: () => {
          editor.deleteSelectedObject();
          editor.focus();
        },
      },
    }),
  );

  // Shape bar (priority 31) — a selected drawing shape (wins over the image bar).
  manager.register(
    createShapeContextToolbar({
      anchorRect: () => editor.getSelectedShapeRect(),
      canGroup: () => editor.canGroupShapes(),
      canUngroup: () => editor.canUngroupShape(),
      actions: {
        fill: (btn) => shapeFillPopover(btn),
        outline: (btn) => shapeOutlinePopover(btn),
        addText: () => { editor.addTextToSelectedShape(); },
        wrapInline: () => withSelectedObject((id) => editor.dispatch(setShapeProps(id, { wrap: "block", align: "center", anchor: null }))),
        wrapSquare: () => withSelectedObject((id) => editor.dispatch(setShapeProps(id, { wrap: "square", align: "left", anchor: null }))),
        layerBehind: () => withSelectedObject((id) => editor.dispatch(setShapeLayer(id, true))),
        layerInFront: () => withSelectedObject((id) => editor.dispatch(setShapeLayer(id, false))),
        alignLeft: () => editor.align("left"),
        alignCenter: () => editor.align("center"),
        alignRight: () => editor.align("right"),
        bringToFront: () => withSelectedObject((id) => editor.dispatch(bringShapeToFront(id))),
        sendToBack: () => withSelectedObject((id) => editor.dispatch(sendShapeToBack(id))),
        group: () => { editor.groupShapes(); editor.focus(); },
        ungroup: () => { editor.ungroupShape(); editor.focus(); },
        remove: () => {
          editor.deleteSelectedObject();
          editor.focus();
        },
      },
    }),
  );

  // Table bar (priority 28) — a multi-cell selection (2+ rows/columns).
  manager.register(
    createTableContextToolbar({
      anchorRect: () => editor.getCellSelectionRect(),
      actions: {
        mergeCells: () => editor.dispatch(mergeCellsCmd()),
        insertRowAbove: () => editor.dispatch(insertTableRowCmd("above")),
        insertRowBelow: () => editor.dispatch(insertTableRowCmd("below")),
        insertColumnLeft: () => editor.dispatch(insertTableColumnCmd("left")),
        insertColumnRight: () => editor.dispatch(insertTableColumnCmd("right")),
        deleteRow: () => editor.dispatch(deleteTableRowCmd()),
        deleteColumn: () => editor.dispatch(deleteTableColumnCmd()),
      },
    }),
  );

  // Hyperlink bar (priority 25) — a collapsed caret inside a link.
  manager.register(
    createLinkContextToolbar({
      linkUrl: () => editor.linkAtCaret(),
      hasRangeSelection: isRangeSelection,
      inToc: () => editor.caretInToc(),
      anchorRect: () => editor.getSelectionAnchorRect(),
      actions: {
        open: (url) => window.open(url, "_blank", "noopener,noreferrer"),
        edit: (anchorEl, url) => openLinkDialog(anchorEl, url),
        copy: (url) => void navigator.clipboard?.writeText(url),
        remove: () => {
          editor.dispatch(setLinkCmd(null));
          editor.focus();
        },
      },
    }),
  );

  // Text format bar (priority 20) — over a range (always) or a caret (config.onCaret).
  if (config.floatingToolbar.enabled) {
    manager.register(
      createFloatingFormatBar({
        anchorRect: () => editor.getSelectionAnchorRect(),
        format: () => editor.currentFormat(),
        hasRangeSelection: isRangeSelection,
        fonts: () => toolbarFonts(config.fonts),
        pxToPt: (px) => Math.round(sharedPxToPt(px) * 2) / 2,
        focus: () => editor.focus(),
        config: config.floatingToolbar,
        runCustomButton,
        actions: {
          toggle: (key) => editor.toggleStyle(key),
          setFontFamily: (value) => editor.setCharStyle({ fontFamily: value }),
          setFontSizePt: (pt) => {
            if (Number.isFinite(pt)) editor.setCharStyle({ fontSizePx: clampSizePx(pt) });
          },
          setColor: (color) => editor.setCharStyle({ color }),
          setHighlight: (color) => editor.dispatch(toggleHighlight(color)),
          clearHighlight: () => editor.setCharStyle({ highlightColor: undefined }),
          clearFormatting: () => {
            editor.dispatch(clearCharFormatting());
            editor.dispatch(applyNamedStyle("Normal"));
          },
        },
      }),
    );
  }

  // Embedder-registered context toolbars (default priority 15, below the built-ins).
  const buildContext = (): ToolbarContext => ({
    format: editor.currentFormat(),
    selection: editor.getSelection(),
    hasRange: isRangeSelection(),
    linkUrl: editor.linkAtCaret(),
    selectionRect: () => editor.getSelectionAnchorRect(),
    objectRect: () => editor.getSelectedObjectRect(),
  });

  // Dispatch a command, then return focus to the editor so typing continues (the
  // context bars act on click, like the ribbon).
  const dispatchFocus = (cmd: Command): void => {
    editor.dispatch(cmd);
    editor.focus();
  };

  // Equation bar (priority 29) — a selected equation block.
  manager.register(
    createEquationContextToolbar({
      anchorRect: () => editor.getSelectedEquationRect(),
      actions: {
        edit: () => editor.editSelectedEquation(), // opens the equation dialog — keep its focus
        alignLeft: () => withSelectedObject((id) => dispatchFocus(setEquationAlignCmd(id, "left"))),
        alignCenter: () => withSelectedObject((id) => dispatchFocus(setEquationAlignCmd(id, "center"))),
        alignRight: () => withSelectedObject((id) => dispatchFocus(setEquationAlignCmd(id, "right"))),
        remove: () => {
          editor.deleteSelectedObject();
          editor.focus();
        },
      },
    }),
  );

  // Review bar (priority 26) — caret inside a comment thread or a suggestion.
  manager.register(
    createReviewContextToolbar({
      review: () => editor.reviewAtCaret(),
      hasRangeSelection: isRangeSelection,
      anchorRect: () => editor.getSelectionAnchorRect(),
      actions: {
        accept: (id) => {
          editor.acceptSuggestion(id);
          editor.focus();
        },
        reject: (id) => {
          editor.rejectSuggestion(id);
          editor.focus();
        },
        open: (id) => editor.revealReview(id), // navigates to the thread — don't steal focus
        resolve: (id, resolved) => {
          editor.resolveThread(id, resolved);
          editor.focus();
        },
      },
    }),
  );

  // Footnote/endnote bar (priority 24) — caret on a note reference marker.
  manager.register(
    createNoteContextToolbar({
      note: () => editor.noteAtCaret(),
      hasRangeSelection: isRangeSelection,
      anchorRect: () => editor.getSelectionAnchorRect(),
      actions: {
        goTo: (info) => {
          editor.goToNote(info.kind, info.id); // jumps into the note body
          editor.focus();
        },
        remove: (info) => dispatchFocus(deleteNoteCmd(info.kind, info.id)),
      },
    }),
  );

  // TOC bar (priority 22) — caret inside a Table-of-Contents entry.
  manager.register(
    createTocContextToolbar({
      inToc: () => editor.caretInToc(),
      hasRangeSelection: isRangeSelection,
      anchorRect: () => editor.getSelectionAnchorRect(),
      update: () => {
        // Full regeneration (Word's F9 "update entire table"): re-lists headings and
        // refreshes entry text, preserving the TOC's look. `recalculateToc` alone only
        // rewrites a literal "\t<number>", which imported entries don't carry (their
        // page number is paint-live) — so it would silently do nothing here.
        editor.dispatch(updateTocFieldCmd());
        editor.focus();
      },
    }),
  );

  // List-item bar (priority 18) — caret in a bulleted/numbered list.
  manager.register(
    createListContextToolbar({
      listKind: () => editor.currentFormat().listKind,
      hasRangeSelection: isRangeSelection,
      anchorRect: () => editor.getSelectionAnchorRect(),
      actions: {
        promote: () => dispatchFocus(changeListLevel(-1)),
        demote: () => dispatchFocus(changeListLevel(1)),
        removeList: () => dispatchFocus(toggleList(editor.currentFormat().listKind === "bullet" ? "bullet" : "decimal")),
      },
    }),
  );

  // Empty-paragraph insert menu (priority 16) — Notion-style ＋ on an empty line.
  // Reusable so both the ＋ chip and the "/" keyboard shortcut (S4) open it.
  const openInsertMenuAt = (clientX: number, clientY: number): void => {
    {
        const insert = (cmd: Command): (() => void) => () => dispatchFocus(cmd);
        const entries: MenuEntry[] = [
          { kind: "item", label: "Heading 1", onClick: insert(applyNamedStyle("Heading1")) },
          { kind: "item", label: "Heading 2", onClick: insert(applyNamedStyle("Heading2")) },
          { kind: "item", label: "Heading 3", onClick: insert(applyNamedStyle("Heading3")) },
          { kind: "sep" },
          { kind: "item", label: "Bulleted list", onClick: insert(toggleList("bullet")) },
          { kind: "item", label: "Numbered list", onClick: insert(toggleList("decimal")) },
          { kind: "sep" },
          { kind: "item", label: "Table", onClick: insert(insertTable(3, 3)) },
          { kind: "item", label: "Page break", onClick: insert(insertPageBreak()) },
          { kind: "item", label: "Table of contents", onClick: insert(insertTocCmd()) },
          { kind: "item", label: "Footnote", onClick: insert(insertFootnoteCmd()) },
          { kind: "sep" },
          // Shapes were reachable only from the ribbon Insert → shapes popover; surface
          // them in the inline ＋ menu too so inline-menu users can discover them (A4).
          {
            kind: "submenu",
            label: "Shape",
            icon: ICONS.shapeRect,
            items: [
              { kind: "item", label: "Rectangle", icon: ICONS.shapeRect, onClick: () => editor.insertShape("rect") },
              { kind: "item", label: "Rounded rectangle", icon: ICONS.shapeRoundRect, onClick: () => editor.insertShape("roundRect") },
              { kind: "item", label: "Ellipse", icon: ICONS.shapeEllipse, onClick: () => editor.insertShape("ellipse") },
              { kind: "item", label: "Triangle", icon: ICONS.shapeTriangle, onClick: () => editor.insertShape("triangle") },
              { kind: "item", label: "Diamond", icon: ICONS.shapeDiamond, onClick: () => editor.insertShape("diamond") },
              { kind: "item", label: "Right arrow", icon: ICONS.shapeRightArrow, onClick: () => editor.insertShape("rightArrow") },
              { kind: "item", label: "Left arrow", icon: ICONS.shapeLeftArrow, onClick: () => editor.insertShape("leftArrow") },
              { kind: "item", label: "Line", icon: ICONS.shapeLine, onClick: () => editor.insertShape("line") },
              { kind: "sep" },
              { kind: "item", label: "Text Box", icon: ICONS.shapeTextBox, onClick: () => editor.insertTextBox() },
            ],
          },
        ];
        showContextMenu(clientX, clientY, entries);
    }
  };
  manager.register(
    createInsertMenuToolbar({
      onEmptyParagraph: () => editor.caretInEmptyParagraph(),
      anchorRect: () => editor.getSelectionAnchorRect(),
      openMenu: (anchorEl) => {
        const r = anchorEl.getBoundingClientRect();
        openInsertMenuAt(r.left, r.bottom + 2);
      },
    }),
  );

  // "/" at the start of an empty paragraph opens the block inserter — the keyboard
  // door into the ＋ Insert menu (S4). Skipped while a form control (filter box,
  // dialog input) has focus so it never hijacks a literal slash typed there.
  app.addEventListener("keydown", (e: KeyboardEvent) => {
    if (e.key !== "/" || e.ctrlKey || e.metaKey || e.altKey) return;
    const ae = document.activeElement;
    if (ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA" || ae.tagName === "SELECT")) return;
    if (!editor.caretInEmptyParagraph()) return;
    e.preventDefault();
    const r = editor.getSelectionAnchorRect();
    if (r) openInsertMenuAt(r.left, r.top + r.height + 4);
  });

  for (const spec of config.contextToolbars) {
    manager.register(createCustomContextToolbar(spec, { context: buildContext, runButton: runCustomButton }));
  }

  refreshContextToolbars = manager.refresh;
  scheduleRefreshContextToolbars = manager.scheduleRefresh;
}

// ---- page setup ------------------------------------------------------------
// Page layout lives in the draggable showPageLayout dialog (ui/pageLayout.ts),
// opened from the Layout-tab button above. The old inline preset panel was retired.

// ---- find & replace bar (Ctrl+F) -------------------------------------------
{
  const bar = document.createElement("div");
  bar.className = "cw-float-panel";
  bar.style.cssText =
    "position:fixed;top:46px;right:24px;display:none;gap:4px;align-items:center;" +
    "background:#fff;border:1px solid #dadce0;border-radius:8px;padding:6px 8px;" +
    "box-shadow:0 2px 8px rgba(0,0,0,.18);z-index:10;font-size:13px;";
  const mkInput = (placeholder: string, width: number): HTMLInputElement => {
    const i = document.createElement("input");
    i.placeholder = placeholder;
    i.setAttribute("aria-label", placeholder); // the field had no accessible name (L6)
    i.style.cssText = `width:${width}px;height:24px;border:1px solid #dadce0;border-radius:4px;padding:0 6px;font-size:13px;`;
    bar.appendChild(i);
    return i;
  };
  const mkBtn = (label: string, title: string, onClick: () => void): HTMLButtonElement => {
    const b = document.createElement("button");
    b.textContent = label;
    b.title = title;
    b.style.cssText = "height:24px;border:none;background:transparent;cursor:pointer;font-size:13px;color:#3c4043;";
    b.addEventListener("click", onClick);
    bar.appendChild(b);
    return b;
  };
  const findInput = mkInput("Find", 140);
  const counter = document.createElement("span");
  counter.style.cssText = "color:#80868b;min-width:40px;text-align:center;";
  counter.textContent = "";
  counter.setAttribute("aria-live", "polite"); // announce the match count to AT (L6)
  counter.setAttribute("aria-label", "Match count");
  bar.appendChild(counter);
  const update = (s: { index: number; total: number }): void => {
    counter.textContent = s.total > 0 ? `${s.index}/${s.total}` : findInput.value ? "No matches" : "";
  };
  // Match-case / whole-word toggles (editor.search already accepts these).
  let matchCase = false;
  let wholeWord = false;
  const reSearch = (): void => {
    update(findInput.value ? editor.search(findInput.value, { matchCase, wholeWord }) : { index: 0, total: 0 });
  };
  const optBtn = (label: string, title: string, get: () => boolean, set: (v: boolean) => void): void => {
    const b = document.createElement("button");
    b.textContent = label;
    b.title = title;
    b.setAttribute("aria-label", title); // a real name, not just the "Aa"/"W" glyph (L6)
    b.style.cssText = "height:24px;min-width:26px;border:1px solid transparent;border-radius:4px;background:transparent;cursor:pointer;font-size:13px;font-weight:600;";
    const paint = (): void => {
      const on = get();
      b.setAttribute("aria-pressed", String(on)); // it's a toggle — expose its state
      b.style.background = on ? "#cfe3fb" : "transparent";
      b.style.borderColor = on ? "#b3d3f5" : "transparent";
      b.style.color = on ? "#0b57d0" : "#3c4043";
    };
    b.addEventListener("click", () => {
      set(!get());
      paint();
      reSearch();
      findInput.focus();
    });
    paint();
    bar.appendChild(b);
  };
  optBtn("Aa", "Match case", () => matchCase, (v) => (matchCase = v));
  optBtn("W", "Match whole word only", () => wholeWord, (v) => (wholeWord = v));
  mkBtn("‹", "Previous (Shift+Enter)", () => update(editor.searchNav(-1)));
  mkBtn("›", "Next (Enter)", () => update(editor.searchNav(1)));
  // Replace is an edit — only in editable mode. Find/navigate stay for the viewer.
  if (!readonly) {
    const replaceInput = mkInput("Replace", 120);
    mkBtn("Replace", "Replace current", () => update(editor.searchReplaceCurrent(replaceInput.value)));
    mkBtn("All", "Replace all", () => {
      const n = editor.searchReplaceAll(replaceInput.value);
      update(editor.search(findInput.value, { matchCase, wholeWord }));
      counter.textContent += ` (${n} replaced)`;
    });
  }
  let findSurface: SurfaceHandle | null = null;
  const close = (): void => {
    bar.style.display = "none";
    editor.searchClear();
    findSurface?.release();
    findSurface = null;
    editor.focus();
  };
  mkBtn("×", "Close (Esc)", close);
  document.body.appendChild(bar);
  detachables.push(bar);

  openFind = (): void => {
    const wasOpen = bar.style.display === "flex";
    bar.style.display = "flex";
    findInput.select();
    findInput.focus();
    reSearch();
    // Register with the surface manager so the bar stacks ABOVE any open dialog
    // (it used to be z-index 10, behind them) and shares the one Escape handler.
    if (!wasOpen) findSurface = openSurface({ el: bar, close, kind: "overlay" });
  };

  findInput.addEventListener("input", reSearch);
  bar.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") {
      update(editor.searchNav(ev.shiftKey ? -1 : 1));
      ev.preventDefault();
    }
    // Escape is handled centrally by the surface manager (closes the topmost).
    ev.stopPropagation(); // typing in the bar never reaches the editor keymap
  });
  window.addEventListener("keydown", (ev) => {
    if ((ev.ctrlKey || ev.metaKey) && ev.key.toLowerCase() === "f") {
      ev.preventDefault();
      openFind();
    }
  }, { signal: teardown.signal });
}

// Dev hook for in-browser verification (break-rule scans, perf probes, and the
// snapshot/replay round-trip: serialize the doc at load, reconstruct from the
// editor's change log, compare). Last-mounted instance wins the global.
const persist = { serializeDocument, deserializeDocument, reconstruct, rehydrateDocMedia, mediaStore };
window.__cw = { doc, tree, engine, editor, createLayoutEngine, sampleDoc, stressDoc, persist, share: goOnlineWithCurrentDoc };

// Hand the wrapper a control surface for the mounted editor.
// Disposer for the WebMCP agent tools (set below when runtime.agentTools is on).
let disposeAgentTools: (() => void) | null = null;

const handle: EditorHandle = {
  getDocument: () => editor.getDocument(),
  createChild: () => editor.createChild(),
  setDocument: (d) => setDocumentFromApi(d),
  openDocx: (file) => openDocxFile(file),
  exportDocx: () => exportBlob("docx"),
  exportPdf: () => exportBlob("pdf"),
  share: () => goOnlineWithCurrentDoc(),
  getDocId: () => collabId,
  getShareLink: () => shareLink,
  getSelection: () => editor.getSelection(),
  insertText: (text) => editor.dispatch(insertTextCmd(text)),
  getMode: () => editor.getMode(),
  setMode: (mode) => editor.setMode(mode),
  getReview: () => editor.getReview(),
  getKnownUsers: () => editor.getKnownUsers(),
  setKnownUsers: (users) => editor.setKnownUsers(users),
  acceptSuggestion: (id) => editor.acceptSuggestion(id),
  rejectSuggestion: (id) => editor.rejectSuggestion(id),
  acceptAllSuggestions: () => editor.acceptAllSuggestions(),
  rejectAllSuggestions: () => editor.rejectAllSuggestions(),
  addComment: (body, mentions) => editor.addComment(body, mentions),
  replyToComment: (threadId, body, mentions) => editor.replyToComment(threadId, body, mentions),
  resolveThread: (threadId, resolved) => editor.resolveThread(threadId, resolved),
  runCommand: (id) => runRegisteredCommand(id),
  setCustomBlockData: (id, data) => editor.dispatch(setCustomBlockDataCmd(id, data)),
  setDecorations: (specs) => editor.setDecorations(specs),
  clearDecorations: () => editor.setDecorations([]),
  invalidateDecorations: () => editor.invalidateDecorations(),
  destroy: () => {
    disposeAgentTools?.(); // unregister WebMCP tools before tearing the editor down
    closeDevPanel(); // remove the floating inspector (body-level) if it's open
    unregisterMediaRetention(); // deliberately no eviction — see registerMediaRetention
    teardown.abort(); // drop every global window/document listener this instance added
    compactRO?.disconnect(); // ResizeObserver isn't signal-aware — stop it explicitly
    for (const el of detachables) el.remove(); // body-level floats (find bar, image bar, …)
    sync?.destroy();
    editor.destroy();
    shell.root.remove();
  },
};
// Context for custom ribbon buttons: the handle + macro helpers (getSelection /
// insertText live on the handle) plus an event emitter and a teardown hook.
ribbonCtx = {
  ...handle,
  emit: (name, payload) => runtime.onEvent?.({ type: "custom", name, payload }),
  registerCleanup: (target) => {
    if (typeof target === "function") teardown.signal.addEventListener("abort", () => target(), { once: true });
    else detachables.push(target);
  },
};

// ===== commands + keymap registry (opt-in via the `commands` option) =========
// Register the embedder's commands, then wire their keybindings to one keydown
// listener on the editor container. The built-in editing keymap (input/keymap.ts)
// is bound on the SAME element earlier (during createEditor), so it fires first
// and calls preventDefault for its chords; we skip anything already handled
// (ev.defaultPrevented) so custom bindings can never clobber Ctrl+B/I/U/Z/Y or
// Ctrl+Enter. Listener is tied to teardown.signal, so destroy() removes it.
if (runtime.commands?.length) {
  for (const cmd of runtime.commands) {
    if (commandRegistry.has(cmd.id)) {
      console.warn(`[canvas-word] commands: duplicate command id "${cmd.id}" ignored`);
      continue;
    }
    commandRegistry.set(cmd.id, cmd);
  }
  // Resolve "Mod" to the concrete platform modifier (⌘ on macOS, Ctrl elsewhere)
  // for both conflict detection and matching.
  const isMac = typeof navigator !== "undefined" && /Mac|iP(hone|ad|od)/.test(navigator.platform || navigator.userAgent || "");
  const { bindings, conflicts, duplicateIds } = resolveCommandBindings(runtime.commands, isMac);
  for (const c of conflicts)
    console.warn(`[canvas-word] commands: keybinding "${c.raw}" for "${c.commandId}" is already bound to "${c.heldBy}" — ignored`);
  for (const id of duplicateIds)
    console.warn(`[canvas-word] commands: duplicate command id "${id}" ignored`);
  if (bindings.length) {
    app.addEventListener(
      "keydown",
      (ev) => {
        if (ev.defaultPrevented) return; // built-in chords win
        for (const b of bindings) {
          if (chordMatches(b.chord, ev, isMac)) {
            ev.preventDefault();
            runRegisteredCommand(b.commandId);
            return;
          }
        }
      },
      { signal: teardown.signal },
    );
  }
}

runtime.onLoadProgress?.({ phase: "ready", percent: 1, loaded: 0, total: 0 });
runtime.onReady?.(handle);
runtime.onEvent?.({ type: "ready" });

// WebMCP agent tooling (opt-in). Lazy-load the polyfill + bridge so non-agent
// embedders never pull them into their bundle; initialize navigator.modelContext,
// then register the tools wrapping the live editor.
if (runtime.agentTools) {
  const cfg = typeof runtime.agentTools === "object" ? runtime.agentTools : {};
  void Promise.all([import("@mcp-b/global"), import("./agent/webmcp")])
    .then(([global, bridge]) => {
      global.initializeWebModelContext();
      disposeAgentTools = bridge.registerAgentTools(
        editor,
        { docId: () => collabId, setDocument: setDocumentFromApi },
        cfg,
      );
    })
    .catch((e) => console.error("[wordcanvas] failed to enable agent tools:", e));
}
}
