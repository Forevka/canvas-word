// Runtime config bridge between the WordCanvas wrapper and the editor app
// module. WordCanvas builds a runtime and passes it to mountEditorApp(runtime),
// which mounts one editor and calls onReady with a handle. The runtime is a plain
// per-call value (no module singleton), so multiple editors coexist on one page.

import type { DocSelection, Document, Fragment, ReviewLayer, UserInfo } from "@cw/shared";
import type { ChildDocument, EditMode, FieldResolver } from "../index";
import type { ExportWarning } from "../export/exportDocument";
import type { DefaultStyleOverrides, EditorBehavior, EditorTheme, FontsConfig } from "../config";
import type { CustomizeRibbon } from "../ribbon";
import type { LoadProgress } from "./loadProgress";

export type { EditMode, FieldResolver };

/** The file format an export/save produces. */
export type SaveFormat = "docx" | "pdf";

/** Payload handed to a host `onSave` handler when the user triggers an export
 *  while `onSave` is configured (see WordCanvasOptions.onSave). */
export interface SaveEvent {
  /** The exported file as a Blob with the correct MIME type set. */
  blob: Blob;
  /** The same file content as raw bytes (handy for FormData / a Node Buffer). */
  bytes: Uint8Array;
  /** Which format the user exported. */
  format: SaveFormat;
  /** Deduplicated lossy-mapping notes from the export (usually empty). */
  warnings: ExportWarning[];
}

/** Host hook invoked instead of the built-in download when the user exports. */
export type SaveHandler = (event: SaveEvent) => void | Promise<void>;

/** Opt-in WebMCP agent tooling. `true` registers all tool buckets; an object
 *  restricts capabilities or namespaces tool names. */
export interface AgentToolsOptions {
  /** Which tool buckets to register. Default: all three.
   *  read = read & inspect (incl. layout dump); suggest = suggestions + comments;
   *  edit = direct text/format edits. */
  capabilities?: ("read" | "suggest" | "edit")[];
  /** Optional prefix so several editors on one page don't collide (e.g. "doc1"
   *  → "doc1_get_document"). */
  name?: string;
}

/** Initial view-chrome state — what the reader sees on open. Every field is
 *  optional; omit one to keep its built-in default. Lets an embedder fully
 *  configure the surface (panels, rulers, grid, ribbon, zoom) per mount. */
export interface WordCanvasViewOptions {
  // ----- rulers & grid -----
  /** Show the horizontal ruler. Default true (always hidden in readonly). */
  ruler?: boolean;
  /** Show the vertical ruler. Default true (always hidden in readonly). */
  verticalRuler?: boolean;
  /** Show the drawing-grid overlay on every page. Default false. */
  grid?: boolean;
  /** Snap dragged anchored objects to the grid. Default false. */
  snapToGrid?: boolean;
  /** Grid step in document px (96dpi). Default 24 (1/4 inch). */
  gridSpacingPx?: number;
  /** Show non-printing formatting marks (spaces, tabs, paragraph ends, line
   *  breaks) on open. Default false. */
  formattingMarks?: boolean;
  // ----- side panels & drawers -----
  /** Open the Outline / navigation pane (left drawer). Default true. */
  outline?: boolean;
  /** Open the Bookmarks panel. Default false. */
  bookmarks?: boolean;
  /** Open the Review pane (track changes + comments). Default false. */
  reviewPane?: boolean;
  /** Open the Activity panel (online only — who edited & when). Default false. */
  activity?: boolean;
  // ----- chrome -----
  /** Show the ribbon toolbar. Default true (always hidden in readonly). When
   *  false the editor stays fully editable but chromeless — the side drawers it
   *  wires up (outline, bookmarks) still work. */
  toolbar?: boolean;
  /** Start with the ribbon body collapsed (tab strip stays). Default false. */
  ribbonCollapsed?: boolean;
  /** Show the status bar (page/word count + zoom). Default true. */
  statusBar?: boolean;
  /** Show the File ▸ Export **PDF** button. Default true. Set false when the
   *  embedder ships its own export/save pipeline (e.g. via `onSave` or the
   *  `exportPdf()` handle method) and doesn't want the built-in button. */
  exportPdf?: boolean;
  /** Show the File ▸ Export **DOCX** button. Default true. Set false when the
   *  embedder ships its own export/save pipeline (e.g. via `onSave` or the
   *  `exportDocx()` handle method) and doesn't want the built-in button. */
  exportDocx?: boolean;
  /** Initial presentational zoom (1 = 100%, clamped to [0.25, 5]). Default 1. */
  zoom?: number;
}

/** One other collaborator currently in the document. */
export interface Participant {
  siteId: string;
  user?: UserInfo | undefined;
}

/** Events the WordCanvas wrapper re-emits to the embedder. */
export type WordCanvasEvent =
  | { type: "ready" }
  | { type: "shared"; docId: string; url: string }
  | { type: "userEntered"; siteId: string; user?: UserInfo | undefined }
  | { type: "userLeave"; siteId: string; user?: UserInfo | undefined }
  | { type: "presence"; participants: Participant[] }
  | { type: "modeChanged"; mode: EditMode }
  | { type: "reviewChanged"; review: ReviewLayer }
  | { type: "custom"; name: string; payload?: unknown };

/** Handle the editor app exposes back to the WordCanvas wrapper. */
export interface EditorHandle {
  getDocument(): Document;
  /** Create a child document sharing this editor's live styles/fonts/theme (see
   *  ChildDocument) — renders or edits a content slice on canvas. */
  createChild(): ChildDocument;
  /** Replace the open document with a programmatically-built one (e.g. from
   *  the DocumentBuilder). The input is cloned; like openDocx, this is a NEW
   *  document — any live collab session is dropped (the next Share forks).
   *  Discards undo history and unsaved edits; preserves zoom + scroll. */
  setDocument(doc: Document): void;
  /** Open a .docx (auto-publishes when online); resolves when loaded. */
  openDocx(file: File | ArrayBuffer): Promise<void>;
  /** Export the current document to a .docx Blob (track changes baked to the
   *  original baseline, matching the toolbar's Export). For a custom Save button. */
  exportDocx(): Promise<Blob>;
  /** Export the current document to a PDF Blob. */
  exportPdf(): Promise<Blob>;
  /** Publish the current document and resolve its shareable link (online only). */
  share(): Promise<string>;
  getDocId(): string | null;
  getShareLink(): string | null;
  /** The current selection/caret, or null when the editor isn't focused. */
  getSelection(): DocSelection | null;
  /** Insert plain text at the caret, replacing any selection. */
  insertText(text: string): void;
  // ---- review layer (track changes + comments) ----------------------------
  getMode(): EditMode;
  setMode(mode: EditMode): boolean;
  getReview(): ReviewLayer;
  getKnownUsers(): UserInfo[];
  setKnownUsers(users: UserInfo[]): void;
  acceptSuggestion(id: string): void;
  rejectSuggestion(id: string): void;
  acceptAllSuggestions(): void;
  rejectAllSuggestions(): void;
  addComment(body: Fragment, mentions?: UserInfo[]): string | null;
  replyToComment(threadId: string, body: Fragment, mentions?: UserInfo[]): void;
  resolveThread(threadId: string, resolved?: boolean): void;
  destroy(): void;
}

export interface WordCanvasRuntime {
  container: HTMLElement;
  /** Backend base URL. Present ⇒ online (sync, publish, share); absent ⇒ offline. */
  backendUrl?: string | undefined;
  /** Join an existing collaboration session on load (online only). */
  collabId?: string | undefined;
  /** Caller-supplied identity — stamped on changes/presence (attribution). */
  user?: UserInfo | undefined;
  /** Override how a share link is surfaced; default shows a built-in dialog. */
  onShareLink?: ((url: string, docId: string) => void) | undefined;
  /** When set, the toolbar Export buttons hand the produced file to this hook
   *  instead of triggering a browser download — route it to your own pipeline. */
  onSave?: SaveHandler | undefined;
  /** Mount view-only: hide the editing chrome and make every mutation a no-op
   *  (the document is still selectable, copyable, and live for remote edits). */
  readonly?: boolean | undefined;
  /** Initial editor mode ("edit" | "suggest" | "view"). Defaults to "view" when
   *  `readonly`, else "edit". */
  mode?: EditMode | undefined;
  /** Modes the user may switch to (constrains the picker + setMode). */
  allowedModes?: EditMode[] | undefined;
  /** Users that can be @-mentioned in comments (embedder-supplied roster). */
  knownUsers?: UserInfo[] | undefined;
  /** Called once the editor is mounted and ready. */
  onReady?: ((handle: EditorHandle) => void) | undefined;
  /** First-load progress sink (JS-chunk download + font fetch). Lets the embedder
   *  drive a loading bar; fires until `phase: "ready"` at `percent: 1`. */
  onLoadProgress?: ((p: LoadProgress) => void) | undefined;
  /** Sink for collaboration events (presence, share, ready) → WordCanvas.on(...). */
  onEvent?: ((ev: WordCanvasEvent) => void) | undefined;
  /** Resolve a custom field's content from the host backend. When set, right-
   *  clicking a custom field offers "Update Field (<name>)", which calls this and
   *  splices the returned OOXML in as the field's new result. */
  resolveField?: FieldResolver | undefined;
  /** Expose the editor to AI agents over WebMCP (navigator.modelContext). `true`
   *  registers all tool buckets; an object restricts/namespaces them. The polyfill
   *  is lazy-loaded only when this is set, so other embedders pay nothing. */
  agentTools?: boolean | AgentToolsOptions | undefined;
  /** Initial view-chrome state (rulers, grid, snap). Omit ⇒ built-in defaults. */
  view?: WordCanvasViewOptions | undefined;
  /** Color theme override (per-instance). Omit ⇒ the built-in look. */
  theme?: EditorTheme | undefined;
  /** Override the LIBRARY's built-in default run/paragraph styles for NEW/blank
   *  documents + the fallback stylesheet. A loaded .docx keeps its own defaults. */
  overrideDefaultStyles?: DefaultStyleOverrides | undefined;
  /** Behavior tuning (zoom step/clamp, indent step, default grid spacing). */
  behavior?: EditorBehavior | undefined;
  /** Custom fonts (URL-loaded faces + sizing) and toolbar disables. */
  fonts?: FontsConfig | undefined;
  /** Customize the ribbon: reorder/remove built-ins by id and add custom tabs,
   *  groups, and buttons. Called once at mount with a mutation API. */
  customizeRibbon?: CustomizeRibbon | undefined;
}
