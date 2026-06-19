// Public type surface for the @forevka/wordcanvas (WordCanvas) package.
// Hand-written so the published types stay small, self-contained, and stable
// regardless of internal refactors — no dependency on internal workspace types.

import type { Block, CharStyle, DocPosition, Document, ParaStyle, Run } from "./model";

export type { Document } from "./model";

/** Caller-supplied identity, used for change attribution and presence. */
export interface UserInfo {
  id: string;
  firstName: string;
  lastName: string;
}

/** One other collaborator currently in the document (online mode). */
export interface Participant {
  siteId: string;
  user?: UserInfo;
}

/** Editor mode: "edit" (raw editing), "suggest" (edits become tracked changes),
 *  or "view" (read-only). The three are mutually exclusive. */
export type EditMode = "edit" | "suggest" | "view";

// ===== Review layer (track changes + comments) =============================
// An attributed, anchored overlay on top of the core document. `getReview()`
// returns a snapshot; the `reviewChanged` event carries the full layer.

/** Rich-text comment body — the same style-homogeneous run list the document
 *  uses (so a comment composes/renders like any other content). */
export type Fragment = Run[];

/** A range into the live document; `start === end` is a point (caret) anchor. */
export interface ReviewAnchor {
  start: DocPosition;
  end: DocPosition;
}

export type SuggestionKind = "insert" | "delete" | "format" | "structural";

/** A tracked structural edit (split/merge/block/table op). `op`/`inverse` are the
 *  internal core ops captured at intercept time — opaque on the public surface
 *  (not part of the stable contract); `blockId` is the block the record hangs on. */
export interface StructuralChange {
  op: unknown;
  inverse: unknown;
  blockId: string;
}

/** One tracked change. Text-bearing kinds (insert/delete/format) cover live text;
 *  `structural` records carry a degenerate point anchor + the applied op. */
export interface Suggestion {
  /** Globally unique id. */
  id: string;
  kind: SuggestionKind;
  /** Covers text present in the live document. */
  anchor: ReviewAnchor;
  /** Attribution (the embedder owns identity; the layer just presents it). */
  author: UserInfo;
  /** Wall-clock ms; display/sort only. */
  createdAt: number;
  /** format only: the applied char-style patch + its inverse (for reject). */
  patch?: Partial<CharStyle>;
  inverse?: Partial<CharStyle>;
  /** structural only: the applied core op, its inverse, and the host block. */
  structural?: StructuralChange;
  /** Optional grouping so related records accept/reject as a unit. */
  groupId?: string;
}

export interface Comment {
  id: string;
  author: UserInfo;
  body: Fragment;
  createdAt: number;
  editedAt?: number;
  /** Resolved identities of users @-mentioned in this comment. */
  mentions?: UserInfo[];
}

export type ThreadStatus = "open" | "resolved";

export interface CommentThread {
  id: string;
  /** Highlighted span; `start === end` ⇒ point comment. */
  anchor: ReviewAnchor;
  status: ThreadStatus;
  /** Ordered; `comments[0]` is the root, the rest are replies. */
  comments: Comment[];
}

/** Snapshot of the review overlay (suggestions + comment threads). */
export interface ReviewLayer {
  docId: string;
  /** Core document version these anchors were last rebased to. */
  baseVersion: number;
  suggestions: Suggestion[];
  threads: CommentThread[];
}

// ===== Theming & behavior (see WordCanvasOptions.theme / .behavior) =========

/** Ruler band styling (the strip the horizontal + vertical rulers paint). */
export interface RulerTheme {
  /** Band fill behind the ticks. */
  bg?: string;
  /** The lighter "content area" portion of the band (within the margins). */
  content?: string;
  /** Tick + boundary line color. */
  line?: string;
  /** Number label color. */
  label?: string;
  /** CSS font for the number labels. */
  font?: string;
}

/** Editor color theme — every field optional; omit to keep the built-in value.
 *  Affects the on-screen editor only (exported PDFs keep the built-in look). */
export interface EditorTheme {
  /** Gray gutter behind the pages (the scroll area + ruler troughs). */
  canvasBackground?: string;
  /** Native table gridlines when a cell carries no explicit border. */
  grid?: string;
  /** Drawing-grid overlay mesh (the dotted snap grid). */
  gridMesh?: string;
  /** External hyperlink color. */
  externalLink?: string;
  /** UI accent (band-edit boundary and similar chrome affordances). */
  accent?: string;
  /** Footnote separator rule. */
  footnoteRule?: string;
  /** Non-printing formatting marks (space dots, tab arrows, pilcrows). */
  formattingMark?: string;
  /** TOC dot-leader color. */
  tocLeader?: string;
  /** Placeholder fill for an image whose bitmap hasn't loaded. */
  imagePlaceholder?: string;
  /** Rule between newspaper columns. */
  columnSeparator?: string;
  /** Blinking text caret color. */
  caret?: string;
  /** Find-highlight color. */
  searchHighlight?: string;
  /** Comment pin fill once its thread is resolved. */
  reviewPinResolved?: string;
  /** Comment pin stroke. */
  reviewPinStroke?: string;
  /** Gap (px) between stacked pages. */
  pageGapPx?: number;
  /** Ruler band styling. */
  ruler?: RulerTheme;
}

/** Editor behavior tuning — every field optional. */
export interface EditorBehavior {
  /** Multiplicative zoom step for +/- and Ctrl+wheel. Default 1.1. */
  zoomStep?: number;
  /** Absolute minimum zoom. Default 0.25. */
  zoomMin?: number;
  /** Absolute maximum zoom. Default 5. */
  zoomMax?: number;
  /** Indent / outdent step in px per toolbar press. Default 36. */
  indentStepPx?: number;
  /** Default drawing-grid spacing in px (view.gridSpacingPx still overrides). Default 24. */
  gridSpacingPx?: number;
}

/** Overrides the LIBRARY's built-in default run/paragraph styles for NEW/blank
 *  documents and the fallback stylesheet — NOT a loaded .docx's own defaults. */
export interface DefaultStyleOverrides {
  /** Body (Normal) font family, e.g. "Times New Roman, serif". */
  fontFamily?: string;
  /** Body font size in px (96 dpi). */
  fontSizePx?: number;
  /** Body text color (any CSS color string). */
  color?: string;
  /** Default paragraph line height (multiplier). */
  lineHeight?: number;
  /** Heading font family (Title/Heading1/Heading2). Omit to keep the built-in. */
  headingFontFamily?: string;
}

/** A ready-made dark-canvas theme (darkens the chrome/gutter; pages stay white
 *  for print fidelity). Spread + tweak it for your own theme. */
export declare const darkCanvasTheme: EditorTheme;

// ===========================================================================

export interface WordCanvasOptions {
  /** Element to mount the editor into. */
  container: HTMLElement;
  /** Backend base URL (e.g. "https://api.example.com"). Online iff provided;
   *  omit for a fully offline editor (no sync, publish, or share). */
  backendUrl?: string;
  /** Open this document on load (online only) — e.g. the id returned by an
   *  upload. Canonical name; `collabId` is the deprecated alias. */
  docId?: string;
  /** @deprecated Use `docId`. Join an existing collaboration session on load. */
  collabId?: string;
  /** Caller-supplied identity (attribution + presence). The embedder owns auth. */
  user?: UserInfo;
  /** Override how a share link is surfaced (default: a built-in dialog). */
  onShareLink?: (url: string, docId: string) => void;
  /** Route exports to your own pipeline. When set, the toolbar's Export (PDF /
   *  DOCX) buttons hand the produced file to this callback — a Blob plus raw
   *  bytes, the format, and any export warnings (see SaveEvent) — instead of
   *  triggering a browser download. Return a promise to keep the UI responsive
   *  while you upload. Omit to keep the default download behaviour. (For a fully
   *  custom button, call `exportDocx()` / `exportPdf()` on the instance.) */
  onSave?: SaveHandler;
  /** Mount as a view-only viewer: the document renders and stays selectable and
   *  copyable, but the editing chrome is hidden and every mutation is a no-op.
   *  In an online session a read-only client still receives live remote edits.
   *  Equivalent to `mode: "view"`. */
  readonly?: boolean;
  /** Initial editor mode: "edit" (default), "suggest" (edits become tracked
   *  changes), or "view" (read-only). Defaults to "view" when `readonly`. */
  mode?: EditMode;
  /** Restrict which modes the user can switch to (constrains the mode picker and
   *  setMode). Omit ⇒ all three. e.g. ["suggest","view"] locks out raw editing. */
  allowedModes?: EditMode[];
  /** Users that can be @-mentioned in comments. The embedder owns this roster;
   *  update it later with `setKnownUsers`. */
  knownUsers?: UserInfo[];
  /** Resolve a custom (developer-defined) field's content from your backend. When
   *  set, right-clicking a custom field offers "Update Field (<name>)"; the
   *  callback receives the field's name + verbatim instruction and returns the new
   *  result to splice in (see FieldResult). */
  resolveField?: FieldResolver;
  /** Expose this editor to AI agents over WebMCP (https://webmcp.dev), the standard
   *  `navigator.modelContext` API. `true` registers the full tool set — read &
   *  inspect (including a layout-geometry dump for debugging rendering / text-
   *  placement issues), plus suggestions, comments, and direct edits. Pass an
   *  object to restrict capabilities or namespace the tool names. The WebMCP
   *  polyfill is lazy-loaded only when this is set, so it adds nothing for
   *  embedders that don't opt in. Connect an agent via the WebMCP browser
   *  extension / Chrome DevTools MCP. */
  agentTools?: boolean | AgentToolsOptions;
  /** Initial view-chrome state — what the reader sees on open (rulers, grid,
   *  panels, ribbon, status bar, the Export buttons, zoom). Omit any field to keep
   *  its default. See WordCanvasViewOptions. */
  view?: WordCanvasViewOptions;
  /** Color theme — pass a (partial) `EditorTheme` to recolor the editor chrome:
   *  the gray canvas/gutter, gridlines, hyperlink/accent colors, formatting marks,
   *  TOC leaders, rulers, caret, etc. Omit any field to keep its built-in value.
   *  A ready-made `darkCanvasTheme` is exported to spread + tweak. Affects the
   *  on-screen editor only (exported PDFs keep the built-in look). */
  theme?: EditorTheme;
  /** Override the LIBRARY's built-in default run/paragraph styles (body font,
   *  size, color, line height, heading font) for NEW/blank documents and the
   *  fallback stylesheet. IMPORTANT: a loaded .docx keeps its OWN defaults
   *  (w:docDefaults / Normal) — set defaults there, or here, not both. */
  overrideDefaultStyles?: DefaultStyleOverrides;
  /** Behavior tuning: zoom step (default 1.1), zoom clamp (0.25–5), indent step
   *  (36px), and default drawing-grid spacing (24px; `view.gridSpacingPx` wins). */
  behavior?: EditorBehavior;
  /** Track first-load progress so you can show a loader while the big chunks
   *  stream. Fires for the editor JS chunk download (`phase: "bundle"`,
   *  indeterminate) and the bundled font fetch (`phase: "fonts"`, the dominant
   *  ~9 MB cost — a smooth, size-weighted bar), then once at `phase: "ready"`
   *  with `percent: 1`. Read `percent` (0..1, monotonic) to drive a progress bar. */
  onLoadProgress?: (progress: LoadProgress) => void;
}

/** Initial view-chrome state (see WordCanvasOptions.view). Every field is
 *  optional; omit one to keep its built-in default. */
export interface WordCanvasViewOptions {
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
  /** Open the Outline / navigation pane (left drawer). Default true. */
  outline?: boolean;
  /** Open the Bookmarks panel. Default false. */
  bookmarks?: boolean;
  /** Open the Review pane (track changes + comments). Default false. */
  reviewPane?: boolean;
  /** Open the Activity panel (online only — who edited & when). Default false. */
  activity?: boolean;
  /** Show the ribbon toolbar. Default true (always hidden in readonly). When
   *  false the editor stays fully editable but chromeless. */
  toolbar?: boolean;
  /** Start with the ribbon body collapsed (tab strip stays). Default false. */
  ribbonCollapsed?: boolean;
  /** Show the status bar (page/word count + zoom). Default true. */
  statusBar?: boolean;
  /** Show the File ▸ Export **PDF** button. Default true. Set false when the
   *  embedder ships its own export/save pipeline (e.g. via `onSave` or the
   *  `exportPdf()` method) and doesn't want the built-in button. */
  exportPdf?: boolean;
  /** Show the File ▸ Export **DOCX** button. Default true. Set false when the
   *  embedder ships its own export/save pipeline (e.g. via `onSave` or the
   *  `exportDocx()` method) and doesn't want the built-in button. */
  exportDocx?: boolean;
  /** Initial presentational zoom (1 = 100%, clamped to [0.25, 5]). Default 1. */
  zoom?: number;
}

/** First-load progress (see WordCanvasOptions.onLoadProgress). */
export interface LoadProgress {
  /** Coarse stage: "bundle" = downloading the editor JS chunk (indeterminate —
   *  a dynamic import can't be byte-measured); "fonts" = fetching the bundled
   *  fonts (measurable); "ready" = mount complete. */
  phase: "bundle" | "fonts" | "ready";
  /** Overall startup progress, 0..1, monotonic. Drive a progress bar with this. */
  percent: number;
  /** Bytes fetched / total for the current measurable phase (fonts). Both 0 when
   *  the phase is indeterminate (bundle). Informational. */
  loaded: number;
  total: number;
}

/** Opt-in WebMCP agent tooling (see WordCanvasOptions.agentTools). */
export interface AgentToolsOptions {
  /** Which tool buckets to register. Default: all three.
   *  - `read`  — read & inspect (get_document, get_selection, search_document,
   *    inspect_layout, get_document_stats);
   *  - `suggest` — set_mode, comments, tracked-change suggestions, accept/reject;
   *  - `edit`  — replace_text, insert_text, format_text, set_alignment,
   *    select_range, undo/redo, set_document. */
  capabilities?: ("read" | "suggest" | "edit")[];
  /** Optional prefix so several editors on one page don't collide
   *  (e.g. "doc1" → "doc1_get_document"). */
  name?: string;
}

/** Request passed to a custom-field resolver (see WordCanvasOptions.resolveField). */
export interface FieldResolveRequest {
  /** The field's id within the document. */
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
 *    styles come through (media survives export);
 *  - or an **OOXML fragment string** (w:p / w:tbl, or a w:document) for simple,
 *    text-only results (no embedded media). */
export type FieldResult = string | ArrayBuffer | Uint8Array | Blob;

/** Host hook producing a field's result for a resolve request. */
export type FieldResolver = (req: FieldResolveRequest) => Promise<FieldResult>;

/** The file format an export/save produces. */
export type SaveFormat = "docx" | "pdf";

/** A deduplicated lossy-mapping note from an export. `count` = times it fired. */
export interface ExportWarning {
  code: string;
  detail?: string;
  count: number;
}

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

/** What to render into a child-document surface (see WordCanvas.createChild). */
export type ChildContent =
  | { kind: "blocks"; blocks: Block[] }
  | { kind: "fragment"; fragment: { blocks: { runs: Run[]; style: ParaStyle }[]; inline?: boolean } }
  | { kind: "runs"; runs: Run[]; paraStyle?: ParaStyle }
  /** Body OOXML (w:p / w:tbl, or a w:document). Media-free (no images by r:id). */
  | { kind: "ooxml"; ooxml: string }
  /** A real sample of a named style ("AaBbCc" in the style's resolved formatting). */
  | { kind: "styleSample"; styleId: string; sampleText?: string };

export interface ChildRenderOptions {
  /** Layout/wrap width in CSS px. Default: target.clientWidth. */
  widthPx?: number;
  /** Child section margins (all sides), CSS px. Default 0. */
  padding?: number;
  /** Cap the displayed height; content beyond it is clipped. */
  maxHeightPx?: number;
  /** Fixed presentational scale (1 = 100%). Ignored when `fit` is set. */
  zoom?: number;
  /** Auto-scale the content to fit `widthPx` × `maxHeightPx` (compact swatches). */
  fit?: boolean;
}

/** A live interactive editor over a child slice (canvas-native editing). */
export interface ChildEditorHandle {
  /** The edited content as model blocks (hand back to splice into the parent). */
  getBlocks(): Block[];
  getDocument(): Document;
  /** Insert an image at the caret. The bytes are registered in the parent
   *  document's shared media store (content-addressed), so the image renders here,
   *  survives serialize/export, and resolves after commit-back. Dimensions default
   *  to the image's natural size. */
  insertImage(bytes: Uint8Array, mime: string, opts?: { widthPx?: number; heightPx?: number }): Promise<void>;
  destroy(): void;
}

/** A lightweight sibling document that shares the parent editor's live styles,
 *  fonts and theme and renders/edits a content slice on canvas. */
export interface ChildDocument {
  /** Paint a static preview of `content` into `target`. Safe to call repeatedly. */
  render(target: HTMLElement, content: ChildContent, opts?: ChildRenderOptions): void;
  /** Mount an interactive editor over `content` inside `target`. */
  mountEditor(target: HTMLElement, content: ChildContent, opts?: ChildRenderOptions): ChildEditorHandle;
  /** Tear down every surface/editor this child created. */
  destroy(): void;
}

/** Event name → payload, for `on(...)`/`off(...)`. */
export interface WordCanvasEventMap {
  ready: Record<string, never>;
  shared: { docId: string; url: string };
  userEntered: { siteId: string; user?: UserInfo };
  userLeave: { siteId: string; user?: UserInfo };
  presence: { participants: Participant[] };
  /** The editor mode changed (picker or setMode). */
  modeChanged: { mode: EditMode };
  /** The review overlay changed (suggestion/comment added, resolved, rebased).
   *  Carries the full layer so panels can re-render. */
  reviewChanged: { review: ReviewLayer };
}

/** Handle resolved once the editor is mounted (via `whenReady()`). */
export interface EditorHandle {
  /** Snapshot of the current document (plain data). */
  getDocument(): Document;
  /** Create a child document sharing this editor's live styles/fonts/theme. */
  createChild(): ChildDocument;
  /** Replace the open document with a programmatically-built one (e.g. a
   *  DocumentBuilder result). Cloned; drops undo history and any live collab
   *  session; preserves zoom + scroll. */
  setDocument(doc: Document): void;
  /** Open a .docx (auto-publishes when online); resolves when loaded. */
  openDocx(file: File | ArrayBuffer): Promise<void>;
  /** Export the current document to a .docx Blob (track changes baked to the
   *  original baseline, matching the toolbar's Export). For a custom Save button. */
  exportDocx(): Promise<Blob>;
  /** Export the current document to a PDF Blob (see `exportDocx`). */
  exportPdf(): Promise<Blob>;
  /** Publish the current document and resolve its shareable link (online only). */
  share(): Promise<string>;
  getDocId(): string | null;
  getShareLink(): string | null;
  // ---- review layer (track changes + comments) ----------------------------
  /** The current editor mode. */
  getMode(): EditMode;
  /** Switch mode; returns false if the mode isn't allowed. */
  setMode(mode: EditMode): boolean;
  /** Snapshot of the review overlay (suggestions + comment threads). */
  getReview(): ReviewLayer;
  /** The effective @-mentionable roster (configured base + live editors). */
  getKnownUsers(): UserInfo[];
  /** Replace the configured base roster (live editors stay merged on top). */
  setKnownUsers(users: UserInfo[]): void;
  acceptSuggestion(id: string): void;
  rejectSuggestion(id: string): void;
  acceptAllSuggestions(): void;
  rejectAllSuggestions(): void;
  /** Add a comment thread anchored to the current selection. Returns the thread
   *  id, or null if there's no selection. */
  addComment(body: Fragment, mentions?: UserInfo[]): string | null;
  replyToComment(threadId: string, body: Fragment, mentions?: UserInfo[]): void;
  resolveThread(threadId: string, resolved?: boolean): void;
  destroy(): void;
}

export declare class WordCanvas {
  constructor(opts: WordCanvasOptions);
  /** Subscribe to a collaboration event. Returns an unsubscribe function. */
  on<E extends keyof WordCanvasEventMap>(event: E, handler: (data: WordCanvasEventMap[E]) => void): () => void;
  /** Unsubscribe a handler previously passed to `on()`. */
  off<E extends keyof WordCanvasEventMap>(event: E, handler: (data: WordCanvasEventMap[E]) => void): void;
  /** Resolves once the editor is mounted and ready. */
  whenReady(): Promise<EditorHandle>;
  /** Create a child document that shares this editor's live styles, fonts and
   *  theme — render a real, document-styled preview of a content slice onto a
   *  canvas (instead of an HTML approximation), or mount a canvas-native editor
   *  over a slice. Returned synchronously; render() before ready is buffered.
   *  (mountEditor() requires the editor to be ready — await whenReady() first.) */
  createChild(): ChildDocument;
  /** Open a .docx. When online, auto-publishes it and surfaces a share link. */
  openDocx(file: File | ArrayBuffer): Promise<void>;
  /** Export the current document to a .docx Blob — track changes baked to the
   *  original baseline, exactly like the toolbar's Export. Wire it to your own
   *  Save button and POST it anywhere. Resolves once the editor is ready. */
  exportDocx(): Promise<Blob>;
  /** Export the current document to a PDF Blob (see `exportDocx`). */
  exportPdf(): Promise<Blob>;
  /** Replace the open document with a programmatically-built one (e.g. a
   *  DocumentBuilder result). The input is cloned. Like openDocx, this starts a
   *  NEW document: undo history and any live collab session are dropped (the
   *  next share() forks). Zoom and scroll are preserved, so calling this on
   *  every data change gives a stable live preview. */
  setDocument(doc: Document): Promise<void>;
  /** Publish the current document and resolve its shareable link (online only). */
  share(): Promise<string>;
  // ---- review layer (track changes + comments) ----------------------------
  /** Current editor mode. Returns null before the editor is ready. */
  getMode(): EditMode | null;
  /** Switch mode. Resolves to false if the mode isn't allowed (or not ready). */
  setMode(mode: EditMode): Promise<boolean>;
  /** Snapshot of the review overlay (suggestions + comment threads). */
  getReview(): Promise<ReviewLayer>;
  /** The effective @-mentionable roster (configured base + live editors). */
  getKnownUsers(): Promise<UserInfo[]>;
  /** Replace the configured base roster (live editors stay merged on top). */
  setKnownUsers(users: UserInfo[]): Promise<void>;
  acceptSuggestion(id: string): Promise<void>;
  rejectSuggestion(id: string): Promise<void>;
  acceptAllSuggestions(): Promise<void>;
  rejectAllSuggestions(): Promise<void>;
  /** Add a comment thread anchored to the current selection. Resolves to the
   *  thread id, or null if there's no selection. */
  addComment(body: Fragment, mentions?: UserInfo[]): Promise<string | null>;
  replyToComment(threadId: string, body: Fragment, mentions?: UserInfo[]): Promise<void>;
  resolveThread(threadId: string, resolved?: boolean): Promise<void>;
  getDocId(): string | null;
  getShareLink(): string | null;
  destroy(): void;
}
