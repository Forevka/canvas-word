// Public type surface for the @forevka/wordcanvas (WordCanvas) package.
// Hand-written so the published types stay small, self-contained, and stable
// regardless of internal refactors — no dependency on internal workspace types.

import type { Block, Document, ParaStyle, Run } from "./model";

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
  /** Track first-load progress so you can show a loader while the big chunks
   *  stream. Fires for the editor JS chunk download (`phase: "bundle"`,
   *  indeterminate) and the bundled font fetch (`phase: "fonts"`, the dominant
   *  ~9 MB cost — a smooth, size-weighted bar), then once at `phase: "ready"`
   *  with `percent: 1`. Read `percent` (0..1, monotonic) to drive a progress bar. */
  onLoadProgress?: (progress: LoadProgress) => void;
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
  /** Publish the current document and resolve its shareable link (online only). */
  share(): Promise<string>;
  getDocId(): string | null;
  getShareLink(): string | null;
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
  /** Replace the open document with a programmatically-built one (e.g. a
   *  DocumentBuilder result). The input is cloned. Like openDocx, this starts a
   *  NEW document: undo history and any live collab session are dropped (the
   *  next share() forks). Zoom and scroll are preserved, so calling this on
   *  every data change gives a stable live preview. */
  setDocument(doc: Document): Promise<void>;
  /** Publish the current document and resolve its shareable link (online only). */
  share(): Promise<string>;
  getDocId(): string | null;
  getShareLink(): string | null;
  destroy(): void;
}
