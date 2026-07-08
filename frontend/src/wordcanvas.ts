// WordCanvas — the embeddable package entry.
//
//   import { WordCanvas } from "@forevka/wordcanvas";
//   const ed = new WordCanvas({ container, backendUrl: "https://…", user });
//   ed.on("userEntered", ({ user }) => showAvatar(user));
//
// `backendUrl` is the online/offline kill-switch: when set, opening a document
// auto-publishes it and exposes a shareable link, edits sync live, and presence
// events fire; when omitted, the editor runs fully offline. Multiple WordCanvas
// instances can coexist on one page (class-scoped chrome, per-instance runtime).

import type { AgentToolsOptions, EditMode, EditorHandle, FieldResolver, Participant, SaveEvent, SaveFormat, SaveHandler, WordCanvasEvent, WordCanvasRuntime, WordCanvasViewOptions } from "./app/runtime";
import type { ChildContent, ChildDocument, ChildEditorHandle, ChildRenderOptions, FieldResolveRequest, FieldResult } from "./index";
import type { Document, Fragment, ReviewLayer, UserInfo } from "@cw/shared";
import type { CjkConfig, CustomFontDef, CustomFontFaces, DefaultStyleOverrides, EditorBehavior, EditorTheme, FontsConfig } from "./config";
import { darkCanvasTheme } from "./config";
import type { CustomizeRibbon, RibbonActionContext, RibbonApi, RibbonButtonSpec } from "./ribbon";
import { makeFloatingDialog } from "./ui/floatingDialog";
import type { DocSelection } from "@cw/shared";
import { BUNDLE_SHARE, type LoadProgress } from "./app/loadProgress";

export type { Document, UserInfo, Participant, EditMode, ReviewLayer, Fragment, FieldResolver, FieldResolveRequest, FieldResult, AgentToolsOptions, LoadProgress, WordCanvasViewOptions, SaveEvent, SaveFormat, SaveHandler };
export type { ChildDocument, ChildContent, ChildRenderOptions, ChildEditorHandle };
export type { EditorTheme, DefaultStyleOverrides, EditorBehavior, FontsConfig, CjkConfig, CustomFontDef, CustomFontFaces };
export type { CustomizeRibbon, RibbonApi, RibbonButtonSpec, RibbonActionContext, DocSelection };
export { darkCanvasTheme, makeFloatingDialog };

export interface WordCanvasOptions {
  /** Element to mount the editor into. */
  container: HTMLElement;
  /** Backend base URL (e.g. "https://api.example.com"). Online iff provided. */
  backendUrl?: string;
  /** Open this document on load (online only) — e.g. the id returned by an
   *  upload. The canonical name; `collabId` is the deprecated alias. */
  docId?: string;
  /** @deprecated Use `docId`. Join an existing collaboration session on load. */
  collabId?: string;
  /** Caller-supplied identity (attribution + presence). The embedder owns auth. */
  user?: UserInfo;
  /** Override how a share link is surfaced (default: a built-in dialog). */
  onShareLink?: (url: string, docId: string) => void;
  /** Route exports to your own pipeline. When set, the toolbar's Export (PDF /
   *  DOCX) buttons hand the produced file to this callback — a Blob plus raw
   *  bytes, the format, and any export warnings — instead of triggering a browser
   *  download. Return a promise to keep the UI responsive while you upload. Omit
   *  to keep the default download behaviour. (For a fully custom button, call
   *  `exportDocx()` / `exportPdf()` on the instance or its EditorHandle.) */
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
   *  provided, right-clicking a custom field in the document offers "Update Field
   *  (<name>)"; this callback receives the field's name + verbatim instruction and
   *  must return OOXML (w:p / w:tbl, or a full w:document) for its new result. */
  resolveField?: FieldResolver;
  /** Expose this editor to AI agents over [WebMCP](https://webmcp.dev)
   *  (the standard `navigator.modelContext` API). `true` registers the full tool
   *  set (read & inspect — including a layout-geometry dump for debugging render
   *  issues — plus suggestions, comments, and direct edits); pass an object to
   *  restrict capabilities or namespace tool names. The WebMCP polyfill is
   *  lazy-loaded only when this is set, so it adds nothing for embedders that
   *  don't opt in. Connect an agent via the WebMCP browser extension / Chrome
   *  DevTools MCP. */
  agentTools?: boolean | AgentToolsOptions;
  /** Initial view-chrome state — full control over what the reader sees on open:
   *  the horizontal & vertical rulers, drawing grid + snap + step, the Outline /
   *  Bookmarks / Review / Activity panels, the ribbon toolbar (and whether it
   *  starts collapsed), the status bar, and the initial zoom. Omit any field to
   *  keep its default. */
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
  /** Develop mode (default false). When true, the ribbon gains a dedicated
   *  **Developer** tab whose "Inspect Document Tree" button opens a floating
   *  devtools-style panel: it renders the parsed `Document` model as a navigable
   *  tree, highlights a node's painted region on hover, reveals the matching tree
   *  node when you hover the canvas, and shows each node's properties + raw JSON.
   *  Purely a debugging aid — it adds no chrome and runs nothing until the
   *  developer opens it from that tab. Leave off for production embeds. */
  develop?: boolean;
  /** Show the "Organize Pages" button (Layout tab) that opens the visual
   *  page-reorder overlay — drag page thumbnails to re-sequence sections. It
   *  reorders whole page-break/section-delimited groups, so content is never
   *  split. Default true; set false to hide it. */
  organizePages?: boolean;
  /** Supply your OWN fonts, loaded from URLs at runtime. Each custom font needs a
   *  `family` (stored in the model + shown in the toolbar), per-style face URLs
   *  (`regular` required; bold/italic/boldItalic optional, falling back to regular),
   *  and REQUIRED `sizing` ({ ascent, descent } as fractions of em) so the editor
   *  and the PDF/DOCX exporters paginate identically. `disableBuiltin` hides
   *  built-in families (by their original name, e.g. "Calibri") from the toolbar
   *  only — they stay loaded so a loaded .docx that uses them still renders.
   *  TTF/OTF only (WOFF2 is not supported). Font hosts must allow CORS. */
  fonts?: FontsConfig;
  /** CJK (Chinese/Japanese/Korean) tuning. Line-breaking and kinsoku work without
   *  config; set `cjk.locale` ("ja"/"ko"/"zh") to tune locale-specific breaking,
   *  and `cjk.fallbackFont` to a registered `fonts` family so CJK glyphs measure,
   *  export, and embed with a known font instead of an arbitrary system one. */
  cjk?: CjkConfig;
  /** Customize the ribbon toolbar. Called once at mount with a `RibbonApi` to
   *  reorder/remove built-in tabs/groups/buttons (by id — discover them with
   *  `api.tabs()/groups()/items()`) and add your own tabs, groups, and buttons.
   *  A custom button's `onClick` receives a `RibbonActionContext` (the editor
   *  handle + `getSelection`/`insertText`/`emit`/`registerCleanup`) — for macros,
   *  config popups, and informational popups. */
  customizeRibbon?: CustomizeRibbon;
  /** Track first-load progress so you can show a loader while the big chunks
   *  stream. Fires for the editor JS chunk download (`phase: "bundle"`,
   *  indeterminate) and the bundled font fetch (`phase: "fonts"`, the dominant
   *  ~9 MB cost — a smooth, size-weighted bar), then once at `phase: "ready"`
   *  with `percent: 1`. Read `percent` (0..1, monotonic) to drive a progress bar. */
  onLoadProgress?: (progress: LoadProgress) => void;
}

/** Event name → payload, for `on(...)`. */
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
  /** A custom ribbon button emitted an event via `ctx.emit(name, payload)`. */
  custom: { name: string; payload?: unknown };
}

type Handler<E extends keyof WordCanvasEventMap> = (data: WordCanvasEventMap[E]) => void;

export class WordCanvas {
  private handle: EditorHandle | null = null;
  private readonly ready: Promise<EditorHandle>;
  private readonly handlers = new Map<string, Set<(data: unknown) => void>>();

  constructor(opts: WordCanvasOptions) {
    this.ready = new Promise<EditorHandle>((resolve) => {
      const runtime: WordCanvasRuntime = {
        container: opts.container,
        backendUrl: opts.backendUrl,
        // `docId` is the canonical name; fall back to the deprecated `collabId`.
        collabId: opts.docId ?? opts.collabId,
        user: opts.user,
        onShareLink: opts.onShareLink,
        onSave: opts.onSave,
        readonly: opts.readonly,
        mode: opts.mode,
        allowedModes: opts.allowedModes,
        knownUsers: opts.knownUsers,
        resolveField: opts.resolveField,
        agentTools: opts.agentTools,
        view: opts.view,
        theme: opts.theme,
        overrideDefaultStyles: opts.overrideDefaultStyles,
        behavior: opts.behavior,
        develop: opts.develop,
        organizePages: opts.organizePages,
        fonts: opts.fonts,
        cjk: opts.cjk,
        customizeRibbon: opts.customizeRibbon,
        onLoadProgress: opts.onLoadProgress,
        onReady: (h) => {
          this.handle = h;
          resolve(h);
        },
        onEvent: (ev) => this.emit(ev),
      };
      // Lazy-load the editor chunk, then mount this instance. Each call mounts an
      // independent editor (the chunk's module-eval cost is shared, the mount is
      // per-instance), so multiple WordCanvas instances coexist on one page.
      // Bracket the chunk download for the loading bar: a dynamic import() can't
      // be byte-measured, so report 0 here and snap to BUNDLE_SHARE on resolve;
      // mountEditorApp then drives the measurable fonts → ready progress.
      opts.onLoadProgress?.({ phase: "bundle", percent: 0, loaded: 0, total: 0 });
      void import("./editorApp").then((m) => {
        opts.onLoadProgress?.({ phase: "bundle", percent: BUNDLE_SHARE, loaded: 0, total: 0 });
        return m.mountEditorApp(runtime);
      });
    });
  }

  /** Subscribe to a collaboration event. Returns an unsubscribe function. */
  on<E extends keyof WordCanvasEventMap>(event: E, handler: Handler<E>): () => void {
    let set = this.handlers.get(event);
    if (!set) this.handlers.set(event, (set = new Set()));
    set.add(handler as (data: unknown) => void);
    return () => this.off(event, handler);
  }

  /** Unsubscribe a handler previously passed to on(). */
  off<E extends keyof WordCanvasEventMap>(event: E, handler: Handler<E>): void {
    this.handlers.get(event)?.delete(handler as (data: unknown) => void);
  }

  private emit(ev: WordCanvasEvent): void {
    const { type, ...data } = ev;
    for (const h of this.handlers.get(type) ?? []) h(data as unknown);
  }

  /** Resolves once the editor is mounted and ready. */
  whenReady(): Promise<EditorHandle> {
    return this.ready;
  }

  /** Create a child document that shares this editor's live styles, fonts and
   *  theme. Use it to render a real, document-styled preview of a content slice
   *  (e.g. a style sample, a field result, or any blocks/fragment/OOXML) onto a
   *  canvas — instead of an HTML approximation — or to mount a canvas-native
   *  editor over a slice. Returned synchronously; render() calls before the editor
   *  is ready are buffered and flush on ready. (mountEditor() requires the editor
   *  to be ready — await whenReady() first.) */
  createChild(): ChildDocument {
    let real: ChildDocument | null = null;
    let pendingRender: (() => void) | null = null;
    let destroyed = false;
    void this.ready.then((h) => {
      if (destroyed) return;
      real = h.createChild();
      pendingRender?.();
      pendingRender = null;
    });
    return {
      render: (target: HTMLElement, content: ChildContent, opts?: ChildRenderOptions): void => {
        const op = (): void => real?.render(target, content, opts);
        if (real) op();
        else pendingRender = op; // keep only the latest pre-ready render
      },
      mountEditor: (target: HTMLElement, content: ChildContent, opts?: ChildRenderOptions): ChildEditorHandle => {
        if (!real) throw new Error("WordCanvas.createChild(): await whenReady() before mountEditor()");
        return real.mountEditor(target, content, opts);
      },
      destroy: (): void => {
        destroyed = true;
        real?.destroy();
        real = null;
      },
    };
  }

  /** Open a .docx. When online, auto-publishes it and surfaces a share link. */
  async openDocx(file: File | ArrayBuffer): Promise<void> {
    return (await this.ready).openDocx(file);
  }

  /** Export the current document to a .docx Blob — track changes baked to the
   *  original baseline, exactly like the toolbar's Export. Wire it to your own
   *  Save button and POST it anywhere. Resolves once the editor is ready. */
  async exportDocx(): Promise<Blob> {
    return (await this.ready).exportDocx();
  }

  /** Export the current document to a PDF Blob (see `exportDocx`). */
  async exportPdf(): Promise<Blob> {
    return (await this.ready).exportPdf();
  }

  /** Replace the open document with a programmatically-built one (e.g. a
   *  DocumentBuilder result). The input is cloned. Like openDocx, this starts
   *  a NEW document: undo history and any live collab session are dropped
   *  (the next share() forks). Zoom and scroll position are preserved, so
   *  calling this on every data change gives a stable live preview. */
  async setDocument(doc: Document): Promise<void> {
    (await this.ready).setDocument(doc);
  }

  /** Publish the current document and resolve its shareable link (online only). */
  async share(): Promise<string> {
    return (await this.ready).share();
  }

  // ---- review layer (track changes + comments) ----------------------------

  /** Current editor mode. Returns null before the editor is ready. */
  getMode(): EditMode | null {
    return this.handle?.getMode() ?? null;
  }

  /** Switch mode. Resolves to false if the mode isn't allowed (or not ready). */
  async setMode(mode: EditMode): Promise<boolean> {
    return (await this.ready).setMode(mode);
  }

  /** Snapshot of the review overlay (suggestions + comment threads). */
  async getReview(): Promise<ReviewLayer> {
    return (await this.ready).getReview();
  }

  /** The effective @-mentionable roster: the configured base PLUS whoever is
   *  live-editing the document right now (auto-merged from presence). */
  async getKnownUsers(): Promise<UserInfo[]> {
    return (await this.ready).getKnownUsers();
  }

  /** Replace the configured base roster (live editors are still merged on top). */
  async setKnownUsers(users: UserInfo[]): Promise<void> {
    (await this.ready).setKnownUsers(users);
  }

  async acceptSuggestion(id: string): Promise<void> {
    (await this.ready).acceptSuggestion(id);
  }
  async rejectSuggestion(id: string): Promise<void> {
    (await this.ready).rejectSuggestion(id);
  }
  async acceptAllSuggestions(): Promise<void> {
    (await this.ready).acceptAllSuggestions();
  }
  async rejectAllSuggestions(): Promise<void> {
    (await this.ready).rejectAllSuggestions();
  }

  /** Add a comment thread anchored to the current selection (a rich-text body;
   *  `mentions` are @-tagged users). Resolves to the thread id, or null if
   *  there's no selection. */
  async addComment(body: Fragment, mentions?: UserInfo[]): Promise<string | null> {
    return (await this.ready).addComment(body, mentions);
  }
  async replyToComment(threadId: string, body: Fragment, mentions?: UserInfo[]): Promise<void> {
    (await this.ready).replyToComment(threadId, body, mentions);
  }
  async resolveThread(threadId: string, resolved = true): Promise<void> {
    (await this.ready).resolveThread(threadId, resolved);
  }

  /** Current selection/caret. Returns null before the editor is ready or when
   *  the editor isn't focused. */
  getSelection(): DocSelection | null {
    return this.handle?.getSelection() ?? null;
  }

  /** Insert plain text at the caret (replacing any selection). */
  async insertText(text: string): Promise<void> {
    (await this.ready).insertText(text);
  }

  getDocId(): string | null {
    return this.handle?.getDocId() ?? null;
  }

  getShareLink(): string | null {
    return this.handle?.getShareLink() ?? null;
  }

  destroy(): void {
    this.handle?.destroy();
    this.handle = null;
    this.handlers.clear();
  }
}
