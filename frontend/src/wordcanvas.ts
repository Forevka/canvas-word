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

import type { EditMode, EditorHandle, Participant, WordCanvasEvent, WordCanvasRuntime } from "./app/runtime";
import type { Document, Fragment, ReviewLayer, UserInfo } from "@cw/shared";

export type { Document, UserInfo, Participant, EditMode, ReviewLayer, Fragment };

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
        readonly: opts.readonly,
        mode: opts.mode,
        allowedModes: opts.allowedModes,
        onReady: (h) => {
          this.handle = h;
          resolve(h);
        },
        onEvent: (ev) => this.emit(ev),
      };
      // Lazy-load the editor chunk, then mount this instance. Each call mounts an
      // independent editor (the chunk's module-eval cost is shared, the mount is
      // per-instance), so multiple WordCanvas instances coexist on one page.
      void import("./editorApp").then((m) => m.mountEditorApp(runtime));
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

  /** Open a .docx. When online, auto-publishes it and surfaces a share link. */
  async openDocx(file: File | ArrayBuffer): Promise<void> {
    return (await this.ready).openDocx(file);
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

  /** Add a comment thread anchored to the current selection (a rich-text body).
   *  Resolves to the thread id, or null if there's no selection. */
  async addComment(body: Fragment): Promise<string | null> {
    return (await this.ready).addComment(body);
  }
  async replyToComment(threadId: string, body: Fragment): Promise<void> {
    (await this.ready).replyToComment(threadId, body);
  }
  async resolveThread(threadId: string, resolved = true): Promise<void> {
    (await this.ready).resolveThread(threadId, resolved);
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
