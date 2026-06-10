// Public type surface for the @forevka/wordcanvas (WordCanvas) package.
// Hand-written so the published types stay small, self-contained, and stable
// regardless of internal refactors — no dependency on internal workspace types.

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
  /** Open a .docx. When online, auto-publishes it and surfaces a share link. */
  openDocx(file: File | ArrayBuffer): Promise<void>;
  /** Publish the current document and resolve its shareable link (online only). */
  share(): Promise<string>;
  getDocId(): string | null;
  getShareLink(): string | null;
  destroy(): void;
}
