// Runtime config bridge between the WordCanvas wrapper and the editor app
// module. WordCanvas builds a runtime and passes it to mountEditorApp(runtime),
// which mounts one editor and calls onReady with a handle. The runtime is a plain
// per-call value (no module singleton), so multiple editors coexist on one page.

import type { Document, UserInfo } from "@cw/shared";

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
  | { type: "presence"; participants: Participant[] };

/** Handle the editor app exposes back to the WordCanvas wrapper. */
export interface EditorHandle {
  getDocument(): Document;
  /** Replace the open document with a programmatically-built one (e.g. from
   *  the DocumentBuilder). The input is cloned; like openDocx, this is a NEW
   *  document — any live collab session is dropped (the next Share forks).
   *  Discards undo history and unsaved edits; preserves zoom + scroll. */
  setDocument(doc: Document): void;
  /** Open a .docx (auto-publishes when online); resolves when loaded. */
  openDocx(file: File | ArrayBuffer): Promise<void>;
  /** Publish the current document and resolve its shareable link (online only). */
  share(): Promise<string>;
  getDocId(): string | null;
  getShareLink(): string | null;
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
  /** Mount view-only: hide the editing chrome and make every mutation a no-op
   *  (the document is still selectable, copyable, and live for remote edits). */
  readonly?: boolean | undefined;
  /** Called once the editor is mounted and ready. */
  onReady?: ((handle: EditorHandle) => void) | undefined;
  /** Sink for collaboration events (presence, share, ready) → WordCanvas.on(...). */
  onEvent?: ((ev: WordCanvasEvent) => void) | undefined;
}
