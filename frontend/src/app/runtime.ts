// Runtime config bridge between the WordCanvas wrapper and the editor app
// module. WordCanvas builds a runtime and passes it to mountEditorApp(runtime),
// which mounts one editor and calls onReady with a handle. The runtime is a plain
// per-call value (no module singleton), so multiple editors coexist on one page.

import type { Document, Fragment, ReviewLayer, UserInfo } from "@cw/shared";
import type { EditMode, FieldResolver } from "../index";

export type { EditMode, FieldResolver };

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
  | { type: "reviewChanged"; review: ReviewLayer };

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
  /** Sink for collaboration events (presence, share, ready) → WordCanvas.on(...). */
  onEvent?: ((ev: WordCanvasEvent) => void) | undefined;
  /** Resolve a custom field's content from the host backend. When set, right-
   *  clicking a custom field offers "Update Field (<name>)", which calls this and
   *  splices the returned OOXML in as the field's new result. */
  resolveField?: FieldResolver | undefined;
}
