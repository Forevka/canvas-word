// Runtime config bridge between the WordCanvas wrapper and the editor app
// module. WordCanvas sets the runtime, then dynamic-imports editorApp, whose
// top-level bootstrap reads it via getRuntime() and calls onReady with a handle.
//
// v1 supports a single embedded editor per page (editorApp's top-level code runs
// once on first import; its CSS is id-based). Constructing a second WordCanvas
// throws.

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
  /** Called once the editor is mounted and ready. */
  onReady?: ((handle: EditorHandle) => void) | undefined;
  /** Sink for collaboration events (presence, share, ready) → WordCanvas.on(...). */
  onEvent?: ((ev: WordCanvasEvent) => void) | undefined;
}

let current: WordCanvasRuntime | null = null;

export function setRuntime(r: WordCanvasRuntime): void {
  current = r;
}

export function getRuntime(): WordCanvasRuntime {
  if (!current) throw new Error("WordCanvas runtime not configured (set it before importing editorApp)");
  return current;
}
