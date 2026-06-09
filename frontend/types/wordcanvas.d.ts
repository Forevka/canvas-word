// Public type surface for the @cw/frontend (WordCanvas) package. Hand-written so
// the published types stay small and stable regardless of internal refactors.

export interface WordCanvasOptions {
  /** Element to mount the editor into. */
  container: HTMLElement;
  /** Backend base URL (e.g. "https://api.example.com"). Online iff provided;
   *  omit for a fully offline editor (no sync, publish, or share). */
  backendUrl?: string;
  /** Join an existing collaboration session on load (online only). */
  collabId?: string;
  /** Override how a share link is surfaced (default: a built-in dialog). */
  onShareLink?: (url: string, docId: string) => void;
}

export declare class WordCanvas {
  constructor(opts: WordCanvasOptions);
  /** Resolves once the editor is mounted and ready. */
  whenReady(): Promise<unknown>;
  /** Open a .docx. When online, auto-publishes it and surfaces a share link. */
  openDocx(file: File | ArrayBuffer): Promise<void>;
  /** Publish the current document and resolve its shareable link (online only). */
  share(): Promise<string>;
  getDocId(): string | null;
  getShareLink(): string | null;
  destroy(): void;
}
