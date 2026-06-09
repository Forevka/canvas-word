// WordCanvas — the embeddable package entry.
//
//   import { WordCanvas } from "@cw/frontend";
//   const ed = new WordCanvas({ container, backendUrl: "https://…" });
//
// `backendUrl` is the online/offline kill-switch: when set, opening a document
// auto-publishes it and exposes a shareable link, and edits sync live; when
// omitted, the editor runs fully offline. v1 supports one WordCanvas per page.

import { setRuntime, type EditorHandle } from "./app/runtime";
import { ensureWordCanvasStyles } from "./ui/styles";

export interface WordCanvasOptions {
  /** Element to mount the editor into. */
  container: HTMLElement;
  /** Backend base URL (e.g. "https://api.example.com"). Online iff provided. */
  backendUrl?: string;
  /** Join an existing collaboration session on load (online only). */
  collabId?: string;
  /** Override how a share link is surfaced (default: a built-in dialog). */
  onShareLink?: (url: string, docId: string) => void;
}

let mounted = false;

export class WordCanvas {
  private handle: EditorHandle | null = null;
  private readonly ready: Promise<EditorHandle>;

  constructor(opts: WordCanvasOptions) {
    if (mounted) {
      throw new Error("WordCanvas: only one instance per page is supported in this version");
    }
    mounted = true;
    ensureWordCanvasStyles();
    this.ready = new Promise<EditorHandle>((resolve) => {
      setRuntime({
        container: opts.container,
        backendUrl: opts.backendUrl,
        collabId: opts.collabId,
        onShareLink: opts.onShareLink,
        onReady: (h) => {
          this.handle = h;
          resolve(h);
        },
      });
      // Evaluating the editor app (once) mounts the UI and calls onReady.
      void import("./editorApp");
    });
  }

  /** Resolves once the editor is mounted and ready. */
  whenReady(): Promise<EditorHandle> {
    return this.ready;
  }

  /** Open a .docx. When online, auto-publishes it and surfaces a share link. */
  async openDocx(file: File | ArrayBuffer): Promise<void> {
    return (await this.ready).openDocx(file);
  }

  /** Publish the current document and resolve its shareable link (online only). */
  async share(): Promise<string> {
    return (await this.ready).share();
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
    mounted = false;
  }
}
