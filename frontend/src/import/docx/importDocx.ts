// Main-thread API — the only module the app imports.
//
//   const { doc, warnings } = await importDocx(file);
//
// Runs the pipeline in a lazily-created singleton worker (terminated after 30s
// idle); the file buffer is transferred, not cloned. Import is perceptually
// free for the UI no matter how big the document is.

import { ImportError, type FromWorker, type ImportPhase, type ImportResult, type ToWorker } from "./types";
import { BAND_CONTAINERS, type Block } from "@cw/shared";

export type { ImportPhase, ImportResult, ImportWarning } from "./types";
export { ImportError } from "./types";

export interface ImportOptions {
  onProgress?: (phase: ImportPhase, pct: number) => void;
}

interface Pending {
  resolve: (result: ImportResult) => void;
  reject: (err: Error) => void;
  onProgress?: ((phase: ImportPhase, pct: number) => void) | undefined;
}

const IDLE_TERMINATE_MS = 30_000;

let worker: Worker | null = null;
let idleTimer: ReturnType<typeof setTimeout> | undefined;
let nextId = 0;
const pending = new Map<number, Pending>();

function ensureWorker(): Worker {
  if (idleTimer !== undefined) {
    clearTimeout(idleTimer);
    idleTimer = undefined;
  }
  if (worker) return worker;
  worker = new Worker(new URL("./worker.ts", import.meta.url), { type: "module" });
  worker.onmessage = (e: MessageEvent<FromWorker>) => {
    const msg = e.data;
    const job = pending.get(msg.id);
    if (!job) return;
    switch (msg.type) {
      case "progress":
        job.onProgress?.(msg.phase, msg.pct);
        return;
      case "done":
        pending.delete(msg.id);
        // Re-home blobs onto the main thread BEFORE resolving (and before the
        // worker can idle-terminate, which revokes its blob: URLs) so images
        // survive later repaints. The worker is still alive here, so its URLs
        // are fetchable.
        void rehomeMedia(msg.result).then(() => {
          job.resolve(msg.result);
          if (pending.size === 0) scheduleIdleTerminate();
        });
        return;
      case "error":
        pending.delete(msg.id);
        job.reject(new ImportError(msg.code, msg.message));
        if (pending.size === 0) scheduleIdleTerminate();
        return;
    }
  };
  worker.onerror = (e) => {
    // Worker-level failure (e.g. failed to load): fail everything in flight.
    const err = new Error(`docx import worker failed: ${e.message}`);
    for (const job of pending.values()) job.reject(err);
    pending.clear();
    worker?.terminate();
    worker = null;
  };
  return worker;
}

function scheduleIdleTerminate(): void {
  if (idleTimer !== undefined) clearTimeout(idleTimer);
  idleTimer = setTimeout(() => {
    worker?.terminate();
    worker = null;
    idleTimer = undefined;
  }, IDLE_TERMINATE_MS);
}

/** The pipeline runs in the worker, so the media blob: URLs it mints belong to
 *  the worker's context and are revoked when it idle-terminates — after which any
 *  repaint (a page scrolled into view, a background-tab paint) reloads a dead URL
 *  and shows a permanent gray box. Re-create each blob on the MAIN thread, where
 *  it lives as long as the document, and rewrite the image srcs to match. Called
 *  while the worker is still alive, so its URLs are still fetchable. */
async function rehomeMedia(result: ImportResult): Promise<void> {
  if (result.mediaUrls.length === 0) return;
  const remap = new Map<string, string>();
  await Promise.all(
    result.mediaUrls.map(async (oldUrl) => {
      try {
        const blob = await (await fetch(oldUrl)).blob();
        remap.set(oldUrl, URL.createObjectURL(blob));
      } catch {
        // Worker URL already gone / unreachable — leave the original src as-is.
      }
    }),
  );
  if (remap.size === 0) return;
  const rewrite = (blocks: Block[] | undefined): void => {
    for (const b of blocks ?? []) {
      if (b.kind === "image") {
        const next = remap.get(b.src);
        if (next) b.src = next;
      } else if (b.kind === "table") {
        for (const row of b.rows) for (const cell of row.cells) rewrite(cell.blocks);
      }
    }
  };
  rewrite(result.doc.blocks);
  for (const band of BAND_CONTAINERS) rewrite(result.doc.section[band]);
  // The app now references the main-thread URLs; the worker's originals die with
  // it. Hand back the live set so the caller revokes the right ones on discard.
  result.mediaUrls = [...remap.values()];
}

export async function importDocx(file: File | Blob | ArrayBuffer, opts: ImportOptions = {}): Promise<ImportResult> {
  const buf = file instanceof ArrayBuffer ? file : await file.arrayBuffer();
  const w = ensureWorker();
  const id = nextId++;
  return new Promise<ImportResult>((resolve, reject) => {
    pending.set(id, { resolve, reject, onProgress: opts.onProgress });
    const msg: ToWorker = { id, buf };
    w.postMessage(msg, [buf]); // transfer, zero-copy
  });
}
