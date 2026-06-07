// Main-thread API — the only module the app imports.
//
//   const { doc, warnings } = await importDocx(file);
//
// Runs the pipeline in a lazily-created singleton worker (terminated after 30s
// idle); the file buffer is transferred, not cloned. Import is perceptually
// free for the UI no matter how big the document is.

import { ImportError, type FromWorker, type ImportPhase, type ImportResult, type ToWorker } from "./types";

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
        job.resolve(msg.result);
        break;
      case "error":
        pending.delete(msg.id);
        job.reject(new ImportError(msg.code, msg.message));
        break;
    }
    if (pending.size === 0) scheduleIdleTerminate();
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
