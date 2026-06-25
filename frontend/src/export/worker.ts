// Export worker — transport only; all logic is in pipeline.ts. Mirrors the import
// worker. The result bytes are transferred back zero-copy.

import "./shared/workerGlobals"; // MUST be first — installs process/Buffer before pdfkit loads
import { runExport } from "./pipeline";
import type { FromExportWorker, ToExportWorker } from "./types";

interface WorkerScope {
  postMessage(message: FromExportWorker, transfer?: Transferable[]): void;
  onmessage: ((e: MessageEvent<ToExportWorker>) => void) | null;
}

const ctx = self as unknown as WorkerScope;

// Run jobs one at a time. runExport mutates the process-global analyzer state
// (CJK locale / fallback font) for a job's duration and restores it after, so
// overlapping jobs would clobber each other — chain each onto a shared queue.
let queue: Promise<void> = Promise.resolve();

ctx.onmessage = (e: MessageEvent<ToExportWorker>): void => {
  const { id, doc, format, images, fonts, cjk } = e.data;
  queue = queue.then(async () => {
    try {
      const { bytes, warnings } = await runExport(doc, format, images, fonts, cjk);
      ctx.postMessage({ id, type: "done", bytes, warnings }, [bytes.buffer]);
    } catch (err: unknown) {
      ctx.postMessage({ id, type: "error", message: err instanceof Error ? err.message : String(err) });
    }
  });
};
