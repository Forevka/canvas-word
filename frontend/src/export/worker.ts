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

ctx.onmessage = (e: MessageEvent<ToExportWorker>): void => {
  const { id, doc, format, images, fonts, cjk } = e.data;
  runExport(doc, format, images, fonts, cjk)
    .then(({ bytes, warnings }) => {
      ctx.postMessage({ id, type: "done", bytes, warnings }, [bytes.buffer]);
    })
    .catch((err: unknown) => {
      ctx.postMessage({ id, type: "error", message: err instanceof Error ? err.message : String(err) });
    });
};
