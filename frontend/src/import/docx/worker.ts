// Worker entry — message transport only; all logic lives in pipeline.ts.
// tsconfig uses the DOM lib (this is mostly a window app), so the worker
// globals are typed locally instead of pulling in lib.webworker.

import { runImport } from "./pipeline";
import { ImportError, type FromWorker, type ToWorker } from "./types";

interface WorkerScope {
  postMessage(message: FromWorker): void;
  onmessage: ((e: MessageEvent<ToWorker>) => void) | null;
}

const ctx = self as unknown as WorkerScope;

ctx.onmessage = (e: MessageEvent<ToWorker>) => {
  const { id, buf } = e.data;
  try {
    const result = runImport(new Uint8Array(buf), (phase, pct) => {
      ctx.postMessage({ id, type: "progress", phase, pct });
    });
    ctx.postMessage({ id, type: "done", result });
  } catch (err) {
    if (err instanceof ImportError) {
      ctx.postMessage({ id, type: "error", code: err.code, message: err.message });
    } else {
      ctx.postMessage({
        id,
        type: "error",
        code: "MALFORMED_XML",
        message: err instanceof Error ? err.message : String(err),
      });
    }
  }
};
