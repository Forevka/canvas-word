// pdfkit (and its readable-stream / zlib deps) reference the Node globals
// `process` and `Buffer` as bare identifiers. The node-polyfills plugin provides
// the MODULES but its global injection doesn't reliably reach a separate worker
// entry, so we install the globals ourselves — imported FIRST in worker.ts, before
// the pipeline (and thus pdfkit) is evaluated, since ES imports run in source
// order. On Node (vitest) these resolve to the real builtins and the guards
// no-op, so this is safe everywhere.

import process from "process";
import { Buffer } from "buffer";

const g = globalThis as unknown as { process?: unknown; Buffer?: unknown; global?: unknown };
if (typeof g.process === "undefined") g.process = process;
if (typeof g.Buffer === "undefined") g.Buffer = Buffer;
if (typeof g.global === "undefined") g.global = globalThis;
