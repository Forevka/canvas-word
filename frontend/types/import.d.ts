// Public types for @forevka/wordcanvas/import — the headless .docx → Document
// import pipeline. Hand-written (like wordcanvas.d.ts) so the published surface
// stays self-contained: the model types come from ./model, never from the private
// @cw/shared workspace package. Runs in the browser (the `default` build) and on
// Node (the `node` build, used by a server to parse uploads).

import type { Document } from "./model";

export type { Document } from "./model";

export type ImportPhase = "unzip" | "styles" | "parse" | "map";

/** A lossy-mapping decision — surfaced, not swallowed (deduplicated by code). */
export interface ImportWarning {
  code: string;
  message: string;
}

/** A collected embedded image (Node/backend path — see RunImportOpts.collectMediaBytes). */
export interface ImportMedia {
  /** The synthetic src the mapper put on the ImageBlock (e.g. "cw-media:0"). */
  src: string;
  bytes: Uint8Array;
  mime: string;
}

export interface RunImportOpts {
  onProgress?: (phase: ImportPhase, pct: number) => void;
  /** Collect raw image bytes (and skip blob: URL creation) for Node/backend
   *  callers that content-address media themselves. */
  collectMediaBytes?: boolean;
}

export interface ImportResult {
  doc: Document;
  /** Every lossy mapping decision. */
  warnings: ImportWarning[];
  /** blob: URLs created for embedded media; caller revokes when the doc is
   *  discarded (browser path). Empty on the Node path. */
  mediaUrls: string[];
  /** Raw embedded image bytes — populated only with { collectMediaBytes: true }.
   *  Each `src` matches the ImageBlock.src the mapper assigned. */
  media: ImportMedia[];
}

/** Parse .docx bytes into the document model. Pure / DOM-free. */
export declare function runImport(
  bytes: Uint8Array,
  onProgress?: (phase: ImportPhase, pct: number) => void,
  opts?: RunImportOpts,
): ImportResult;
