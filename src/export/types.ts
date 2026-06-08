// Shared export types — small and transport-friendly (worker postMessage).

import type { Document } from "../model/document";

export type ExportFormat = "pdf" | "docx";

/** A deduplicated, lossy-mapping note. `count` = how many times it fired. */
export interface ExportWarning {
  code: string;
  detail?: string;
  count: number;
}

export interface ExportResult {
  bytes: Uint8Array;
  warnings: ExportWarning[];
}

export type ExportPhase = "measure" | "layout" | "render" | "zip";

/** src -> bytes for embedded images, resolved on the main thread (blob: URLs are
 *  invalid in the worker / Node) and passed through to the pipeline. */
export type ImageBytes = Record<string, Uint8Array>;

export interface ToExportWorker {
  id: number;
  doc: Document;
  format: ExportFormat;
  images?: ImageBytes;
}

export type FromExportWorker =
  | { id: number; type: "progress"; phase: ExportPhase; pct: number }
  | { id: number; type: "done"; bytes: Uint8Array; warnings: ExportWarning[] }
  | { id: number; type: "error"; message: string };
