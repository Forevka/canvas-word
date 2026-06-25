// Shared export types — small and transport-friendly (worker postMessage).

import type { Document } from "@cw/shared";
import type { CustomFontPayload } from "../fonts/customRegistry";

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
  /** Custom fonts (defs + fetched face bytes), resolved on the main thread. */
  fonts?: CustomFontPayload;
  /** CJK locale + fallback-font tuning (mirrors the editor's `cjk` config). The
   *  fallback family must also be among `fonts` to embed. */
  cjk?: { locale?: string; fallbackFont?: string };
}

export type FromExportWorker =
  | { id: number; type: "progress"; phase: ExportPhase; pct: number }
  | { id: number; type: "done"; bytes: Uint8Array; warnings: ExportWarning[] }
  | { id: number; type: "error"; message: string };
