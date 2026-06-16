// Public types for @forevka/wordcanvas/export — the headless Document → .docx/.pdf
// export pipeline. Hand-written + self-contained (model types from ./model). The
// `node` build embeds the bundled clone fonts (read from disk) and a bundled
// pdfkit; the `default` build is the browser variant. Call (and await)
// installMeasureHost() from "@forevka/wordcanvas/export/measure" before exporting
// so layout/PDF rendering have a DOM-free measurement context.

import type { Document } from "./model";

export type { Document } from "./model";

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

/** src → bytes for embedded images, resolved by the caller (blob: URLs are invalid
 *  in the worker / on Node) and passed through to the pipeline. */
export type ImageBytes = Record<string, Uint8Array>;

/** Render a document to PDF or DOCX bytes. */
export declare function runExport(
  doc: Document,
  format: ExportFormat,
  images?: ImageBytes,
): Promise<ExportResult>;
