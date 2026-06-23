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

// ===== Custom fonts =========================================================

/** Per-style face URLs for a custom font (`regular` required; the rest fall back). */
export interface CustomFontFaces {
  regular: string;
  bold?: string;
  italic?: string;
  boldItalic?: string;
}

/** One custom font: family name (model + render name), face URLs, and REQUIRED
 *  `sizing` ({ ascent, descent } as fractions of em) for pagination parity. */
export interface CustomFontDef {
  family: string;
  faces: CustomFontFaces;
  sizing: { ascent: number; descent: number };
  label?: string;
}

/** Fully-populated custom-fonts config. */
export interface ResolvedFontsConfig {
  disableBuiltin: string[];
  fonts: CustomFontDef[];
}

/** One fetched custom face's bytes, ready for fontkit/pdfkit. */
export interface CustomFontFaceBytes {
  family: string;
  style: "Regular" | "Bold" | "Italic" | "BoldItalic";
  bytes: Uint8Array;
}

/** Custom fonts carried into the export pipeline: defs (for resolution/metrics) +
 *  fetched face bytes (for measuring + embedding). */
export interface CustomFontPayload {
  defs: ResolvedFontsConfig;
  faces: CustomFontFaceBytes[];
}

/** Render a document to PDF or DOCX bytes. Pass `fonts` to embed custom faces (the
 *  caller fetches their bytes; on Node use a disk cache to reuse across renders). */
export declare function runExport(
  doc: Document,
  format: ExportFormat,
  images?: ImageBytes,
  fonts?: CustomFontPayload,
): Promise<ExportResult>;
