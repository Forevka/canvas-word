// Public types for @forevka/wordcanvas/recalc-docx — patch a .docx's cached TOC /
// cross-reference page numbers IN PLACE (drift-free: the file is never re-exported,
// so every field and unmodeled byte survives). Hand-written + self-contained.
// Await installMeasureHost() (from "@forevka/wordcanvas/export/measure") first —
// the layout pass measures text.

import type { Document } from "./model";

export type { Document } from "./model";

export interface PatchResult {
  bytes: Uint8Array;
  /** PAGEREF cached numbers rewritten to a new value. */
  changed: number;
  /** PAGEREF results left untouched because the cached value was non-arabic. */
  skipped: number;
}

/** Rewrite cached PAGEREF page numbers from a bookmark → page-number map,
 *  preserving every field and other part. */
export declare function patchTocPageNumbers(
  docxBytes: Uint8Array,
  pages: Record<string, string>,
): PatchResult;

/** Lay `doc` out and patch the original `docxBytes` from the resulting page map. */
export declare function patchTocFromLayout(docxBytes: Uint8Array, doc: Document): PatchResult;
