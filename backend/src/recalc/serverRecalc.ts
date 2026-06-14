// Server-side TOC recalculation: use THIS project's layout engine to compute each
// heading's real page, then PATCH those numbers into the uploaded .docx in place —
// rewriting only the cached PAGEREF result digits and preserving every field
// (live TOC, PAGEREF, footer PAGE/NUMPAGES) and every other part untouched. This
// is the faithful, drift-free path for a renderer-less caller (e.g. a C# pipeline
// that emits a real Word TOC field but cannot compute page numbers).
//
// installMeasureHost() injects the DOM-free fontkit measure context so layout runs
// in Node. We import only to LAY OUT — the output is the original bytes, not a
// re-export — so there is no import↔export pagination drift.

import { runImport } from "@forevka/wordcanvas/import";
import { installMeasureHost } from "@forevka/wordcanvas/export/measure";
import { patchTocFromLayout } from "@forevka/wordcanvas/recalc-docx";

// Layout is CPU-heavy; cap concurrency so a burst can't starve the event loop /
// collab traffic. Mirrors serverExport.ts.
const MAX_CONCURRENT = Number(process.env.RECALC_CONCURRENCY ?? process.env.EXPORT_CONCURRENCY ?? 2);
let active = 0;
const queue: (() => void)[] = [];

function acquire(): Promise<void> {
  if (active < MAX_CONCURRENT) {
    active++;
    return Promise.resolve();
  }
  return new Promise<void>((resolve) => queue.push(resolve));
}

function release(): void {
  active--;
  const next = queue.shift();
  if (next) {
    active++;
    next();
  }
}

export interface RecalcDocxResult {
  bytes: Uint8Array;
  /** TOC/cross-reference page numbers rewritten to a new value. */
  changed: number;
  /** PAGEREF results left untouched because the cached value was non-arabic
   *  (roman/alpha page numbers are not recalculated — see patchTocDocx). */
  skipped: number;
  warnings: { code: string; message: string }[];
}

/** Recalculate a .docx's TOC page numbers and return the patched .docx. Stateless;
 *  the file is preserved except the cached page numbers. */
export async function recalcDocxBytes(bytes: Uint8Array): Promise<RecalcDocxResult> {
  const result = runImport(bytes);

  await acquire();
  try {
    await installMeasureHost();
    const { bytes: out, changed, skipped } = patchTocFromLayout(bytes, result.doc);
    return { bytes: out, changed, skipped, warnings: result.warnings };
  } finally {
    release();
  }
}
