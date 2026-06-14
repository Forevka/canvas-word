// Server-side headless TOC generation: fill in a TOC field's result using this
// project's layout engine. Patch-in-place (the field result is spliced into the
// original document.xml), so it's drift-free. installMeasureHost() makes layout
// run in Node; the concurrency gate mirrors serverExport/serverRecalc.

import { installMeasureHost } from "@forevka/wordcanvas/export/measure";
import { generateTocInDocx, type GenerateTocResult } from "@forevka/wordcanvas/generate-toc";
import type { TocOptions } from "@cw/shared";

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

/** Generate the TOC field result in a .docx and return the patched bytes. */
export async function generateTocBytes(bytes: Uint8Array, opts: TocOptions = {}): Promise<GenerateTocResult> {
  await acquire();
  try {
    await installMeasureHost();
    return generateTocInDocx(bytes, opts);
  } finally {
    release();
  }
}
