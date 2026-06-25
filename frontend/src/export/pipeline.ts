// Pure export entry — no DOM, no worker. Runs in the worker AND under vitest.
// Image bytes must already be resolved (the main thread does that; blob: URLs are
// invalid here).

import type { Block, Document } from "@cw/shared";
import type { ExportFormat, ExportResult, ImageBytes } from "./types";
import { renderPdf } from "./pdf/renderPdf";
import { writeDocx } from "./docx/writeDocx";
import { installMeasureHost } from "./shared/measureHost";
import { registerCustomFontBytes, clearCustomFontBytes } from "./shared/fontRegistry";
import {
  createFontRegistry,
  defaultFontRegistry,
  setActiveFontRegistry,
  type CustomFontPayload,
} from "../fonts/customRegistry";
import { createLayoutEngine } from "../layout/engine";
import { setCjkFallbackFont, setCjkLocale } from "../layout/prepareCache";
import { pageOfBlockMap } from "../recalc/recalcToc";

/** CJK tuning for an export job (mirror of the editor's `cjk` config). The
 *  fallback family must also appear in the `fonts` payload so it embeds. */
export interface CjkExportConfig {
  locale?: string;
  fallbackFont?: string;
}

/** Does the doc carry a generated table of contents (tocEntry paragraphs)? Their
 *  page numbers are paint-only in the model, so docx export needs a layout pass to
 *  bake the cached PAGEREF numbers. */
function hasTocEntries(doc: Document): boolean {
  const scan = (blocks: Block[]): boolean =>
    blocks.some((b) =>
      b.kind === "paragraph"
        ? !!b.style.tocEntry
        : b.kind === "table" && b.rows.some((r) => r.cells.some((c) => scan(c.blocks))),
    );
  return scan(doc.blocks);
}

export async function runExport(
  doc: Document,
  format: ExportFormat,
  images: ImageBytes = {},
  fonts?: CustomFontPayload,
  cjk?: CjkExportConfig,
): Promise<ExportResult> {
  // CJK fallback/locale touch the process-global analyzer; set for the duration
  // of the job. (The fallback font itself must be in `fonts` to embed.)
  if (cjk) {
    setCjkLocale(cjk.locale);
    setCjkFallbackFont(cjk.fallbackFont);
  }
  try {
    return await runExportInner(doc, format, images, fonts);
  } finally {
    if (cjk) {
      setCjkFallbackFont(null);
      setCjkLocale(undefined);
    }
  }
}

async function runExportInner(
  doc: Document,
  format: ExportFormat,
  images: ImageBytes,
  fonts?: CustomFontPayload,
): Promise<ExportResult> {
  // Each job owns its custom-font state: renderPdf builds + tears down a per-job
  // CustomFontRegistry internally, so a later export that omits a family can't
  // inherit an earlier job's fonts.
  if (format === "pdf") return renderPdf(doc, { images, ...(fonts ? { fonts } : {}) });

  // DOCX writes family names verbatim, but a live TOC field export still needs a
  // layout pass for each heading's real page (cached PAGEREF result). Scope that
  // pass's custom fonts to a per-job registry, just like the PDF path.
  if (!hasTocEntries(doc)) return writeDocx(doc, images);
  await installMeasureHost(); // idempotent; awaited BEFORE the synchronous font span
  const fontReg = createFontRegistry();
  if (fonts) fontReg.register(fonts.defs);
  setActiveFontRegistry(fontReg);
  if (fonts) registerCustomFontBytes(fonts.faces);
  try {
    const tocPages = pageOfBlockMap(createLayoutEngine(fontReg).layout(doc));
    return await writeDocx(doc, images, tocPages);
  } finally {
    clearCustomFontBytes(fontReg);
    setActiveFontRegistry(defaultFontRegistry());
  }
}
