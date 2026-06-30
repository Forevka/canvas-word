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
import { createLayoutEngine, type LayoutEngineOptions } from "../layout/engine";
import { ARABIC_FONT_FAMILY, CJK_FONT_FAMILY } from "../fonts/clones";
import { pageOfBlockMap } from "../recalc/recalcToc";

/** CJK + Arabic tuning for an export job (mirror of the editor's `cjk` config). The
 *  CJK fallback defaults to the bundled `NotoSansSC` face and the Arabic fallback to
 *  `NotoSansArabic` — both are built-ins that embed without a `fonts` payload entry.
 *  Pass `fallbackFont: ""` / `arabicFallbackFont: ""` to opt out per script. */
export interface CjkExportConfig {
  locale?: string;
  fallbackFont?: string;
  arabicFallbackFont?: string;
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
  // Resolve the CJK config once and THREAD it into every layout this job runs
  // (renderPdf's layout + the docx live-TOC pass). No process-global state and no
  // lock: each layout engine re-asserts its own fallback/locale synchronously at the
  // top of each (await-free) layout pass, so concurrent exports never cross-
  // contaminate. The fallback defaults to the bundled CJK face so a consumer that
  // calls runExport without a cjk config still gets Chinese glyphs (not tofu);
  // `fallbackFont: ""` opts out (the engine treats empty as "no fallback").
  const cjkOpts: LayoutEngineOptions = {
    cjkFallback: cjk?.fallbackFont ?? CJK_FONT_FAMILY,
    arabicFallback: cjk?.arabicFallbackFont ?? ARABIC_FONT_FAMILY,
    ...(cjk?.locale !== undefined ? { cjkLocale: cjk.locale } : {}),
  };
  return runExportInner(doc, format, images, cjkOpts, fonts);
}

async function runExportInner(
  doc: Document,
  format: ExportFormat,
  images: ImageBytes,
  cjkOpts: LayoutEngineOptions,
  fonts?: CustomFontPayload,
): Promise<ExportResult> {
  // Each job owns its custom-font state: renderPdf builds + tears down a per-job
  // CustomFontRegistry internally, so a later export that omits a family can't
  // inherit an earlier job's fonts.
  if (format === "pdf") return renderPdf(doc, { images, cjk: cjkOpts, ...(fonts ? { fonts } : {}) });

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
    const tocPages = pageOfBlockMap(createLayoutEngine(fontReg, cjkOpts).layout(doc));
    return await writeDocx(doc, images, tocPages);
  } finally {
    clearCustomFontBytes(fontReg);
    setActiveFontRegistry(defaultFontRegistry());
  }
}
