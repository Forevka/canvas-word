// Pure export entry — no DOM, no worker. Runs in the worker AND under vitest.
// Image bytes must already be resolved (the main thread does that; blob: URLs are
// invalid here).

import type { Block, Document } from "@cw/shared";
import type { ExportFormat, ExportResult, ImageBytes } from "./types";
import { renderPdf } from "./pdf/renderPdf";
import { writeDocx } from "./docx/writeDocx";
import { installMeasureHost } from "./shared/measureHost";
import { registerCustomFontBytes } from "./shared/fontRegistry";
import { registerCustomFonts, type CustomFontPayload } from "../fonts/customRegistry";
import { createLayoutEngine } from "../layout/engine";
import { pageOfBlockMap } from "../recalc/recalcToc";

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
): Promise<ExportResult> {
  // Register custom fonts (defs → resolution/metrics; bytes → fontkit/pdfkit) BEFORE
  // any layout. This worker/process has its own global registry, so the config must
  // be (re)applied per export — it's keyed by family, so repeats are cheap.
  if (fonts) {
    registerCustomFonts(fonts.defs);
    registerCustomFontBytes(fonts.faces);
  }
  if (format === "pdf") return renderPdf(doc, { images });
  // Live TOC field export needs each heading's real page for the cached PAGEREF
  // result — lay the doc out once (measure host is idempotent).
  let tocPages: Map<string, number> | undefined;
  if (hasTocEntries(doc)) {
    await installMeasureHost();
    tocPages = pageOfBlockMap(createLayoutEngine().layout(doc));
  }
  return writeDocx(doc, images, tocPages);
}
