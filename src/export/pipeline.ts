// Pure export entry — no DOM, no worker. Runs in the worker AND under vitest.
// Image bytes must already be resolved (the main thread does that; blob: URLs are
// invalid here).

import type { Document } from "../model/document";
import type { ExportFormat, ExportResult, ImageBytes } from "./types";
import { renderPdf } from "./pdf/renderPdf";
import { writeDocx } from "./docx/writeDocx";

export async function runExport(
  doc: Document,
  format: ExportFormat,
  images: ImageBytes = {},
): Promise<ExportResult> {
  if (format === "pdf") return renderPdf(doc, { images });
  return writeDocx(doc, images);
}
