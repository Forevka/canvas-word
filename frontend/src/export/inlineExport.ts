// Main-thread (inline) export — used ONLY when the document contains a custom
// block whose type is registered here. Custom blocks paint via a registered
// function held in the block registry, which is a MAIN-THREAD module singleton;
// the export worker has its own, empty registry, so it can't render them. Running
// the (isomorphic, DOM-free) pipeline inline lets the block's paint() reach the
// PDF via the Canvas2D→pdfkit shim. Every other export still goes through the
// worker (exportDocument gates on this), so the common path stays off the main
// thread.
//
// workerGlobals installs process/Buffer before pdfkit is evaluated — imported
// FIRST, exactly like worker.ts.
import "./shared/workerGlobals";
import { runExport } from "./pipeline";
import type { CjkExportConfig } from "./pipeline";
import type { ExportFormat, ExportResult, ImageBytes } from "./types";
import type { CustomFontPayload } from "../fonts/customRegistry";
import type { Document } from "@cw/shared";

export async function runExportInline(
  doc: Document,
  format: ExportFormat,
  images: ImageBytes,
  fonts: CustomFontPayload | undefined,
  cjk: CjkExportConfig | undefined,
): Promise<ExportResult> {
  return runExport(doc, format, images, fonts, cjk);
}
