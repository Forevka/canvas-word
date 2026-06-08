// LayoutTree -> PDF via pdfkit. Page-accurate: it lays the document out with the
// SAME engine the editor uses (over the fontkit measure host, so widths match the
// embedded fonts), then paints each page with paintBlock — the inverse of the
// canvas renderer. Output is vector + selectable text, fonts subset-embedded.
//
// Coordinates: the model is CSS px (96dpi); PDF is points (72dpi). Each page is
// sized in points and the whole page is scaled by 72/96, so paintBlock keeps
// drawing in document px, unaware of the unit change.

import PDFDocument from "pdfkit";
import type { Document } from "../../model/document";
import type { ExportResult, ImageBytes } from "../types";
import { WarningSink } from "../warnings";
import { installMeasureHost } from "../shared/measureHost";
import { resolveFont } from "../shared/fontRegistry";
import { paintBlock, type PaintCtx } from "./paintBlock";

const PT = 72 / 96; // px -> pt

export interface RenderPdfOptions {
  images?: ImageBytes;
}

export async function renderPdf(doc: Document, opts: RenderPdfOptions = {}): Promise<ExportResult> {
  await installMeasureHost();
  // Imported after the host is installed: the engine measures lazily at layout(),
  // but this keeps the ordering contract explicit and avoids any module-load DOM.
  const { createLayoutEngine } = await import("../../layout/engine");
  const tree = createLayoutEngine().layout(doc);

  const warnings = new WarningSink();
  const images = opts.images ?? {};
  // font:false skips pdfkit's eager Helvetica.afm load (a Node `fs` read that
  // throws in the browser worker, where fs is unavailable) — we always register
  // and select bundled fonts before drawing any text.
  const pdf = new PDFDocument({ autoFirstPage: false, compress: true, font: false } as unknown as ConstructorParameters<
    typeof PDFDocument
  >[0]);
  const bytesPromise = collect(pdf);

  const registered = new Set<string>();
  const fontName = (family: string, bold: boolean, italic: boolean): string => {
    const r = resolveFont(family, bold, italic);
    if (!registered.has(r.file)) {
      pdf.registerFont(r.file, toBuffer(r.bytes));
      registered.add(r.file);
      if (r.substituted) warnings.warn("font-substituted", family);
    }
    return r.file;
  };

  const ctx: PaintCtx = {
    doc: pdf,
    font: fontName,
    image: (src) => images[src],
    warn: (code, detail) => warnings.warn(code, detail),
  };

  for (const page of tree.pages) {
    pdf.addPage({ size: [page.widthPx * PT, page.heightPx * PT], margin: 0 });
    pdf.save();
    pdf.scale(PT); // every subsequent draw is in document px

    pdf.rect(0, 0, page.widthPx, page.heightPx).fill("#ffffff");

    for (const block of page.blocks) paintBlock(ctx, block);

    if (page.footnoteRuleY !== undefined) {
      const cw = page.widthPx - page.marginPx.left - page.marginPx.right;
      pdf.undash().lineWidth(1).strokeColor("#80868b");
      pdf
        .moveTo(page.marginPx.left, page.footnoteRuleY + 0.5)
        .lineTo(page.marginPx.left + cw / 3, page.footnoteRuleY + 0.5)
        .stroke();
    }

    if (page.header) for (const b of page.header) paintBlock(ctx, b);
    if (page.footer) for (const b of page.footer) paintBlock(ctx, b);

    pdf.restore();
  }

  pdf.end();
  const bytes = await bytesPromise;
  return { bytes, warnings: warnings.list() };
}

function collect(pdf: PDFKit.PDFDocument): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = [];
    pdf.on("data", (c: Uint8Array) => chunks.push(c));
    pdf.on("end", () => {
      let total = 0;
      for (const c of chunks) total += c.length;
      const out = new Uint8Array(total);
      let off = 0;
      for (const c of chunks) {
        out.set(c, off);
        off += c.length;
      }
      resolve(out);
    });
    pdf.on("error", reject);
  });
}

function toBuffer(bytes: Uint8Array): Buffer {
  return Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}
