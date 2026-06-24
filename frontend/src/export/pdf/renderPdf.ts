// LayoutTree -> PDF via pdfkit. Page-accurate: it lays the document out with the
// SAME engine the editor uses (over the fontkit measure host, so widths match the
// embedded fonts), then paints each page with paintBlock — the inverse of the
// canvas renderer. Output is vector + selectable text, fonts subset-embedded.
//
// Coordinates: the model is CSS px (96dpi); PDF is points (72dpi). Each page is
// sized in points and the whole page is scaled by 72/96, so paintBlock keeps
// drawing in document px, unaware of the unit change.

import PDFDocument from "pdfkit";
import type { Document } from "@cw/shared";
import { detectTocHeadings, textOfRuns } from "@cw/shared";
import type { PlacedBlock } from "../../layout/layoutTree";
import type { ExportResult, ImageBytes } from "../types";
import { WarningSink } from "../warnings";
import { installMeasureHost } from "../shared/measureHost";
import { resolveFont, registerCustomFontBytes, clearCustomFontBytes } from "../shared/fontRegistry";
import {
  createFontRegistry,
  defaultFontRegistry,
  setActiveFontRegistry,
  type CustomFontPayload,
  type CustomFontRegistry,
} from "../../fonts/customRegistry";
import { paintBlock, type PaintCtx } from "./paintBlock";
import {
  COLUMN_SEPARATOR_COLOR,
  FOOTNOTE_RULE_COLOR,
  FOOTNOTE_RULE_WIDTH_FRACTION,
  pageBorderSegments,
} from "../../paint/paintStyle";

const PT = 72 / 96; // px -> pt

export interface RenderPdfOptions {
  images?: ImageBytes;
  /** This export job's custom fonts (defs + fetched face bytes). Scoped to the
   *  render so concurrent jobs never share custom-font state. */
  fonts?: CustomFontPayload;
}

/** pdfkit's outline addItem accepts these at runtime (@types only declares
 *  `expanded`); fit:false + top makes the destination an XYZ at that y. */
interface OutlineItemOptions {
  pageNumber?: number;
  top?: number;
  fit?: boolean;
  expanded?: boolean;
}

/** Public entry: own a per-job CustomFontRegistry for the whole render so concurrent
 *  exports never share custom-font state, register this job's custom defs+bytes and
 *  make the registry active, then render. The registry's bytes are dropped (and the
 *  process default restored) in finally. */
export async function renderPdf(doc: Document, opts: RenderPdfOptions = {}): Promise<ExportResult> {
  const fontReg = createFontRegistry();
  try {
    // Inside the try so the finally still restores the active registry and clears
    // this job's bytes if registerCustomFontBytes throws on an invalid face.
    setActiveFontRegistry(fontReg);
    if (opts.fonts) {
      fontReg.register(opts.fonts.defs);
      registerCustomFontBytes(opts.fonts.faces);
    }
    return await renderPdfInner(doc, opts, fontReg);
  } finally {
    clearCustomFontBytes(fontReg);
    setActiveFontRegistry(defaultFontRegistry());
  }
}

async function renderPdfInner(doc: Document, opts: RenderPdfOptions, fontReg: CustomFontRegistry): Promise<ExportResult> {
  await installMeasureHost();
  // Imported after the host is installed: the engine measures lazily at layout(),
  // but this keeps the ordering contract explicit and avoids any module-load DOM.
  const { createLayoutEngine } = await import("../../layout/engine");
  // layout() re-asserts fontReg active (no await between it and the paint loop), so
  // the whole resolution span sees this job's fonts even if another job swapped the
  // active registry during the awaits above.
  const tree = createLayoutEngine(fontReg).layout(doc);

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

  // --- internal navigation -------------------------------------------------
  // Targets = every bookmark-start block ∪ every TOC entry's heading. Each gets
  // a named destination at its top (keyed by blockId); TOC targets and "#anchor"
  // links resolve to those names so the PDF jumps like Ctrl+click in the editor.
  const targetBlockIds = new Set<string>();
  for (const r of Object.values(doc.bookmarks ?? {})) targetBlockIds.add(r.start.blockId);
  for (const b of doc.blocks) {
    if (b.kind === "paragraph" && b.style.tocEntry) targetBlockIds.add(b.style.tocEntry.targetId);
  }

  // First-occurrence page index + top-Y (px) of every laid-out block (incl. cells).
  const pos = new Map<string, { pageIndex: number; y: number }>();
  const indexPositions = (blocks: PlacedBlock[], pageIndex: number): void => {
    for (const b of blocks) {
      if (!pos.has(b.blockId)) pos.set(b.blockId, { pageIndex, y: b.y });
      if (b.table) for (const row of b.table.rows) for (const cell of row.cells) indexPositions(cell.blocks, pageIndex);
    }
  };
  for (const page of tree.pages) indexPositions(page.blocks, page.index);

  const destName = (blockId: string): string | undefined =>
    targetBlockIds.has(blockId) && pos.has(blockId) ? `b_${blockId}` : undefined;

  // Destinations grouped by the page they live on — registered while that page
  // is current (addNamedDestination anchors to pdf.page and ignores the CTM).
  const destsByPage = new Map<number, { name: string; y: number }[]>();
  for (const id of targetBlockIds) {
    const p = pos.get(id);
    if (!p) continue;
    let list = destsByPage.get(p.pageIndex);
    if (!list) destsByPage.set(p.pageIndex, (list = []));
    list.push({ name: `b_${id}`, y: p.y });
  }

  const ctx: PaintCtx = {
    doc: pdf,
    font: fontName,
    image: (src) => images[src],
    warn: (code, detail) => warnings.warn(code, detail),
    destName: (ref) => {
      const id = ref.blockId ?? (ref.anchor ? doc.bookmarks?.[ref.anchor]?.start.blockId : undefined);
      return id ? destName(id) : undefined;
    },
  };

  for (const page of tree.pages) {
    pdf.addPage({ size: [page.widthPx * PT, page.heightPx * PT], margin: 0 });
    // Register this page's named destinations (points; pdfkit flips y by page height).
    for (const d of destsByPage.get(page.index) ?? []) {
      pdf.addNamedDestination(d.name, "XYZ", null, d.y * PT, null);
    }
    pdf.save();
    pdf.scale(PT); // every subsequent draw is in document px

    pdf.rect(0, 0, page.widthPx, page.heightPx).fill(page.pageColorHex ?? "#ffffff");

    // page borders (w:pgBorders) — behind content, identical geometry to canvas.
    if (page.pageBorders) {
      for (const s of pageBorderSegments(page.pageBorders, page.widthPx, page.heightPx, page.marginPx)) {
        pdf.lineWidth(s.widthPx).strokeColor(s.color);
        if (s.dash.length) pdf.dash(s.dash[0]!, { space: s.dash[1] ?? s.dash[0]! });
        else pdf.undash();
        pdf.moveTo(s.x1, s.y1).lineTo(s.x2, s.y2).stroke();
      }
      pdf.undash();
    }

    for (const block of page.blocks) paintBlock(ctx, block);

    // column separator rules (w:cols/@w:sep).
    if (page.columnSeparatorsX) {
      pdf.undash().lineWidth(1).strokeColor(COLUMN_SEPARATOR_COLOR);
      for (const x of page.columnSeparatorsX) {
        pdf.moveTo(x + 0.5, page.contentTopPx).lineTo(x + 0.5, page.contentBottomPx).stroke();
      }
    }

    if (page.footnoteRuleY !== undefined) {
      const cw = page.widthPx - page.marginPx.left - page.marginPx.right;
      pdf.undash().lineWidth(1).strokeColor(FOOTNOTE_RULE_COLOR);
      pdf
        .moveTo(page.marginPx.left, page.footnoteRuleY + 0.5)
        .lineTo(page.marginPx.left + cw * FOOTNOTE_RULE_WIDTH_FRACTION, page.footnoteRuleY + 0.5)
        .stroke();
    }

    if (page.header) for (const b of page.header) paintBlock(ctx, b);
    if (page.footer) for (const b of page.footer) paintBlock(ctx, b);

    pdf.restore();
  }

  addOutline(pdf, doc, pos);

  pdf.end();
  const bytes = await bytesPromise;
  return { bytes, warnings: warnings.list() };
}

/** Build the PDF document outline (bookmarks panel) from the document's headings,
 *  nesting by level. Called after every page exists so item destinations can
 *  reference pages by index. Headings never laid out are skipped. */
function addOutline(
  pdf: PDFKit.PDFDocument,
  doc: Document,
  pos: Map<string, { pageIndex: number; y: number }>,
): void {
  const headings = detectTocHeadings(doc, undefined, 9);
  if (headings.length === 0) return;
  const stack: { level: number; item: PDFKit.PDFOutline }[] = [];
  for (const { block, level } of headings) {
    const p = pos.get(block.id);
    if (!p) continue;
    while (stack.length > 0 && stack[stack.length - 1]!.level >= level) stack.pop();
    const parent = stack.length > 0 ? stack[stack.length - 1]!.item : pdf.outline;
    const title = textOfRuns(block.runs).trim() || "(untitled)";
    const opts: OutlineItemOptions = { pageNumber: p.pageIndex, top: p.y * PT, fit: false, expanded: true };
    const item = parent.addItem(title, opts);
    stack.push({ level, item });
  }
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
