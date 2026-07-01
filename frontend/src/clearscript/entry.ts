// ClearScript host entry — a single self-contained IIFE bundle that exposes the
// headless WordCanvas pipelines (Builder, DOCX import, PDF/DOCX export) to a
// bare-V8 host (.NET ClearScript). There is NO DOM, NO Node, NO fs/fetch here:
//
//   - Fonts are injected by the host via `installFont(file, bytes)` (the C# side
//     reads the bundled .ttf clones and hands their bytes in). Because that
//     registers the bundled faces up-front, `installMeasureHost()` inside the
//     export pipeline short-circuits its fs/fetch font read and only wires up the
//     fontkit-backed measurement globals — so the exact same layout/measure path
//     the Node backend uses runs under V8 with zero I/O.
//   - pdfkit / fontkit / fflate are bundled with node-builtin polyfills
//     (esbuild-plugin-polyfill-node) + timer/process shims (see build script
//     banner), so the only host responsibility is fonts + bytes marshalling.
//
// The document model never has to cross the V8 boundary as data: the host keeps
// the imported/built `doc` as an opaque script object and only binary blobs
// (docx in, pdf/docx out) are marshalled.

import PDFDocument from "pdfkit";
import { DocumentBuilder } from "../builder/documentBuilder";
import { cm, inches, PAGE_SIZES, pt, twips } from "../builder/units";
import { bytesToDataUrl } from "../builder/media";
import { runImport } from "../import/docx/pipeline";
import { runExport } from "../export/pipeline";
import { registerFont } from "../export/shared/fontRegistry";
import { installMeasureHost } from "../export/shared/measureHost";
import { generateTocInDocx } from "../recalc/generateTocDocx";
import { patchTocFromLayout } from "../recalc/patchTocDocx";
import { recalcToc } from "../recalc/recalcToc";
import { DocumentEditor, generateTocIntoDoc } from "@cw/shared";
import type { Document, TocOptions } from "@cw/shared";
import { getPages } from "../layout/pages";
import {
  findText,
  mapPages,
  queryBlockPath,
  queryBookmarks,
  queryEndnotes,
  queryFields,
  queryFootnotes,
  queryListItems,
  queryParagraphs,
  queryPositionOfText,
  queryRangeText,
  querySdts,
  querySdtValue,
  querySections,
  queryStyles,
  type PageInfo as PageInfoDto,
} from "./queryBridge";
import type { CustomFontPayload } from "../fonts/customRegistry";
import type { CjkExportConfig } from "../export/pipeline";
import type { ImageBytes } from "../export/types";

// pdfkit's PDFDocument is a Node Readable that pushes content proactively (its
// _read is a no-op) and relies on the event loop's flow machinery to emit 'data'/
// 'end'. Under bare V8 (no real event loop) that flow deadlocks mid-stream for
// multi-page docs. Since pdfkit generates everything SYNCHRONOUSLY (deflateSync),
// we capture pushed chunks and replay them as synchronous 'data'+'end' the moment
// end() returns — renderPdf's collect() then resolves with zero dependence on
// stream flow. Patched once, on the SAME bundled pdfkit class renderPdf imports.
{
  const proto = PDFDocument.prototype as unknown as {
    __wcSyncCollect?: boolean;
    end: (...args: unknown[]) => unknown;
    push: (chunk: Uint8Array | null) => boolean;
    emit: (event: string, ...args: unknown[]) => boolean;
    __wcChunks?: Uint8Array[];
  };
  if (!proto.__wcSyncCollect) {
    proto.__wcSyncCollect = true;
    const realEnd = proto.end;
    proto.push = function (this: typeof proto, chunk: Uint8Array | null): boolean {
      if (chunk != null) (this.__wcChunks ||= []).push(chunk);
      return true;
    };
    proto.end = function (this: typeof proto, ...args: unknown[]): unknown {
      // Opt-in deterministic output (globalThis.__wc_fixedDate = epoch ms): pin the
      // PDF CreationDate/ModDate so the body is byte-reproducible across runs/hosts.
      const fixed = (globalThis as unknown as { __wc_fixedDate?: number }).__wc_fixedDate;
      if (fixed != null) {
        const self = this as unknown as { info?: Record<string, unknown> };
        const info = self.info ?? {};
        info.CreationDate = new Date(fixed);
        info.ModDate = new Date(fixed);
        self.info = info;
      }
      const r = realEnd.apply(this, args);
      let chunks = this.__wcChunks ?? [];
      // The trailer /ID is pdfkit's md5(info) wrapped in Buffer.from(WordArray) — its
      // bytes differ between Node's native Buffer and the V8 buffer polyfill, so it is
      // the one non-content field that isn't host-stable. In deterministic mode pin it
      // to a hash of the CONTENT (identical across hosts), making output byte-identical.
      if (fixed != null) chunks = [pinPdfId(concatChunks(chunks))];
      for (const c of chunks) this.emit("data", c);
      this.emit("end");
      return r;
    };
  }
}

function concatChunks(chunks: Uint8Array[]): Uint8Array {
  let total = 0;
  for (const c of chunks) total += c.length;
  const out = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) {
    out.set(c, off);
    off += c.length;
  }
  return out;
}

/** Overwrite the trailer `/ID [<hex> <hex>]` with a host-independent hash of the rest
 *  of the document (the two hex spans excluded), so the file id is deterministic and
 *  identical across the Node and V8 hosts. No-op if the marker isn't found. */
function pinPdfId(buf: Uint8Array): Uint8Array {
  // Find the LAST "/ID [<" (ASCII) — it lives in the trailer near EOF.
  const M = [0x2f, 0x49, 0x44, 0x20, 0x5b, 0x3c]; // "/ID [<"
  let at = -1;
  for (let i = buf.length - M.length; i >= 0 && at < 0; i--) {
    let ok = true;
    for (let j = 0; j < M.length; j++) if (buf[i + j] !== M[j]) { ok = false; break; }
    if (ok) at = i;
  }
  if (at < 0) return buf;
  const h1 = at + 6; // first 32 hex digits
  const h2 = h1 + 32 + 3; // skip "> <"
  if (h2 + 32 > buf.length) return buf;

  // 4-lane FNV-1a (Math.imul = exact 32-bit, identical on any V8) over the buffer,
  // skipping the two id spans so the hash can't depend on the old (host-divergent) id.
  const lanes = [0x811c9dc5, 0x01000193, 0x85ebca77, 0xc2b2ae3d];
  for (let i = 0; i < buf.length; i++) {
    if ((i >= h1 && i < h1 + 32) || (i >= h2 && i < h2 + 32)) continue;
    const l = i & 3;
    lanes[l] = Math.imul(lanes[l]! ^ buf[i]!, 0x01000193) >>> 0;
  }
  let hex = "";
  for (let j = 0; j < 4; j++) hex += ("00000000" + (lanes[j]! >>> 0).toString(16)).slice(-8);
  for (let k = 0; k < 32; k++) {
    const code = hex.charCodeAt(k);
    buf[h1 + k] = code;
    buf[h2 + k] = code;
  }
  return buf;
}

/** Register one bundled face's bytes (first-writer-wins, like the Node host). */
function installFont(file: string, bytes: Uint8Array): void {
  registerFont(file, bytes);
}

/** Result of {@link importDocx} — the opaque model plus everything export needs
 *  to faithfully re-emit it (image bytes keyed by their synthetic `cw-media:N`
 *  src, so a later export round-trips embedded images without a DOM). */
interface ImportHandle {
  doc: Document;
  warnings: { code: string; detail?: string }[];
  images: ImageBytes;
  mediaCount: number;
  blockCount: number;
}

function countBlocks(doc: Document): number {
  let n = 0;
  const walk = (blocks: Document["blocks"]): void => {
    for (const b of blocks) {
      n++;
      if (b.kind === "table") for (const r of b.rows) for (const c of r.cells) walk(c.blocks);
    }
  };
  walk(doc.blocks);
  return n;
}

/** Import a .docx (raw bytes) into the model. Always collects media bytes (V8
 *  has no URL.createObjectURL), so ImageBlocks carry `cw-media:N` srcs matched by
 *  the returned `images` map. */
function importDocx(bytes: Uint8Array): ImportHandle {
  const res = runImport(bytes, undefined, { collectMediaBytes: true });
  const images: ImageBytes = {};
  for (const m of res.media) images[m.src] = m.bytes;
  return {
    doc: res.doc as Document,
    warnings: res.warnings.map((w) => ({ code: w.code, ...(w.message ? { detail: w.message } : {}) })),
    images,
    mediaCount: res.media.length,
    blockCount: countBlocks(res.doc as Document),
  };
}

/** Decode any `data:*;base64,...` image sources in the document into the export image
 *  map, so builder-authored / data-URL images embed in PDF + DOCX (the export resolves
 *  image bytes by src). Non-base64 data URLs (e.g. raw SVG, which pdfkit can't embed)
 *  are skipped. Merges over caller-supplied images (which win). */
function withDataUrlImages(doc: Document, images?: ImageBytes): ImageBytes {
  const out: ImageBytes = { ...(images ?? {}) };
  const atob = (globalThis as unknown as { atob(s: string): string }).atob;
  const add = (src: string): void => {
    if (out[src] || !src.startsWith("data:")) return;
    const comma = src.indexOf(",");
    if (comma < 0 || !/;base64/i.test(src.slice(0, comma))) return;
    try {
      const bin = atob(src.slice(comma + 1));
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i) & 0xff;
      out[src] = bytes;
    } catch {
      /* leave unembedded → export draws a placeholder box */
    }
  };
  const walk = (blocks: Document["blocks"]): void => {
    for (const b of blocks) {
      if (b.kind === "image") add(b.src);
      else if (b.kind === "table") for (const r of b.rows) for (const c of r.cells) walk(c.blocks);
    }
  };
  walk(doc.blocks);
  const s = doc.section as unknown as Record<string, Document["blocks"] | undefined>;
  for (const band of ["header", "footer", "headerFirst", "footerFirst", "headerEven", "footerEven"]) {
    const story = s[band];
    if (story) walk(story);
  }
  return out;
}

async function exportPdf(
  doc: Document,
  images?: ImageBytes,
  fonts?: CustomFontPayload,
  cjk?: CjkExportConfig,
): Promise<Uint8Array> {
  const res = await runExport(doc, "pdf", withDataUrlImages(doc, images), fonts, cjk);
  return res.bytes;
}

async function exportDocx(doc: Document, images?: ImageBytes): Promise<Uint8Array> {
  const res = await runExport(doc, "docx", withDataUrlImages(doc, images));
  return res.bytes;
}

/** Generate a Word TOC field's result (entries + live PAGEREF page numbers) IN PLACE
 *  in a .docx that carries a `TOC` field — the layout-engine answer to "Word/OpenXML
 *  can't compute field values without paginating". Drift-free: the original document
 *  is patched + re-zipped, every other byte preserved. This is the Syncfusion
 *  replacement for headless TOC calculation. */
async function generateToc(
  docxBytes: Uint8Array,
  opts?: TocOptions,
): Promise<{ bytes: Uint8Array; generated: number; headings: number; bookmarksSynthesized: number }> {
  await installMeasureHost();
  const r = generateTocInDocx(docxBytes, opts ?? {});
  return { bytes: r.bytes, generated: r.generated, headings: r.headings, bookmarksSynthesized: r.bookmarksSynthesized };
}

/** Recalculate cached TOC/cross-reference (PAGEREF) page numbers IN PLACE (Word's
 *  F9 for page numbers): lay the document out, then rewrite each cached number to the
 *  heading's current page in the ORIGINAL document.xml. Drift-free. */
async function recalcTocPageNumbers(
  docxBytes: Uint8Array,
): Promise<{ bytes: Uint8Array; changed: number; skipped: number }> {
  await installMeasureHost();
  const { doc } = runImport(docxBytes, undefined, { collectMediaBytes: true });
  const r = patchTocFromLayout(docxBytes, doc as Document);
  return { bytes: r.bytes, changed: r.changed, skipped: r.skipped };
}

/** Update fields on an IN-MEMORY document handle (no docx round-trip) — Word's
 *  "Update Fields" for a document already loaded in WordCanvas. (1) Fill an empty TOC
 *  field from the CURRENT headings (so headings added after the TOC are included);
 *  (2) refresh any literal cached TOC page numbers via a layout pass. Returns a NEW
 *  document. Note: for builder/editor TOCs the entry page numbers are layout-resolved
 *  (paint-only) and are baked automatically by export — there is nothing stale to
 *  update; this mainly (re)materializes the entries from the headings. */
async function updateFields(doc: Document, opts?: TocOptions): Promise<Document> {
  await installMeasureHost();
  let next = generateTocIntoDoc(doc, opts ?? {}).doc;
  next = recalcToc(next).doc;
  return next;
}

/** Lay the document out and return the host page map (flattened). Async because
 *  it must ensure text measurement is installed, exactly like export. */
async function layoutPages(doc: Document): Promise<PageInfoDto[]> {
  await installMeasureHost();
  return mapPages(getPages(doc));
}

const api = {
  // Builder surface (the host's typed C# wrapper drives these objects directly).
  DocumentBuilder,
  units: { inches, cm, pt, twips, PAGE_SIZES },
  bytesToDataUrl,
  // Host font injection — call once after the engine loads, before any export.
  installFont,
  // Pipelines.
  importDocx,
  exportPdf,
  exportDocx,
  // TOC / field recalculation (layout-driven; the Syncfusion replacement).
  generateToc,
  recalcTocPageNumbers,
  updateFields,
  // Read-only query surface (the host's WordDocument query methods drive these).
  queryParagraphs,
  findText,
  querySections,
  querySdts,
  querySdtValue,
  queryFields,
  queryBookmarks,
  queryFootnotes,
  queryEndnotes,
  queryListItems,
  queryStyles,
  queryBlockPath,
  queryPositionOfText,
  queryRangeText,
  layoutPages,
  // Edit surface: open a stateful editor over a doc; the host's WordDocumentEditor
  // drives the returned instance's methods (setParagraphText, insertParagraph,
  // undo/redo, …) and reads its `.doc` back for export/query.
  openEditor: (doc: Document): DocumentEditor => new DocumentEditor(doc),
  // Construct a real JS RegExp from a host (C#) source+flags — lets the C#
  // ReplaceAllText expose the JS `pattern: string | RegExp` overload. The pattern
  // is TRUSTED developer input (a .NET binding caller, not an end user); a modest
  // length cap is cheap defense-in-depth against a runaway/catastrophic source.
  newRegExp: (source: string, flags: string): RegExp => {
    if (source.length > 1000) throw new Error("newRegExp: pattern too long (max 1000 chars)");
    return new RegExp(source, flags);
  },
  // Block counting helper (round-trip oracle for the smoke test / benchmark).
  countBlocks,
};

(globalThis as unknown as { WordCanvas: typeof api }).WordCanvas = api;

export default api;
export type { ImportHandle };
