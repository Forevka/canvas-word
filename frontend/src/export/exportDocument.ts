// Main-thread export API — the only module the app imports.
//
//   const { bytes, warnings } = await exportDocument(doc, "pdf");
//
// Runs the pipeline in a lazily-created singleton worker (terminated after 30s
// idle). Embedded image bytes are resolved HERE (blob:/data:/http URLs only
// resolve on the main thread) and handed to the worker, which never touches the
// network or the DOM.

import type { Block, Document } from "@cw/shared";
import type { ExportFormat, ExportResult, FromExportWorker, ImageBytes, ToExportWorker } from "./types";
import {
  CUSTOM_FONT_STYLES,
  faceUrlForStyle,
  type CustomFontFaceBytes,
  type CustomFontPayload,
  type ResolvedFontsConfig,
} from "../fonts/customRegistry";

export type { ExportFormat, ExportResult, ExportWarning } from "./types";

// Intentionally mirrors the same constant in importDocx.ts (both are worker-host files).
const IDLE_TERMINATE_MS = 30_000;

interface Pending {
  resolve: (r: ExportResult) => void;
  reject: (e: Error) => void;
}

let worker: Worker | null = null;
let idleTimer: ReturnType<typeof setTimeout> | undefined;
let nextId = 0;
const pending = new Map<number, Pending>();

function ensureWorker(): Worker {
  if (idleTimer !== undefined) {
    clearTimeout(idleTimer);
    idleTimer = undefined;
  }
  if (worker) return worker;
  worker = new Worker(new URL("./worker.ts", import.meta.url), { type: "module" });
  worker.onmessage = (e: MessageEvent<FromExportWorker>): void => {
    const msg = e.data;
    const job = pending.get(msg.id);
    if (!job) return;
    if (msg.type === "done") {
      pending.delete(msg.id);
      job.resolve({ bytes: msg.bytes, warnings: msg.warnings });
    } else if (msg.type === "error") {
      pending.delete(msg.id);
      job.reject(new Error(msg.message));
    }
    if (pending.size === 0) scheduleIdleTerminate();
  };
  worker.onerror = (e): void => {
    const err = new Error(`export worker failed: ${e.message}`);
    for (const job of pending.values()) job.reject(err);
    pending.clear();
    worker?.terminate();
    worker = null;
  };
  return worker;
}

function scheduleIdleTerminate(): void {
  if (idleTimer !== undefined) clearTimeout(idleTimer);
  idleTimer = setTimeout(() => {
    worker?.terminate();
    worker = null;
    idleTimer = undefined;
  }, IDLE_TERMINATE_MS);
}

/** Every image src referenced anywhere in the document (body, tables, bands,
 *  section breaks). Footnotes hold only paragraphs, so they carry no images. */
function collectImageSrcs(doc: Document): Set<string> {
  const out = new Set<string>();
  const walk = (blocks: Block[] | undefined): void => {
    if (!blocks) return;
    for (const b of blocks) {
      if (b.kind === "image") out.add(b.src);
      else if (b.kind === "table") for (const row of b.rows) for (const cell of row.cells) walk(cell.blocks);
      else if (b.kind === "paragraph" && b.style.sectionBreak) {
        const p = b.style.sectionBreak.props;
        walk(p.header); walk(p.footer); walk(p.headerFirst);
        walk(p.headerEven); walk(p.footerFirst); walk(p.footerEven);
      }
    }
  };
  walk(doc.blocks);
  const s = doc.section;
  walk(s.header); walk(s.footer); walk(s.headerFirst);
  walk(s.headerEven); walk(s.footerFirst); walk(s.footerEven);
  return out;
}

async function resolveImages(doc: Document): Promise<ImageBytes> {
  const srcs = [...collectImageSrcs(doc)];
  const out: ImageBytes = {};
  await Promise.all(
    srcs.map(async (src) => {
      try {
        const res = await fetch(src);
        out[src] = new Uint8Array(await res.arrayBuffer());
      } catch {
        // Leave unresolved — the writer/renderer warns and emits a placeholder.
      }
    }),
  );
  return out;
}

/** Fetch every custom face's bytes on the main thread (deduped by URL). Missing
 *  optional faces fall back to the regular URL. Returns undefined when there are no
 *  custom fonts. A family whose REQUIRED Regular face fails to fetch is dropped
 *  entirely — keeping it would register custom sizing/metrics with no bytes to
 *  embed, so the export would silently fall back to bundled glyphs and desync from
 *  the editor. A missing optional face just falls back to that family's Regular. */
async function resolveFonts(cfg: ResolvedFontsConfig | undefined): Promise<CustomFontPayload | undefined> {
  if (!cfg || cfg.fonts.length === 0) return undefined;
  const cache = new Map<string, Promise<Uint8Array | null>>();
  const fetchBytes = (url: string): Promise<Uint8Array | null> => {
    let p = cache.get(url);
    if (!p) {
      p = fetch(url)
        .then((r) => (r.ok ? r.arrayBuffer().then((b) => new Uint8Array(b)) : null))
        .catch(() => null);
      cache.set(url, p);
    }
    return p;
  };
  const faces: CustomFontFaceBytes[] = [];
  const regularOk = new Set<string>();
  await Promise.all(
    cfg.fonts.flatMap((f) =>
      CUSTOM_FONT_STYLES.map(async (style) => {
        const bytes = await fetchBytes(faceUrlForStyle(f.faces, style));
        if (!bytes) return;
        faces.push({ family: f.family, style, bytes });
        if (style === "Regular") regularOk.add(f.family);
      }),
    ),
  );
  const keptFonts = cfg.fonts.filter((f) => regularOk.has(f.family));
  if (keptFonts.length === 0) return undefined;
  const defs: ResolvedFontsConfig = { disableBuiltin: cfg.disableBuiltin, fonts: keptFonts };
  return { defs, faces: faces.filter((face) => regularOk.has(face.family)) };
}

/** Export a document. Pass the editor instance's resolved `fontConfig` so the worker
 *  embeds exactly that instance's custom faces — the config travels on the call, not
 *  a shared module slot, so two editors with different fonts never cross-export. */
export async function exportDocument(
  doc: Document,
  format: ExportFormat,
  fontConfig?: ResolvedFontsConfig,
): Promise<ExportResult> {
  const [images, fonts] = await Promise.all([resolveImages(doc), resolveFonts(fontConfig)]);
  const w = ensureWorker();
  const id = nextId++;
  return new Promise<ExportResult>((resolve, reject) => {
    pending.set(id, { resolve, reject });
    const msg: ToExportWorker = { id, doc, format, images, ...(fonts ? { fonts } : {}) };
    w.postMessage(msg);
  });
}

export const exportPdf = (doc: Document, fontConfig?: ResolvedFontsConfig): Promise<ExportResult> =>
  exportDocument(doc, "pdf", fontConfig);
export const exportDocx = (doc: Document, fontConfig?: ResolvedFontsConfig): Promise<ExportResult> =>
  exportDocument(doc, "docx", fontConfig);
