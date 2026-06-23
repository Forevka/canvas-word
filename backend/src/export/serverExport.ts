// Server-side .docx / .pdf export: reconstruct a document's head version and run
// the (pure, Node-safe) export pipeline. installMeasureHost() injects the DOM-free
// fontkit measurement context, so layout — and PDF rendering — work headless.
//
// Media bridge: the stored snapshot dropped image src (kept mediaId). The exporter
// keys embedded images by ImageBlock.src, so we set src = mediaId and feed an
// images map keyed by the same id, pulling bytes from the media table.

import { runExport, type ResolvedFontsConfig } from "@forevka/wordcanvas/export";
import { installMeasureHost } from "@forevka/wordcanvas/export/measure";
import { runImport } from "@forevka/wordcanvas/import";
import { applyReviewOp, bakeReview, emptyReview, forEachImage, generateTocIntoDoc, reconstruct, type BakeMode, type TocOptions } from "@cw/shared";
import type { ChangeStore } from "../store/ChangeStore";
import { resolveCustomFontBytes } from "./fontCache";

export type ExportFormat = "docx" | "pdf";

export interface ExportedDoc {
  bytes: Uint8Array;
  title: string | null;
}

// A burst of PDF exports is CPU-heavy (full layout per call). Cap concurrency so
// one batch can't starve the event loop / collab traffic.
const MAX_CONCURRENT = Number(process.env.EXPORT_CONCURRENCY ?? 2);
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

/** Reconstruct + export. Returns null if the document doesn't exist. */
export async function exportDoc(
  docId: string,
  format: ExportFormat,
  store: ChangeStore,
  bakeMode: BakeMode = "reject",
): Promise<ExportedDoc | null> {
  const snap = await store.getSnapshot(docId);
  if (!snap) return null;
  // The document's saved custom-font config (set at create time) — fetched + cached
  // to bytes so server-side layout/embedding matches the editor.
  const fonts = await resolveCustomFontBytes((await store.getFontsConfig(docId)) ?? undefined);
  const changes = await store.getChanges(docId, snap.version);
  const reconstructed = reconstruct(snap.snapshot, changes);

  // Fold the review overlay into a clean document so the review-blind writer emits
  // a plain file. Default "reject" = the original baseline (insertions removed,
  // deletions kept); "accept" = the final text. The core writer never sees review.
  const reviewOps = await store.getReviewOps(docId);
  const review = reviewOps.reduce((l, env) => applyReviewOp(l, env.op).layer, emptyReview(docId));
  const doc = bakeReview(reconstructed, review, bakeMode);

  // Hydrate media: src = mediaId (stable key), bytes pulled from the store.
  const mediaIds = new Set<string>();
  forEachImage(doc, (img) => {
    if (img.mediaId) {
      img.src = img.mediaId;
      mediaIds.add(img.mediaId);
    }
  });
  const images: Record<string, Uint8Array> = {};
  await Promise.all(
    [...mediaIds].map(async (id) => {
      const rec = await store.getMedia(id);
      if (rec) images[id] = rec.bytes;
    }),
  );

  await acquire();
  try {
    await installMeasureHost();
    const { bytes } = await runExport(doc, format, images, fonts);
    const title = await store.getDocumentTitle(docId);
    return { bytes, title };
  } finally {
    release();
  }
}

/** Stateless headless render: a raw .docx in, a PDF out — the "calculate layout +
 *  fields" path for a renderer-less producer (e.g. a C# pipeline). We import the
 *  bytes (collecting embedded media so images render), BUILD the TOC field's entries
 *  from the document's headings styled by `tocOpts` (page numbers resolve during
 *  layout), then render. Footer PAGE/NUMPAGES are substituted per page by the layout
 *  engine, so "Page X of Y" is correct on every page. Nothing is stored. */
export async function renderPdfFromDocx(
  bytes: Uint8Array,
  tocOpts: TocOptions = {},
  fontsConfig?: ResolvedFontsConfig,
): Promise<Uint8Array> {
  const { doc, media } = runImport(bytes, undefined, { collectMediaBytes: true });
  // Build the TOC entries in place (no-op when the document has no TOC field).
  const rendered = doc.tocInstruction !== undefined ? generateTocIntoDoc(doc, tocOpts).doc : doc;
  const images: Record<string, Uint8Array> = {};
  for (const m of media) images[m.src] = m.bytes;
  const fonts = await resolveCustomFontBytes(fontsConfig);

  await acquire();
  try {
    await installMeasureHost();
    const { bytes: out } = await runExport(rendered, "pdf", images, fonts);
    return out;
  } finally {
    release();
  }
}
