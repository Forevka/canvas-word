// Server-side .docx / .pdf export: reconstruct a document's head version and run
// the (pure, Node-safe) export pipeline. installMeasureHost() injects the DOM-free
// fontkit measurement context, so layout — and PDF rendering — work headless.
//
// Media bridge: the stored snapshot dropped image src (kept mediaId). The exporter
// keys embedded images by ImageBlock.src, so we set src = mediaId and feed an
// images map keyed by the same id, pulling bytes from the media table.

import { runExport } from "@forevka/wordcanvas/export";
import { installMeasureHost } from "@forevka/wordcanvas/export/measure";
import { forEachImage, reconstruct } from "@cw/shared";
import type { ChangeStore } from "../store/ChangeStore";

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
): Promise<ExportedDoc | null> {
  const snap = await store.getSnapshot(docId);
  if (!snap) return null;
  const changes = await store.getChanges(docId, snap.version);
  const doc = reconstruct(snap.snapshot, changes);

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
    const { bytes } = await runExport(doc, format, images);
    const title = await store.getDocumentTitle(docId);
    return { bytes, title };
  } finally {
    release();
  }
}
