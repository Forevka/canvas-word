// Document <-> snapshot serialization.
//
// The model is already plain JSON-safe data, so serialization is mostly a deep
// clone. The one non-portable field is an image's `src` (a session-local blob:
// URL): it is dropped on serialize, leaving `mediaId` as the portable handle a
// resolver rehydrates on load (see ./media). A snapshot is the "base version" a
// server hands a client; folding the change log over it reconstructs any later
// version (see ../replay).

import { BAND_CONTAINERS } from "../model/document";
import type { Block, Document, ImageBlock } from "../model/document";

export const SNAPSHOT_VERSION = 1 as const;

export interface SerializedDocument {
  version: typeof SNAPSHOT_VERSION;
  doc: Document;
}

function walkImages(blocks: Block[], fn: (img: ImageBlock) => void): void {
  for (const b of blocks) {
    if (b.kind === "image") fn(b);
    else if (b.kind === "table") {
      for (const row of b.rows) for (const cell of row.cells) walkImages(cell.blocks, fn);
    }
  }
}

/** Visit every image in the document — body, margin bands, and (recursively)
 *  table cells. Footnotes hold only paragraphs, so they carry no images. */
export function forEachImage(doc: Document, fn: (img: ImageBlock) => void): void {
  walkImages(doc.blocks, fn);
  for (const band of BAND_CONTAINERS) {
    const arr = doc.section[band];
    if (arr) walkImages(arr, fn);
  }
}

/** Every distinct mediaId referenced by the document — the media a snapshot
 *  needs bundled alongside it for a faithful reload. */
export function collectMediaIds(doc: Document): Set<string> {
  const ids = new Set<string>();
  forEachImage(doc, (img) => {
    if (img.mediaId) ids.add(img.mediaId);
  });
  return ids;
}

// Document is pure JSON (no Date/Map/Set/typed arrays), so JSON round-trip is a
// correct, dependency-free deep clone — and it naturally drops `undefined`
// optional fields, which suits exactOptionalPropertyTypes.
const clone = <T>(v: T): T => JSON.parse(JSON.stringify(v)) as T;

const isPortableSrc = (src: string): boolean => src.startsWith("data:");

export function serializeDocument(doc: Document): SerializedDocument {
  const copy = clone(doc);
  forEachImage(copy, (img) => {
    // Blob URLs are session-local and unserializable — drop them. mediaId (kept)
    // is how the bytes are found again. Inline data: URLs are already portable,
    // so leave them as a fallback for images without a mediaId.
    if (!isPortableSrc(img.src)) img.src = "";
  });
  return { version: SNAPSHOT_VERSION, doc: copy };
}

export function deserializeDocument(snap: SerializedDocument): Document {
  // Future versions branch on snap.version here. The clone returns a doc whose
  // image `src` is "" (blob-derived) or an inline data: URL; a media resolver
  // fills the blanks from a MediaStore via each image's mediaId (see ./media).
  return clone(snap.doc);
}
