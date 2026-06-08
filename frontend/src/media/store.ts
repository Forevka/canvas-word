// Session media store + resolver — the frontend side of content-addressed media.
//
// Image bytes are registered once under their content address (mediaId =
// sha256(bytes)); the model only ever carries the mediaId. At display time the
// resolver turns a mediaId back into a live blob: URL (cached for the session),
// so a document loaded from a snapshot — whose image srcs are blank — paints
// correctly. This is the inverse of shared/persist/serialize, which strips srcs.

import { BAND_CONTAINERS, MemoryMediaStore, mediaIdForBytes, type Block, type Document } from "@cw/shared";

const store = new MemoryMediaStore();
// mediaId -> blob: URL, live for this session. Avoids minting a second URL for
// bytes we already have one for.
const urlCache = new Map<string, string>();

/** The session media store (bytes by content address). */
export function mediaStore(): MemoryMediaStore {
  return store;
}

/** Hash + register bytes, returning their mediaId. Idempotent: identical bytes
 *  dedupe to one entry. */
export async function registerMediaBytes(bytes: Uint8Array, mime: string): Promise<string> {
  const mediaId = await mediaIdForBytes(bytes);
  if (!store.has(mediaId)) store.put({ mediaId, mime, bytes });
  return mediaId;
}

/** Associate an already-live blob URL with a mediaId (so the import path, which
 *  has just created a URL for these bytes, doesn't make a duplicate). */
export function bindMediaUrl(mediaId: string, url: string): void {
  if (!urlCache.has(mediaId)) urlCache.set(mediaId, url);
}

/** A live blob: URL for a mediaId, or undefined if the bytes aren't in the store
 *  (e.g. a snapshot referencing media that hasn't been fetched yet). */
export function mediaUrl(mediaId: string): string | undefined {
  const cached = urlCache.get(mediaId);
  if (cached) return cached;
  const blob = store.get(mediaId);
  if (!blob) return undefined;
  // Copy into a fresh ArrayBuffer-backed view (satisfies BlobPart's
  // ArrayBufferView<ArrayBuffer> and detaches from any shared backing buffer).
  const url = URL.createObjectURL(new Blob([new Uint8Array(blob.bytes)], { type: blob.mime }));
  urlCache.set(mediaId, url);
  return url;
}

function rewrite(blocks: Block[] | undefined): void {
  for (const b of blocks ?? []) {
    if (b.kind === "image") {
      // Only fill blanks (a snapshot's stripped src); leave live/data: srcs.
      if (b.mediaId && b.src === "") {
        const url = mediaUrl(b.mediaId);
        if (url) b.src = url;
      }
    } else if (b.kind === "table") {
      for (const row of b.rows) for (const cell of row.cells) rewrite(cell.blocks);
    }
  }
}

/** Rehydrate every image in a freshly-deserialized document: fill blank srcs
 *  from the store via each image's mediaId. Mutates in place. Returns the ids
 *  it could NOT resolve (their bytes must be fetched before they'll paint). */
export function rehydrateDocMedia(doc: Document): string[] {
  rewrite(doc.blocks);
  for (const band of BAND_CONTAINERS) rewrite(doc.section[band]);
  const missing: string[] = [];
  const check = (blocks: Block[] | undefined): void => {
    for (const b of blocks ?? []) {
      if (b.kind === "image") {
        if (b.mediaId && b.src === "" && !missing.includes(b.mediaId)) missing.push(b.mediaId);
      } else if (b.kind === "table") {
        for (const row of b.rows) for (const cell of row.cells) check(cell.blocks);
      }
    }
  };
  check(doc.blocks);
  for (const band of BAND_CONTAINERS) check(doc.section[band]);
  return missing;
}
