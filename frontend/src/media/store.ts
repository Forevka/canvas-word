// Session media store + resolver — the frontend side of content-addressed media.
//
// Image bytes are registered once under their content address (mediaId =
// sha256(bytes)); the model only ever carries the mediaId. At display time the
// resolver turns a mediaId back into a live blob: URL (cached for the session),
// so a document loaded from a snapshot — whose image srcs are blank — paints
// correctly. This is the inverse of shared/persist/serialize, which strips srcs.

import { forEachImage, MemoryMediaStore, mediaIdForBytes, type Document } from "@cw/shared";

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

// ---- retention / eviction ---------------------------------------------------
// The store is process-global (content addressing dedupes identical bytes across
// mounts), which historically meant it could only GROW: every image ever loaded
// kept its bytes AND a live blob: URL until page unload. Mounts now register a
// retention provider (the mediaIds their CURRENT document references, read live
// at prune time); pruning evicts everything else. Pruning runs only on the
// undo-safe transition — document replacement rebuilds the editor and drops its
// history, so an evicted id can't resurface via undo — and only when exactly ONE
// mount is registered: with several concurrent mounts, another mount's undo
// history may reference ids outside its current document, so the store keeps
// everything (the pre-eviction behavior). Callers that never register (tests,
// headless jobs) are unaffected.

type MediaIdProvider = () => Iterable<string>;
const retainProviders = new Set<MediaIdProvider>();

/** Every mediaId referenced by a document's images. */
export function mediaIdsOf(doc: Document): Set<string> {
  const ids = new Set<string>();
  forEachImage(doc, (b) => {
    if (b.mediaId) ids.add(b.mediaId);
  });
  return ids;
}

/** Register a mount's retention provider. Returns the unregister function (call
 *  on mount destroy). Unregistering does NOT evict — a destroyed mount's media
 *  may still back a remount rehydrating the same document (React strict-mode
 *  double mounts); the next mount's own prune reclaims it. */
export function registerMediaRetention(provider: MediaIdProvider): () => void {
  retainProviders.add(provider);
  return () => {
    retainProviders.delete(provider);
  };
}

/** Evict stored bytes and revoke blob: URLs for every mediaId no registered
 *  mount retains. Call on document replacement (see retention notes above). */
export function pruneOrphanMedia(): void {
  if (retainProviders.size !== 1) return;
  const keep = new Set<string>();
  for (const provider of retainProviders) for (const id of provider()) keep.add(id);
  for (const id of store.ids()) {
    if (!keep.has(id)) store.delete(id);
  }
  for (const [id, url] of [...urlCache]) {
    if (!keep.has(id)) {
      URL.revokeObjectURL(url);
      urlCache.delete(id);
    }
  }
}

/** Rehydrate every image in a freshly-deserialized document: fill blank srcs
 *  from the store via each image's mediaId. Mutates in place. Returns the ids
 *  it could NOT resolve (their bytes must be fetched before they'll paint). */
export function rehydrateDocMedia(doc: Document): string[] {
  const missing: string[] = [];
  forEachImage(doc, (b) => {
    // Only fill blanks (a snapshot's stripped src); leave live/data: srcs.
    if (b.src !== "") return;
    // Linked ("Link to File") image: no bytes/mediaId — its portable URL rehydrates
    // the src directly (the browser loads it like any other image source).
    if (b.externalSrc) {
      b.src = b.externalSrc;
      return;
    }
    if (!b.mediaId) return;
    const url = mediaUrl(b.mediaId);
    if (url) b.src = url;
    else if (!missing.includes(b.mediaId)) missing.push(b.mediaId);
  });
  return missing;
}
