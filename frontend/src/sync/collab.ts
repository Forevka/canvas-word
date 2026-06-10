// Online document operations: publish a local document to the backend (so it
// gets a server doc id + becomes shareable) and load/join an existing one. Both
// move content-addressed media over the wire so images survive across clients.
// Reuses @cw/shared serialization and the session media store.

import { collectMediaIds, serializeDocument, type Document, type UserInfo } from "@cw/shared";
import { mediaStore, rehydrateDocMedia } from "../media/store";

/** Max media uploads in flight at once. They share one HTTP/2 connection and
 *  each costs a roughly fixed server round-trip (content-hash + idempotent
 *  INSERT) independent of payload size — so the upload is latency-bound, and a
 *  small pool collapses N serial round-trips into a few parallel waves. */
const MEDIA_UPLOAD_CONCURRENCY = 6;

/** gzip a UTF-8 string via the platform CompressionStream (no JS deflate lib).
 *  The document snapshot is large, repetitive JSON that shrinks ~8×, turning the
 *  POST /docs upload from hundreds of KB into tens. */
async function gzipText(text: string): Promise<ArrayBuffer> {
  const stream = new Blob([text]).stream().pipeThrough(new CompressionStream("gzip"));
  return new Response(stream).arrayBuffer();
}

/** Publish a local document: upload its media and create the server document
 *  from the snapshot, attributed to `user` (the creator). The snapshot POST is
 *  independent of media (the server stores media separately, content-addressed
 *  and idempotent), so media uploads and the snapshot POST all run concurrently
 *  rather than serially — turning ~(N×RTT + post) into ~max(a few waves, post). */
export async function publishDocument(
  backendUrl: string,
  doc: Document,
  user?: UserInfo,
): Promise<{ docId: string; version: number }> {
  const store = mediaStore();
  const blobs = [...collectMediaIds(doc)]
    .map((id) => ({ id, blob: store.get(id) }))
    .filter((e): e is { id: string; blob: NonNullable<typeof e.blob> } => e.blob != null);

  // A worker pulls from the shared queue until it's drained; `pool` of them run
  // in parallel (bounded so we don't buffer every image body in memory at once).
  let next = 0;
  const uploadWorker = async (): Promise<void> => {
    while (next < blobs.length) {
      const { id, blob } = blobs[next++]!;
      await fetch(`${backendUrl}/media/${encodeURIComponent(id)}`, {
        method: "PUT",
        headers: { "content-type": blob.mime },
        body: new Uint8Array(blob.bytes),
      });
    }
  };
  const pool = Array.from(
    { length: Math.min(MEDIA_UPLOAD_CONCURRENCY, blobs.length) },
    uploadWorker,
  );

  const snapshotJson = JSON.stringify({ snapshot: serializeDocument(doc), createdBy: user?.id, user });
  const postSnapshot = gzipText(snapshotJson).then((body) =>
    fetch(`${backendUrl}/docs`, {
      method: "POST",
      headers: { "content-type": "application/json", "content-encoding": "gzip" },
      body,
    }),
  );

  const [res] = await Promise.all([postSnapshot, ...pool]);
  if (!res.ok) throw new Error(`publish failed (${res.status})`);
  return (await res.json()) as { docId: string; version: number };
}

/** Load the server's current document (snapshot + replayed log) and fetch any
 *  media this client doesn't already have, so images render for a joiner. */
export async function loadCollabDocument(
  backendUrl: string,
  docId: string,
): Promise<{ doc: Document; version: number }> {
  const res = await fetch(`${backendUrl}/docs/${encodeURIComponent(docId)}/document`);
  if (!res.ok) throw new Error(`collab document ${docId} not found (${res.status})`);
  const payload = (await res.json()) as { doc: Document; version: number };

  const missing = rehydrateDocMedia(payload.doc);
  if (missing.length > 0) {
    const store = mediaStore();
    await Promise.all(
      missing.map(async (id) => {
        try {
          const mr = await fetch(`${backendUrl}/media/${encodeURIComponent(id)}`);
          if (!mr.ok) return;
          const mime = mr.headers.get("content-type") ?? "application/octet-stream";
          store.put({ mediaId: id, mime, bytes: new Uint8Array(await mr.arrayBuffer()) });
        } catch {
          // leave unresolved; the image stays blank rather than breaking load
        }
      }),
    );
    rehydrateDocMedia(payload.doc); // fill the blanks we just fetched
  }
  return { doc: payload.doc, version: payload.version };
}
