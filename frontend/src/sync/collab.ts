// Online document operations: publish a local document to the backend (so it
// gets a server doc id + becomes shareable) and load/join an existing one. Both
// move content-addressed media over the wire so images survive across clients.
// Reuses @cw/shared serialization and the session media store.

import { collectMediaIds, serializeDocument, type Document } from "@cw/shared";
import { mediaStore, rehydrateDocMedia } from "../media/store";

/** Upload referenced media, then create a server document from the snapshot. */
export async function publishDocument(
  backendUrl: string,
  doc: Document,
): Promise<{ docId: string; version: number }> {
  const store = mediaStore();
  for (const id of collectMediaIds(doc)) {
    const blob = store.get(id);
    if (!blob) continue; // bytes not in this session (shouldn't happen for imported media)
    await fetch(`${backendUrl}/media/${encodeURIComponent(id)}`, {
      method: "PUT",
      headers: { "content-type": blob.mime },
      body: new Uint8Array(blob.bytes),
    });
  }
  const res = await fetch(`${backendUrl}/docs`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(serializeDocument(doc)),
  });
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
