// Server-side .docx import: parse an uploaded file into the document model on the
// backend (the import pipeline is pure / Node-safe) and persist it as a new
// document, so a caller can immediately open WordCanvas({ docId }) for review.
//
// Media bridge: runImport with { collectMediaBytes } stamps each ImageBlock with a
// synthetic src ("cw-media:N") and hands back the raw bytes. We content-address
// the bytes into the media table (same sha256 the client uses) and swap each
// block's src for its mediaId — the portable handle serializeDocument keeps.

import { runImport } from "@cw/frontend/import";
import { forEachImage, mediaIdForBytes, serializeDocument, type UserInfo } from "@cw/shared";
import type { ChangeStore } from "../store/ChangeStore";

export interface UploadResult {
  docId: string;
  version: number;
  warnings: { code: string; message: string }[];
}

export async function parseAndStore(
  bytes: Uint8Array,
  filename: string | undefined,
  store: ChangeStore,
  user?: UserInfo,
): Promise<UploadResult> {
  const result = runImport(bytes, undefined, { collectMediaBytes: true });

  // Store each embedded image once (content-addressed) and remember src -> hash.
  const srcToMediaId = new Map<string, string>();
  for (const m of result.media) {
    const hash = await mediaIdForBytes(m.bytes);
    await store.putMedia({ hash, mime: m.mime, bytes: m.bytes });
    srcToMediaId.set(m.src, hash);
  }
  // Rewrite blocks: synthetic src -> mediaId; serialize then drops the src.
  forEachImage(result.doc, (img) => {
    const mediaId = srcToMediaId.get(img.src);
    if (mediaId) {
      img.mediaId = mediaId;
      img.src = "";
    }
  });

  if (user) await store.upsertUser(user);
  const snapshot = serializeDocument(result.doc);
  const title = filename ? stripExt(filename) : undefined;
  const created = await store.createDocument(snapshot, {
    ...(user?.id ? { createdBy: user.id } : {}),
    ...(title ? { title } : {}),
  });
  return { docId: created.docId, version: created.version, warnings: result.warnings };
}

/** "report.docx" -> "report"; tolerant of paths and missing extensions. */
function stripExt(name: string): string {
  const slash = Math.max(name.lastIndexOf("/"), name.lastIndexOf("\\"));
  const base = slash >= 0 ? name.slice(slash + 1) : name;
  const dot = base.lastIndexOf(".");
  return dot > 0 ? base.slice(0, dot) : base;
}
