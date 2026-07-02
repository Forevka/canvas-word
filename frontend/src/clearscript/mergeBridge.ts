// C#-facing merge bridge: merge two in-V8 documents AND union their media byte
// maps (keyed by ImageBlock.src). Kept separate from entry.ts so it is
// unit-testable without the V8 bundle. Pairs with WordDocument.Append in
// dotnet/src/WordCanvas.ClearScript/WordDocument.cs.
//
// The model merge (mergeDocuments) rebases every model id space but leaves
// ImageBlock.src untouched (it is a runtime export key, not model data). Two
// independently-imported documents can each carry `cw-media:0`, `cw-media:1`, …
// for DIFFERENT bytes, so unioning the maps naively would drop images. This
// bridge detects such collisions, reassigns the source's colliding keys, and
// rewrites the src on the merged source-origin image blocks.

import type { Block, Document } from "@cw/shared";
import { BAND_CONTAINERS, mergeDocuments, type MergeOptions, type MergeResult } from "@cw/shared";
import type { ImageBytes } from "../export/types";

export interface MergeBridgeResult {
  doc: Document;
  images: ImageBytes;
  mediaCount: number;
  warnings: MergeResult["warnings"];
}

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a === b) return true;
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/** Union two src-keyed media maps. A source key that collides with a destination
 *  key holding DIFFERENT bytes is reassigned a fresh key; identical bytes dedupe.
 *  Returns the combined map + the src rewrite (oldSrc → newSrc). */
function unionMedia(dest: ImageBytes, src: ImageBytes): { images: ImageBytes; remap: Record<string, string> } {
  const images: ImageBytes = { ...dest };
  const remap: Record<string, string> = {};
  const used = new Set(Object.keys(images));
  let n = 0;
  for (const [key, bytes] of Object.entries(src)) {
    const existing = images[key];
    if (existing !== undefined) {
      if (bytesEqual(existing, bytes)) continue; // identical image → dedupe
      let nk = `cw-media-m${n++}:${key}`;
      while (used.has(nk)) nk = `cw-media-m${n++}:${key}`;
      remap[key] = nk;
      images[nk] = bytes;
      used.add(nk);
    } else {
      images[key] = bytes;
      used.add(key);
    }
  }
  return { images, remap };
}

/** Visit every image block reachable from `blocks`, recursing table cells and
 *  section-break band stories. */
function visitImages(blocks: readonly Block[], fn: (img: Extract<Block, { kind: "image" }>) => void): void {
  for (const b of blocks) {
    if (b.kind === "image") fn(b);
    else if (b.kind === "table") for (const row of b.rows) for (const cell of row.cells) visitImages(cell.blocks, fn);
    else if (b.kind === "paragraph" && b.style.sectionBreak) {
      for (const key of BAND_CONTAINERS) {
        const band = b.style.sectionBreak.props[key];
        if (band) visitImages(band, fn);
      }
    }
  }
}

/** Rewrite `src` on merged SOURCE-origin image blocks per `remap`. Source-origin
 *  blocks are those whose id is a value in the merge id map (freshly cloned by the
 *  merge), so destination images sharing a src key are never touched. */
function rewriteSourceImageSrcs(doc: Document, sourceBlockIds: Set<string>, remap: Record<string, string>): void {
  if (Object.keys(remap).length === 0) return;
  const fix = (img: Extract<Block, { kind: "image" }>): void => {
    if (sourceBlockIds.has(img.id) && remap[img.src]) img.src = remap[img.src]!;
  };
  visitImages(doc.blocks, fix);
  for (const key of BAND_CONTAINERS) {
    const band = doc.section[key];
    if (band) visitImages(band, fix);
  }
  for (const paras of Object.values(doc.footnotes ?? {})) visitImages(paras, fix);
  for (const paras of Object.values(doc.endnotes ?? {})) visitImages(paras, fix);
}

/** Merge `src` into `dest` (model + media). `destImages`/`srcImages` may be
 *  undefined (builder output). */
export function mergeBridge(
  destDoc: Document,
  destImages: ImageBytes | undefined,
  srcDoc: Document,
  srcImages: ImageBytes | undefined,
  opts?: MergeOptions,
): MergeBridgeResult {
  const { images, remap } = unionMedia(destImages ?? {}, srcImages ?? {});
  const result = mergeDocuments(destDoc, srcDoc, opts);
  const sourceBlockIds = new Set(Object.values(result.idMap.blocks));
  rewriteSourceImageSrcs(result.doc, sourceBlockIds, remap);
  return { doc: result.doc, images, mediaCount: Object.keys(images).length, warnings: result.warnings };
}
