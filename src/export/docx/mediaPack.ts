// Collects embedded image bytes into word/media/* parts, deduped by source.
// Image bytes are resolved by the caller (blob: URLs are main-thread only); the
// model's ImageBlock.src is the lookup key.

import type { ImageBytes } from "../types";

interface MediaEntry {
  path: string; // e.g. "media/image1.png"
  bytes: Uint8Array;
  ext: string;
}

function sniff(bytes: Uint8Array): { ext: string; contentType: string } {
  if (bytes.length >= 4 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47)
    return { ext: "png", contentType: "image/png" };
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xd8) return { ext: "jpeg", contentType: "image/jpeg" };
  if (bytes.length >= 3 && bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46)
    return { ext: "gif", contentType: "image/gif" };
  return { ext: "png", contentType: "image/png" }; // best-effort default
}

export class MediaManager {
  private readonly bySrc = new Map<string, MediaEntry>();
  private readonly exts = new Map<string, string>(); // ext -> contentType
  private n = 0;

  constructor(private readonly images: ImageBytes) {}

  /** Register an image src; returns its part path ("media/imageN.ext") or null
   *  when bytes are unavailable. */
  resolve(src: string): string | null {
    const hit = this.bySrc.get(src);
    if (hit) return hit.path;
    const bytes = this.images[src];
    if (!bytes) return null;
    const { ext, contentType } = sniff(bytes);
    const path = `media/image${++this.n}.${ext}`;
    this.bySrc.set(src, { path, bytes, ext });
    this.exts.set(ext, contentType);
    return path;
  }

  /** Parts to add under word/ : [["media/image1.png", bytes], ...]. */
  parts(): [string, Uint8Array][] {
    return [...this.bySrc.values()].map((e) => [e.path, e.bytes]);
  }

  /** Default content-type overrides by extension for [Content_Types].xml. */
  extensions(): [string, string][] {
    return [...this.exts.entries()];
  }
}
