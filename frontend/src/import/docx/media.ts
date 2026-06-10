// Embedded media → blob: URLs. Created where the pipeline runs (worker or
// Node — both have Blob/createObjectURL); the URLs are plain strings that
// survive structured clone and are valid on the main thread (same origin).
//
// Each part (document, header1, …) has its OWN rels, so media stores are
// created per part; the pipeline aggregates the URLs for ImportResult.mediaUrls.

import type { Relationships } from "./relationships";
import { WarningSink, type MediaCollector } from "./types";
import type { Archive } from "./zip";

export interface MediaStore {
  /** blob: URL for an image relationship, or undefined if missing/unsupported. */
  resolve(relId: string): string | undefined;
  /** Every URL this store created — for ImportResult.mediaUrls (caller revokes). */
  urls(): string[];
}

/** Browsers can't decode Windows metafiles (Word loves embedding them) — skip
 *  with a warning rather than producing a broken <img>-equivalent. */
const CONTENT_TYPES: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  webp: "image/webp",
  svg: "image/svg+xml",
};

export function createMediaStore(
  archive: Archive,
  rels: Relationships,
  warnings: WarningSink,
  collect?: MediaCollector,
): MediaStore {
  const byTarget = new Map<string, string | undefined>(); // same image referenced twice → one URL
  const created: string[] = [];

  return {
    resolve(relId) {
      const rel = rels.get(relId);
      if (!rel) {
        warnings.add("media-missing", "An image reference had no matching relationship — skipped.");
        return undefined;
      }
      if (byTarget.has(rel.target)) return byTarget.get(rel.target);

      // Linked ("Link to File") images: the target IS the source. Browsers can
      // load http(s) directly; file:/local paths are unreachable from a page.
      if (rel.external) {
        const loadable = /^https?:\/\//i.test(rel.target);
        if (!loadable) {
          warnings.add(
            "media-external-unreachable",
            "Linked images pointing at local files can't be loaded in the browser — skipped.",
          );
        } else {
          warnings.add("media-external", "Linked (external) images load from their original URL.");
        }
        const src = loadable ? rel.target : undefined;
        byTarget.set(rel.target, src);
        return src;
      }

      let url: string | undefined;
      const ext = rel.target.slice(rel.target.lastIndexOf(".") + 1).toLowerCase();
      const contentType = CONTENT_TYPES[ext];
      if (!contentType) {
        warnings.add(
          "media-unsupported",
          `Some images use formats browsers can't display (.${ext}, typically metafiles) — skipped.`,
        );
      } else {
        const bytes = archive.part(rel.target);
        if (!bytes) {
          warnings.add("media-missing", "An image reference had no matching relationship — skipped.");
        } else {
          // Copy into a fresh buffer: `bytes` may be a view into the archive's
          // big backing buffer, and Blob([view]) would retain all of it.
          const copy = new Uint8Array(bytes);
          if (collect) {
            // Node/backend path: hand the bytes to the collector (which the
            // caller content-addresses) and use the synthetic src it returns —
            // no blob: URL (createObjectURL isn't reliably available off-browser).
            url = collect.add(copy, contentType);
          } else {
            url = URL.createObjectURL(new Blob([copy], { type: contentType }));
            created.push(url);
          }
        }
      }
      byTarget.set(rel.target, url);
      return url;
    },
    urls: () => [...created],
  };
}

/** No-op store for parts without rels. */
export const EMPTY_MEDIA: MediaStore = { resolve: () => undefined, urls: () => [] };
