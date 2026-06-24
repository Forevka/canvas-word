// Server-side custom-font fetcher with an on-disk cache. The export pipeline needs
// each custom face's BYTES (fontkit measures + pdfkit embeds them); a document /
// render request only carries URLs. We fetch each URL once, cache it on disk keyed
// by the URL's sha256 (URLs are treated as immutable content addresses), and reuse
// it across every render — so a burst of exports doesn't re-download the same fonts.
//
// The face-key/style fallback mirrors the frontend (customRegistry.ts): a missing
// bold/italic/boldItalic uses the regular URL, and the registry keys faces by
// (family, style), so resolveFont in the package finds them.

import { createHash } from "node:crypto";
import { mkdir, readFile, rename, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type {
  CustomFontFaceBytes,
  CustomFontPayload,
  ResolvedFontsConfig,
} from "@forevka/wordcanvas/export";

type FontStyleName = CustomFontFaceBytes["style"];
const STYLES: FontStyleName[] = ["Regular", "Bold", "Italic", "BoldItalic"];

interface FaceUrls {
  regular: string;
  bold?: string;
  italic?: string;
  boldItalic?: string;
}

function faceUrlForStyle(faces: FaceUrls, style: FontStyleName): string {
  switch (style) {
    case "Bold":
      return faces.bold ?? faces.regular;
    case "Italic":
      return faces.italic ?? faces.regular;
    case "BoldItalic":
      return faces.boldItalic ?? faces.regular;
    default:
      return faces.regular;
  }
}

/** Host of a URL, for log/error messages — never the full URL (presigned tokens,
 *  internal paths). Returns a placeholder when the URL won't parse. */
function hostOf(url: string): string {
  try {
    return new URL(url).host;
  } catch {
    return "<invalid-url>";
  }
}

/** Raised when a font URL can't be fetched (404 / network / non-OK). The message
 *  carries only the host (not the raw URL) so logs / 400 bodies don't leak tokens
 *  or internal hostnames; the full `url` stays on the instance for internal use. */
export class FontFetchError extends Error {
  constructor(public readonly url: string, message: string) {
    super(`font fetch failed (${hostOf(url)}): ${message}`);
    this.name = "FontFetchError";
  }
}

const CACHE_DIR = process.env.FONT_CACHE_DIR ?? join(tmpdir(), "wordcanvas-fonts");
// Dedupe concurrent fetches of the same URL within a process.
const inflight = new Map<string, Promise<Uint8Array>>();

function cachePath(url: string): string {
  return join(CACHE_DIR, `${createHash("sha256").update(url).digest("hex")}.bin`);
}

/** Fetch a font URL through the disk cache. Reads the cached file if present, else
 *  downloads, writes atomically (temp + rename), and returns the bytes. */
export async function fetchAndCache(url: string): Promise<Uint8Array> {
  const path = cachePath(url);
  try {
    return new Uint8Array(await readFile(path));
  } catch {
    // not cached — fall through to fetch
  }
  let p = inflight.get(url);
  if (!p) {
    p = (async () => {
      // Validate scheme: request-supplied URLs must be plain http(s) (no file:,
      // data:, gopher:, etc. — SSRF surface).
      let parsed: URL;
      try {
        parsed = new URL(url);
      } catch {
        throw new FontFetchError(url, "invalid URL");
      }
      if (parsed.protocol !== "http:" && parsed.protocol !== "https:") {
        throw new FontFetchError(url, "unsupported URL scheme");
      }
      // Optional host allowlist (comma-separated FONT_FETCH_ALLOWED_HOSTS).
      const allowed = (process.env.FONT_FETCH_ALLOWED_HOSTS ?? "")
        .split(",")
        .map((h) => h.trim())
        .filter(Boolean);
      if (allowed.length > 0 && !allowed.includes(parsed.hostname)) {
        throw new FontFetchError(url, "host not allowed");
      }

      const MAX = Number(process.env.FONT_FETCH_MAX_BYTES ?? 10 * 1024 * 1024);
      const controller = new AbortController();
      const timeout = setTimeout(
        () => controller.abort(),
        Number(process.env.FONT_FETCH_TIMEOUT_MS ?? 10000),
      );
      let res: Response;
      try {
        res = await fetch(url, { signal: controller.signal });
      } catch (e) {
        if (e instanceof Error && e.name === "AbortError") throw new FontFetchError(url, "timeout");
        throw new FontFetchError(url, e instanceof Error ? e.message : String(e));
      } finally {
        clearTimeout(timeout);
      }
      if (!res.ok) throw new FontFetchError(url, `HTTP ${res.status}`);
      // Reject oversized fonts early (advertised length), then after download.
      const declared = res.headers.get("content-length");
      if (declared && Number(declared) > MAX) throw new FontFetchError(url, "font too large");
      const bytes = new Uint8Array(await res.arrayBuffer());
      if (bytes.byteLength > MAX) throw new FontFetchError(url, "font too large");
      await mkdir(CACHE_DIR, { recursive: true });
      const tmp = `${path}.${process.pid}.${bytes.byteLength}.tmp`;
      await writeFile(tmp, bytes);
      await rename(tmp, path);
      return bytes;
    })().finally(() => inflight.delete(url));
    inflight.set(url, p);
  }
  return p;
}

/** Resolve a fonts config into a pipeline payload: fetch (cached) each face's bytes
 *  for every style, keyed by (family, style). A face that 404s is dropped (the
 *  pipeline's resolveFont then falls back), so one bad URL doesn't fail the export. */
export async function resolveCustomFontBytes(
  defs: ResolvedFontsConfig | undefined,
): Promise<CustomFontPayload | undefined> {
  if (!defs || defs.fonts.length === 0) return undefined;
  const faces: CustomFontFaceBytes[] = [];
  await Promise.all(
    defs.fonts.flatMap((f) =>
      STYLES.map(async (style) => {
        try {
          const bytes = await fetchAndCache(faceUrlForStyle(f.faces, style));
          faces.push({ family: f.family, style, bytes });
        } catch (e) {
          console.warn(`[wordcanvas] custom font "${f.family}" ${style}: ${e instanceof Error ? e.message : e}`);
        }
      }),
    ),
  );
  return { defs, faces };
}
