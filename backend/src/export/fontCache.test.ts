// fontCache: fetch once, serve from disk thereafter, dedupe concurrent fetches, and
// surface a typed error on a bad URL. Uses a throwaway cache dir + a mocked global
// fetch (no network). FONT_CACHE_DIR is read at module load, so set it before the
// dynamic import.

import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";

const CACHE = mkdtempSync(join(tmpdir(), "wc-fontcache-"));
process.env.FONT_CACHE_DIR = CACHE;

// eslint-disable-next-line @typescript-eslint/consistent-type-imports
let mod: typeof import("./fontCache");

beforeAll(async () => {
  mod = await import("./fontCache");
});

afterAll(() => {
  rmSync(CACHE, { recursive: true, force: true });
  vi.restoreAllMocks();
});

function mockFetch(impl: (url: string) => Response): void {
  vi.stubGlobal(
    "fetch",
    vi.fn((url: string) => Promise.resolve(impl(url))),
  );
}

function okResponse(bytes: Uint8Array): Response {
  return {
    ok: true,
    status: 200,
    headers: new Headers(),
    arrayBuffer: async () => bytes.buffer,
  } as unknown as Response;
}

describe("fontCache", () => {
  it("fetches once, then serves the cached file", async () => {
    const bytes = new Uint8Array([1, 2, 3, 4]);
    const fetchSpy = vi.fn(() => Promise.resolve(okResponse(bytes)));
    vi.stubGlobal("fetch", fetchSpy);

    const url = "https://x/once.ttf";
    const a = await mod.fetchAndCache(url);
    const b = await mod.fetchAndCache(url);
    expect([...a]).toEqual([1, 2, 3, 4]);
    expect([...b]).toEqual([1, 2, 3, 4]);
    expect(fetchSpy).toHaveBeenCalledTimes(1); // second call hit disk
  });

  it("dedupes concurrent fetches of the same URL", async () => {
    const fetchSpy = vi.fn(() => Promise.resolve(okResponse(new Uint8Array([9]))));
    vi.stubGlobal("fetch", fetchSpy);
    const url = "https://x/concurrent.ttf";
    await Promise.all([mod.fetchAndCache(url), mod.fetchAndCache(url), mod.fetchAndCache(url)]);
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  });

  it("throws FontFetchError on a non-OK response", async () => {
    mockFetch(() => ({ ok: false, status: 404 }) as unknown as Response);
    await expect(mod.fetchAndCache("https://x/missing.ttf")).rejects.toBeInstanceOf(mod.FontFetchError);
  });

  it("resolveCustomFontBytes fetches one face per style (regular fallback)", async () => {
    const fetchSpy = vi.fn(() => Promise.resolve(okResponse(new Uint8Array([7]))));
    vi.stubGlobal("fetch", fetchSpy);
    const payload = await mod.resolveCustomFontBytes({
      disableBuiltin: [],
      fonts: [{ family: "Inter", faces: { regular: "https://x/Inter.ttf" }, sizing: { ascent: 0.9, descent: 0.2 } }],
    });
    expect(payload?.faces.map((f) => f.style).sort()).toEqual(["Bold", "BoldItalic", "Italic", "Regular"]);
    // All four styles fall back to the single regular URL → one network fetch.
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  });

  it("returns undefined when there are no custom fonts", async () => {
    expect(await mod.resolveCustomFontBytes(undefined)).toBeUndefined();
    expect(await mod.resolveCustomFontBytes({ disableBuiltin: [], fonts: [] })).toBeUndefined();
  });
});
