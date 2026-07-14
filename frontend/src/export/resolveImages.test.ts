import { afterEach, describe, expect, it, vi } from "vitest";
import type { Document, ImageBlock, SectionProps } from "@cw/shared";
import { resolveImages } from "./exportDocument";

const SECTION: SectionProps = {
  pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
};
const img = (id: string, src: string): ImageBlock => ({ kind: "image", id, revision: 0, src, widthPx: 10, heightPx: 10, align: "left" });
const doc = (srcs: string[]): Document => ({ section: { ...SECTION }, blocks: srcs.map((s, i) => img(`i${i}`, s)) });

afterEach(() => vi.unstubAllGlobals());

describe("resolveImages — external image CORS reporting", () => {
  it("flags an external URL whose fetch rejects (CORS/network) as failedExternal", async () => {
    vi.stubGlobal("fetch", vi.fn().mockRejectedValue(new TypeError("Failed to fetch")));
    const { images, failedExternal } = await resolveImages(doc(["https://cdn.example.com/a.jpg"]));
    expect(failedExternal).toEqual(["https://cdn.example.com/a.jpg"]);
    expect(images["https://cdn.example.com/a.jpg"]).toBeUndefined();
  });

  it("flags an external URL returning a non-OK status", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({ ok: false, status: 403 }));
    const { failedExternal } = await resolveImages(doc(["https://cdn.example.com/forbidden.png"]));
    expect(failedExternal).toEqual(["https://cdn.example.com/forbidden.png"]);
  });

  it("embeds an external URL that fetches successfully (no warning)", async () => {
    const bytes = new Uint8Array([1, 2, 3]).buffer;
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({ ok: true, arrayBuffer: async () => bytes }));
    const { images, failedExternal } = await resolveImages(doc(["https://cdn.example.com/ok.png"]));
    expect(failedExternal).toHaveLength(0);
    expect(images["https://cdn.example.com/ok.png"]).toEqual(new Uint8Array([1, 2, 3]));
  });

  it("does not flag non-external (blob:/data:) srcs that fail — only http(s)", async () => {
    vi.stubGlobal("fetch", vi.fn().mockRejectedValue(new Error("gone")));
    const { failedExternal } = await resolveImages(doc(["blob:https://app/abc", "data:image/png;base64,AAAA"]));
    expect(failedExternal).toHaveLength(0);
  });
});
