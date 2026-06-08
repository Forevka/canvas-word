import { beforeEach, describe, expect, it, vi } from "vitest";
import type { Document, ImageBlock, SectionProps } from "@cw/shared";
import { mediaIdForBytes } from "@cw/shared";
import { mediaStore, mediaUrl, registerMediaBytes, rehydrateDocMedia } from "./store";

// The frontend test env is Node (no URL.createObjectURL) — stub it.
let urlSeq = 0;
beforeEach(() => {
  urlSeq = 0;
  (URL as unknown as { createObjectURL: (b: unknown) => string }).createObjectURL = vi.fn(
    () => `blob:test/${urlSeq++}`,
  );
});

const section = (): SectionProps => ({
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
});

const img = (id: string, mediaId: string | undefined, src: string): ImageBlock => ({
  kind: "image",
  id,
  revision: 0,
  src,
  ...(mediaId ? { mediaId } : {}),
  widthPx: 10,
  heightPx: 10,
  align: "left",
});

describe("media store + resolver", () => {
  it("registers bytes under their content address and dedupes", async () => {
    const bytes = new TextEncoder().encode("PNGDATA");
    const id = await registerMediaBytes(bytes, "image/png");
    expect(id).toBe(await mediaIdForBytes(bytes));
    // same bytes again → same id, still one entry
    await registerMediaBytes(new TextEncoder().encode("PNGDATA"), "image/png");
    expect(mediaStore().ids()).toContain(id);
  });

  it("resolves a mediaId to a cached blob URL", async () => {
    const id = await registerMediaBytes(new TextEncoder().encode("JPEG"), "image/jpeg");
    const a = mediaUrl(id);
    const b = mediaUrl(id);
    expect(a).toBeDefined();
    expect(a).toBe(b); // cached — one URL per session
    expect(mediaUrl("deadbeef")).toBeUndefined(); // unknown id
  });

  it("rehydrates a deserialized doc's blank image srcs from the store", async () => {
    const id = await registerMediaBytes(new TextEncoder().encode("GIF89a"), "image/gif");
    const doc: Document = {
      section: section(),
      blocks: [img("known", id, ""), img("unknown", "notInStore", "")],
    };
    const missing = rehydrateDocMedia(doc);
    expect((doc.blocks[0] as ImageBlock).src).toMatch(/^blob:test\//); // filled
    expect((doc.blocks[1] as ImageBlock).src).toBe(""); // bytes absent → stays blank
    expect(missing).toEqual(["notInStore"]);
  });
});
