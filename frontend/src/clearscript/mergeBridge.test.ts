import { describe, expect, it } from "vitest";
import type { Block, Document, ImageBlock, Paragraph, SectionProps } from "@cw/shared";
import { DEFAULT_CHAR_STYLE, DEFAULT_PARA_STYLE } from "@cw/shared";
import { mergeBridge } from "./mergeBridge";

const sec: SectionProps = { pageWidthPx: 794, pageHeightPx: 1123, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };
const para = (id: string, text: string): Paragraph => ({
  kind: "paragraph", id, revision: 0, runs: [{ text, style: { ...DEFAULT_CHAR_STYLE } }], style: { ...DEFAULT_PARA_STYLE },
});
const img = (id: string, src: string): ImageBlock => ({ kind: "image", id, revision: 0, src, widthPx: 10, heightPx: 10, align: "left" });
const doc = (blocks: Block[]): Document => ({ section: { ...sec }, blocks });

const imgOf = (d: Document): ImageBlock[] => d.blocks.filter((b): b is ImageBlock => b.kind === "image");

describe("mergeBridge media union", () => {
  it("dedupes identical images sharing a src key", () => {
    const bytes = new Uint8Array([1, 2, 3]);
    const dest = doc([img("d1", "cw-media:0")]);
    const src = doc([img("s1", "cw-media:0")]);

    const r = mergeBridge(dest, { "cw-media:0": bytes }, src, { "cw-media:0": new Uint8Array([1, 2, 3]) }, { sectionBreak: "none" });

    expect(Object.keys(r.images)).toHaveLength(1); // identical bytes → one entry
    expect(r.mediaCount).toBe(1);
    // Both image blocks still resolve against the single key.
    for (const i of imgOf(r.doc)) expect(r.images[i.src]).toBeDefined();
  });

  it("reassigns a colliding src key that holds DIFFERENT bytes and rewrites the source block", () => {
    const dest = doc([img("d1", "cw-media:0")]);
    const src = doc([img("s1", "cw-media:0")]);

    const r = mergeBridge(
      dest,
      { "cw-media:0": new Uint8Array([1, 1, 1]) },
      src,
      { "cw-media:0": new Uint8Array([9, 9, 9]) },
      { sectionBreak: "none" },
    );

    expect(Object.keys(r.images)).toHaveLength(2); // distinct images kept apart
    const images = imgOf(r.doc);
    // The destination image keeps cw-media:0; the source image was rewritten.
    const destImg = images.find((i) => i.id === "d1")!;
    expect(destImg.src).toBe("cw-media:0");
    const srcImg = images.find((i) => i.id !== "d1")!;
    expect(srcImg.src).not.toBe("cw-media:0");
    // Every image resolves, and to the RIGHT bytes.
    expect(r.images[destImg.src]).toEqual(new Uint8Array([1, 1, 1]));
    expect(r.images[srcImg.src]).toEqual(new Uint8Array([9, 9, 9]));
  });

  it("carries builder docs (no images) through cleanly", () => {
    const r = mergeBridge(doc([para("d1", "A")]), undefined, doc([para("s1", "B")]), undefined, { sectionBreak: "none" });
    expect(r.mediaCount).toBe(0);
    expect(r.doc.blocks).toHaveLength(2);
  });
});
