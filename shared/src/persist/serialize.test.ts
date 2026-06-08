import { describe, expect, it } from "vitest";
import type { Document, ImageBlock, Paragraph, SectionProps, TableBlock } from "../model/document";
import { mediaIdForBytes } from "./media";
import { collectMediaIds, deserializeDocument, forEachImage, serializeDocument } from "./serialize";

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

const para = (id: string, text: string): Paragraph => ({
  kind: "paragraph",
  id,
  revision: 0,
  runs: [{ text, style: { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" } }],
  style: { align: "left", lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 },
});

const tableWithImage = (cellImg: ImageBlock): TableBlock => ({
  kind: "table",
  id: "t1",
  revision: 0,
  rows: [{ cells: [{ id: "c1", blocks: [cellImg] }] }],
});

describe("serializeDocument / deserializeDocument", () => {
  it("drops blob: src but keeps mediaId; round-trips structure", () => {
    const doc: Document = {
      section: section(),
      blocks: [para("p1", "hello"), img("i1", "hashA", "blob:http://x/abc")],
    };
    const snap = serializeDocument(doc);
    expect(snap.version).toBe(1);
    const storedImg = snap.doc.blocks[1] as ImageBlock;
    expect(storedImg.src).toBe(""); // blob URL dropped
    expect(storedImg.mediaId).toBe("hashA"); // portable handle kept

    const back = deserializeDocument(snap);
    expect(back.blocks[0]).toEqual(doc.blocks[0]); // paragraph identical
    expect((back.blocks[1] as ImageBlock).mediaId).toBe("hashA");
    // serialize must not mutate the input doc
    expect((doc.blocks[1] as ImageBlock).src).toBe("blob:http://x/abc");
  });

  it("keeps inline data: URLs (already portable)", () => {
    const doc: Document = { section: section(), blocks: [img("i1", undefined, "data:image/png;base64,AAAA")] };
    const snap = serializeDocument(doc);
    expect((snap.doc.blocks[0] as ImageBlock).src).toBe("data:image/png;base64,AAAA");
  });

  it("collects mediaIds across body, table cells, and bands", () => {
    const doc: Document = {
      section: { ...section(), header: [img("h1", "hashHeader", "blob:1")] },
      blocks: [img("i1", "hashBody", "blob:2"), tableWithImage(img("i2", "hashCell", "blob:3"))],
    };
    expect(collectMediaIds(doc)).toEqual(new Set(["hashBody", "hashCell", "hashHeader"]));

    const seen: string[] = [];
    forEachImage(doc, (im) => seen.push(im.id));
    expect(seen.sort()).toEqual(["h1", "i1", "i2"]);
  });
});

describe("mediaIdForBytes", () => {
  it("is the sha256 hex of the bytes (deterministic, content-addressed)", async () => {
    // sha256("") = e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855
    expect(await mediaIdForBytes(new Uint8Array())).toBe(
      "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
    );
    const abc = new TextEncoder().encode("abc");
    // sha256("abc") = ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad
    expect(await mediaIdForBytes(abc)).toBe(
      "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad",
    );
    // identical bytes → identical id (dedup)
    expect(await mediaIdForBytes(new TextEncoder().encode("abc"))).toBe(await mediaIdForBytes(abc));
  });
});
