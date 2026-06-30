import { describe, expect, it } from "vitest";
import type { Document, ImageBlock, SectionProps } from "@cw/shared";
import { applyOp } from "@cw/shared";
import { setImageCropCmd } from "./commands";
import type { Command, EditorState } from "./state";

const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

const img = (id: string, crop?: ImageBlock["crop"]): ImageBlock => ({
  kind: "image", id, revision: 0, src: "blob:x", widthPx: 100, heightPx: 100, align: "center",
  ...(crop ? { crop } : {}),
});
const docOf = (...blocks: ImageBlock[]): Document => ({ section: SECTION, blocks });
const imageOf = (doc: Document, id: string): ImageBlock => doc.blocks.find((b) => b.id === id) as ImageBlock;

const trnOf = (state: EditorState, cmd: Command) => {
  const trn = cmd(state);
  if (!trn) throw new Error("no transaction");
  return trn;
};

describe("setImageCropCmd — interactive crop (a:srcRect)", () => {
  const CROP = { left: 0.1, top: 0.2, right: 0.05, bottom: 0.15 };

  it("writes the crop insets onto the selected image", () => {
    const trn = trnOf({ doc: docOf(img("a")), selection: null }, setImageCropCmd("a", CROP));
    let doc = docOf(img("a"));
    for (const op of trn.ops) doc = applyOp(doc, op).doc;
    expect(imageOf(doc, "a").crop).toEqual(CROP);
  });

  it("is one undoable step: the op inverse restores the prior (un-cropped) state", () => {
    const base = docOf(img("a"));
    const trn = trnOf({ doc: base, selection: null }, setImageCropCmd("a", CROP));
    const res = applyOp(base, trn.ops[0]!);
    expect(imageOf(res.doc, "a").crop).toEqual(CROP);
    const undone = applyOp(res.doc, res.inverse).doc;
    expect(imageOf(undone, "a").crop).toBeUndefined(); // back to no crop
  });

  it("clears an existing crop when passed null, and undo restores it", () => {
    const base = docOf(img("a", CROP));
    const trn = trnOf({ doc: base, selection: null }, setImageCropCmd("a", null));
    const res = applyOp(base, trn.ops[0]!);
    expect(imageOf(res.doc, "a").crop).toBeUndefined();
    const undone = applyOp(res.doc, res.inverse).doc;
    expect(imageOf(undone, "a").crop).toEqual(CROP);
  });

  it("bails (no transaction) when the target id is not an image", () => {
    expect(setImageCropCmd("nope", CROP)({ doc: docOf(img("a")), selection: null })).toBeNull();
  });

  it("supports a transient origin for live previews (no committed step)", () => {
    const trn = trnOf({ doc: docOf(img("a")), selection: null }, setImageCropCmd("a", CROP, "transient"));
    expect(trn.origin).toBe("transient");
  });
});
