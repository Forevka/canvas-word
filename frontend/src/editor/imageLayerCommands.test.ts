import { describe, expect, it } from "vitest";
import type { Document, ImageBlock, Paragraph, SectionProps } from "@cw/shared";
import { applyOp } from "@cw/shared";
import { setImageLayer, bringImageToFront, sendImageToBack, moveAnchoredImage } from "./commands";
import type { Command } from "./state";
import type { EditorState } from "./state";

const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

const img = (id: string, anchor?: ImageBlock["anchor"], wrap?: ImageBlock["wrap"]): ImageBlock => ({
  kind: "image", id, revision: 0, src: "blob:x", widthPx: 100, heightPx: 100, align: "center",
  ...(anchor ? { anchor } : {}),
  ...(wrap ? { wrap } : {}),
});
const docOf = (...blocks: ImageBlock[] | Paragraph[]): Document => ({ section: SECTION, blocks });

const apply = (state: EditorState, cmd: Command): Document => {
  const trn = cmd(state);
  if (!trn) throw new Error("no transaction");
  let doc = state.doc;
  for (const op of trn.ops) doc = applyOp(doc, op).doc;
  return doc;
};
const imageOf = (doc: Document, id: string): ImageBlock => doc.blocks.find((b) => b.id === id) as ImageBlock;

describe("image layer / z-order commands", () => {
  it("setImageLayer lifts an in-line image into the behind layer and clears wrap", () => {
    const a = img("a", undefined, "square");
    const doc = apply({ doc: docOf(a), selection: null }, setImageLayer("a", true));
    const out = imageOf(doc, "a");
    expect(out.anchor?.behind).toBe(true);
    expect(out.wrap).toBeUndefined(); // exclusive — wrap cleared
  });

  it("setImageLayer keeps an existing anchor's offsets, only flipping the layer", () => {
    const a = img("a", { behind: true, offsetXPx: -98, offsetYPx: -238, relFromH: "margin", relFromV: "paragraph" });
    const doc = apply({ doc: docOf(a), selection: null }, setImageLayer("a", false));
    expect(imageOf(doc, "a").anchor).toMatchObject({ behind: false, offsetXPx: -98, offsetYPx: -238 });
  });

  it("bringImageToFront sets the front layer with the highest z", () => {
    const a = img("a", { behind: true, offsetXPx: 0, offsetYPx: 0, relFromH: "page", relFromV: "page", z: 3 });
    const b = img("b", { behind: false, offsetXPx: 0, offsetYPx: 0, relFromH: "page", relFromV: "page", z: 7 });
    const doc = apply({ doc: docOf(a, b), selection: null }, bringImageToFront("a"));
    const out = imageOf(doc, "a");
    expect(out.anchor?.behind).toBe(false);
    expect(out.anchor?.z).toBe(8); // max(3,7) + 1
  });

  it("sendImageToBack sets the behind layer with the lowest z", () => {
    const a = img("a", { behind: false, offsetXPx: 0, offsetYPx: 0, relFromH: "page", relFromV: "page", z: 2 });
    const b = img("b", { behind: true, offsetXPx: 0, offsetYPx: 0, relFromH: "page", relFromV: "page", z: -1 });
    const doc = apply({ doc: docOf(a, b), selection: null }, sendImageToBack("a"));
    const out = imageOf(doc, "a");
    expect(out.anchor?.behind).toBe(true);
    expect(out.anchor?.z).toBe(-2); // min(2,-1) - 1
  });

  it("moveAnchoredImage sets new absolute offsets, preserving the rest of the anchor", () => {
    const a = img("a", { behind: true, offsetXPx: 0, offsetYPx: 0, relFromH: "margin", relFromV: "paragraph", z: 5 });
    const doc = apply({ doc: docOf(a), selection: null }, moveAnchoredImage("a", -120, 40));
    expect(imageOf(doc, "a").anchor).toMatchObject({ offsetXPx: -120, offsetYPx: 40, behind: true, relFromH: "margin", z: 5 });
  });

  it("moveAnchoredImage is a no-op on an in-flow (non-anchored) image", () => {
    const a = img("a", undefined, "square");
    expect(moveAnchoredImage("a", 10, 10)({ doc: docOf(a), selection: null })).toBeNull();
  });

  it("clearing the anchor (wrap toggle) is undoable back to the anchored state", () => {
    const a = img("a", { behind: true, offsetXPx: -98, offsetYPx: -238, relFromH: "margin", relFromV: "paragraph" });
    const base = docOf(a);
    const res = applyOp(base, { type: "setImageProps", blockId: "a", patch: { wrap: "block", anchor: null } });
    expect((res.doc.blocks[0] as ImageBlock).anchor).toBeUndefined();
    expect((res.doc.blocks[0] as ImageBlock).wrap).toBe("block");
    // Inverse restores the anchor and clears the wrap.
    const undone = applyOp(res.doc, res.inverse).doc;
    expect((undone.blocks[0] as ImageBlock).anchor).toMatchObject({ behind: true, offsetXPx: -98 });
    expect((undone.blocks[0] as ImageBlock).wrap).toBeUndefined();
  });
});
