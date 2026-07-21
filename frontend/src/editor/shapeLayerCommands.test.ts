// Issue #217 (Part 4 of #206) — shape wrap / float anchor / z-order commands.
// Mirrors imageLayerCommands.test.ts: the shape layer commands reuse the same
// anchorFor / anchorZRange helpers as images, over a stacking space SHARED with
// images, so a shape can be lifted above/below images too.
import { describe, expect, it } from "vitest";
import type { Document, ImageBlock, SectionProps, ShapeBlock, TableBlock, Paragraph } from "@cw/shared";
import { applyOp } from "@cw/shared";
import { setShapeLayer, bringShapeToFront, sendShapeToBack, moveAnchoredShape, setShapeProps } from "./commands";
import type { Command, EditorState } from "./state";

const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

const shape = (id: string, anchor?: ShapeBlock["anchor"], wrap?: ShapeBlock["wrap"]): ShapeBlock => ({
  kind: "shape", id, revision: 0, geometry: { preset: "rect" }, widthPx: 100, heightPx: 100, align: "center",
  ...(anchor ? { anchor } : {}),
  ...(wrap ? { wrap } : {}),
});
const img = (id: string, anchor?: ImageBlock["anchor"]): ImageBlock => ({
  kind: "image", id, revision: 0, src: "blob:x", widthPx: 100, heightPx: 100, align: "center",
  ...(anchor ? { anchor } : {}),
});
const docOf = (...blocks: Document["blocks"]): Document => ({ section: SECTION, blocks });

const apply = (state: EditorState, cmd: Command): Document => {
  const trn = cmd(state);
  if (!trn) throw new Error("no transaction");
  let doc = state.doc;
  for (const op of trn.ops) doc = applyOp(doc, op).doc;
  return doc;
};
const shapeOf = (doc: Document, id: string): ShapeBlock => doc.blocks.find((b) => b.id === id) as ShapeBlock;

describe("shape layer / z-order commands", () => {
  it("setShapeLayer lifts an in-line shape into the behind layer and clears wrap", () => {
    const doc = apply({ doc: docOf(shape("a", undefined, "square")), selection: null }, setShapeLayer("a", true));
    const out = shapeOf(doc, "a");
    expect(out.anchor?.behind).toBe(true);
    expect(out.wrap).toBeUndefined(); // exclusive — wrap cleared
  });

  it("setShapeLayer keeps an existing anchor's offsets, only flipping the layer", () => {
    const a = shape("a", { behind: true, offsetXPx: -98, offsetYPx: -238, relFromH: "margin", relFromV: "paragraph" });
    const doc = apply({ doc: docOf(a), selection: null }, setShapeLayer("a", false));
    expect(shapeOf(doc, "a").anchor).toMatchObject({ behind: false, offsetXPx: -98, offsetYPx: -238 });
  });

  it("bringShapeToFront sets the front layer with the highest z", () => {
    const a = shape("a", { behind: true, offsetXPx: 0, offsetYPx: 0, relFromH: "page", relFromV: "page", z: 3 });
    const b = shape("b", { behind: false, offsetXPx: 0, offsetYPx: 0, relFromH: "page", relFromV: "page", z: 7 });
    const doc = apply({ doc: docOf(a, b), selection: null }, bringShapeToFront("a"));
    const out = shapeOf(doc, "a");
    expect(out.anchor?.behind).toBe(false);
    expect(out.anchor?.z).toBe(8); // max(3,7) + 1
  });

  it("sendShapeToBack sets the behind layer with the lowest z", () => {
    const a = shape("a", { behind: false, offsetXPx: 0, offsetYPx: 0, relFromH: "page", relFromV: "page", z: 2 });
    const b = shape("b", { behind: true, offsetXPx: 0, offsetYPx: 0, relFromH: "page", relFromV: "page", z: -1 });
    const doc = apply({ doc: docOf(a, b), selection: null }, sendShapeToBack("a"));
    const out = shapeOf(doc, "a");
    expect(out.anchor?.behind).toBe(true);
    expect(out.anchor?.z).toBe(-2); // min(2,-1) - 1
  });

  it("z-order is shared with images: bringShapeToFront clears an anchored image on top", () => {
    const anchoredImg = img("im", { behind: false, offsetXPx: 0, offsetYPx: 0, relFromH: "page", relFromV: "page", z: 5 });
    const s = shape("a", { behind: false, offsetXPx: 0, offsetYPx: 0, relFromH: "page", relFromV: "page", z: 1 });
    const doc = apply({ doc: docOf(anchoredImg, s), selection: null }, bringShapeToFront("a"));
    expect(shapeOf(doc, "a").anchor?.z).toBe(6); // above the z=5 image
  });

  it("moveAnchoredShape sets new absolute offsets, preserving the rest of the anchor", () => {
    const a = shape("a", { behind: true, offsetXPx: 0, offsetYPx: 0, relFromH: "margin", relFromV: "paragraph", z: 5 });
    const doc = apply({ doc: docOf(a), selection: null }, moveAnchoredShape("a", -120, 40));
    expect(shapeOf(doc, "a").anchor).toMatchObject({ offsetXPx: -120, offsetYPx: 40, behind: true, relFromH: "margin", z: 5 });
  });

  it("moveAnchoredShape is a no-op on an in-flow (non-anchored) shape", () => {
    expect(moveAnchoredShape("a", 10, 10)({ doc: docOf(shape("a", undefined, "square")), selection: null })).toBeNull();
  });

  it("clearing the anchor (wrap toggle) is undoable back to the anchored state", () => {
    const a = shape("a", { behind: true, offsetXPx: -98, offsetYPx: -238, relFromH: "margin", relFromV: "paragraph" });
    const res = applyOp(docOf(a), { type: "setShapeProps", blockId: "a", patch: { wrap: "block", anchor: null } });
    expect((res.doc.blocks[0] as ShapeBlock).anchor).toBeUndefined();
    expect((res.doc.blocks[0] as ShapeBlock).wrap).toBe("block");
    // Inverse restores the anchor and clears the wrap.
    const undone = applyOp(res.doc, res.inverse).doc;
    expect((undone.blocks[0] as ShapeBlock).anchor).toMatchObject({ behind: true, offsetXPx: -98 });
    expect((undone.blocks[0] as ShapeBlock).wrap).toBeUndefined();
  });

  it("setShapeProps anchors a shape inside a table cell (cell-aware, like images)", () => {
    const cellShape = shape("cs");
    const para: Paragraph = { kind: "paragraph", id: "p", revision: 0, runs: [{ text: "", style: {} as never }], style: {} as never };
    const table: TableBlock = { kind: "table", id: "t", revision: 0, rows: [{ cells: [{ id: "c", blocks: [cellShape, para] }] }] };
    const doc: Document = { section: SECTION, blocks: [table] };
    const trn = setShapeProps("cs", { anchor: { behind: true, offsetXPx: 5, offsetYPx: 5, relFromH: "margin", relFromV: "paragraph" } })({ doc, selection: null });
    expect(trn).not.toBeNull();
    let out = doc;
    for (const op of trn!.ops) out = applyOp(out, op).doc;
    const updated = (out.blocks[0] as TableBlock).rows[0]!.cells[0]!.blocks[0] as ShapeBlock;
    expect(updated.anchor?.behind).toBe(true);
  });
});
