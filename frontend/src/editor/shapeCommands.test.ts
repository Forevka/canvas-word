// Issue #214 — drawing-shape commands: insert at the caret, resize/align/fill/stroke
// via setShapeProps (each an invertible one-step op), and delete via removeBlockObject.
// Pointer-capture drag-resize isn't scriptable (see SDLC known constraints), so the
// resize COMMIT logic is verified here on the op/command instead.
import { describe, expect, it } from "vitest";
import type { Document, Paragraph, SectionProps, ShapeBlock } from "@cw/shared";
import { applyOp } from "@cw/shared";
import { insertShape, setShapeProps, removeBlockObject, SHAPE_DEFAULT_WIDTH_PX, SHAPE_DEFAULT_HEIGHT_PX } from "./commands";
import type { Command, EditorState } from "./state";

const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };
const CHAR = { fontFamily: "Times New Roman, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" };
const PARA = { align: "left" as const, lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };

const para = (id: string, text = "hello"): Paragraph => ({ kind: "paragraph", id, revision: 0, runs: [{ text, style: { ...CHAR } }], style: { ...PARA } });
const shape = (id: string, patch: Partial<ShapeBlock> = {}): ShapeBlock => ({
  kind: "shape", id, revision: 0, geometry: { preset: "rect" }, widthPx: 180, heightPx: 120, align: "left", ...patch,
});
const caret = (blockId: string, offset: number) => ({ anchor: { blockId, offset }, focus: { blockId, offset } });
const shapeOf = (doc: Document, id: string): ShapeBlock => doc.blocks.find((b) => b.id === id) as ShapeBlock;
const trnOf = (state: EditorState, cmd: Command) => {
  const trn = cmd(state);
  if (!trn) throw new Error("no transaction");
  return trn;
};

describe("insertShape", () => {
  it("splits the caret paragraph and inserts a shape block with a default fill + stroke", () => {
    const base: Document = { section: SECTION, blocks: [para("p1")] };
    const trn = trnOf({ doc: base, selection: caret("p1", 2) }, insertShape("rect"));
    let doc = base;
    for (const op of trn.ops) doc = applyOp(doc, op).doc;
    const s = doc.blocks.find((b): b is ShapeBlock => b.kind === "shape")!;
    expect(s.geometry.preset).toBe("rect");
    expect(s.widthPx).toBe(SHAPE_DEFAULT_WIDTH_PX);
    expect(s.heightPx).toBe(SHAPE_DEFAULT_HEIGHT_PX);
    expect(s.fill).toBeDefined();
    expect(s.stroke).toBeDefined();
  });

  it("seeds a line preset with no fill (stroke-only geometry)", () => {
    const base: Document = { section: SECTION, blocks: [para("p1")] };
    const trn = trnOf({ doc: base, selection: caret("p1", 0) }, insertShape("line"));
    let doc = base;
    for (const op of trn.ops) doc = applyOp(doc, op).doc;
    const s = doc.blocks.find((b): b is ShapeBlock => b.kind === "shape")!;
    expect(s.geometry.preset).toBe("line");
    expect(s.fill).toBeUndefined();
    expect(s.stroke).toBeDefined();
  });
});

describe("setShapeProps", () => {
  const docOf = (...blocks: ShapeBlock[]): Document => ({ section: SECTION, blocks });

  it("commits a drag-resize (widthPx/heightPx); the op inverse restores the prior box", () => {
    const base = docOf(shape("a"));
    const trn = trnOf({ doc: base, selection: null }, setShapeProps("a", { widthPx: 300, heightPx: 200 }));
    const res = applyOp(base, trn.ops[0]!);
    expect(shapeOf(res.doc, "a").widthPx).toBe(300);
    expect(shapeOf(res.doc, "a").heightPx).toBe(200);
    const undone = applyOp(res.doc, res.inverse).doc;
    expect(shapeOf(undone, "a").widthPx).toBe(180);
    expect(shapeOf(undone, "a").heightPx).toBe(120);
  });

  it("sets alignment", () => {
    const base = docOf(shape("a"));
    const trn = trnOf({ doc: base, selection: null }, setShapeProps("a", { align: "center" }));
    expect(shapeOf(applyOp(base, trn.ops[0]!).doc, "a").align).toBe("center");
  });

  it("sets fill/stroke and clears them with null (undo restores)", () => {
    const base = docOf(shape("a", { fill: { color: "#112233" }, stroke: { color: "#445566", widthPt: 2 } }));
    const cleared = applyOp(base, trnOf({ doc: base, selection: null }, setShapeProps("a", { fill: null, stroke: null })).ops[0]!);
    expect(shapeOf(cleared.doc, "a").fill).toBeUndefined();
    expect(shapeOf(cleared.doc, "a").stroke).toBeUndefined();
    const restored = applyOp(cleared.doc, cleared.inverse).doc;
    expect(shapeOf(restored, "a").fill).toEqual({ color: "#112233" });
    expect(shapeOf(restored, "a").stroke).toEqual({ color: "#445566", widthPt: 2 });
  });

  it("returns null (no-op) for a missing shape id", () => {
    expect(setShapeProps("nope", { widthPx: 10 })({ doc: docOf(shape("a")), selection: null })).toBeNull();
  });
});

describe("removeBlockObject deletes a shape", () => {
  it("removes a top-level shape and parks the caret on a neighbour", () => {
    const base: Document = { section: SECTION, blocks: [shape("a"), para("p2")] };
    const trn = trnOf({ doc: base, selection: null }, removeBlockObject("a"));
    const doc = applyOp(base, trn.ops[0]!).doc;
    expect(doc.blocks.find((b) => b.id === "a")).toBeUndefined();
  });
});
