// Issue #214 — drawing-shape commands: insert at the caret, resize/align/fill/stroke
// via setShapeProps (each an invertible one-step op), and delete via removeBlockObject.
// Pointer-capture drag-resize isn't scriptable (see SDLC known constraints), so the
// resize COMMIT logic is verified here on the op/command instead.
import { describe, expect, it } from "vitest";
import type { Document, Paragraph, SectionProps, ShapeBlock } from "@cw/shared";
import { applyOp } from "@cw/shared";
import { insertShape, insertTextBox, addShapeText, setShapeProps, removeBlockObject, clampAnchorOffsetToPage, SHAPE_DEFAULT_WIDTH_PX, SHAPE_DEFAULT_HEIGHT_PX } from "./commands";
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

  it("uses the caller-supplied id so the host can select + reveal it (A1)", () => {
    const base: Document = { section: SECTION, blocks: [para("p1")] };
    const trn = trnOf({ doc: base, selection: caret("p1", 2) }, insertShape("ellipse", "mine"));
    let doc = base;
    for (const op of trn.ops) doc = applyOp(doc, op).doc;
    expect(doc.blocks.find((b) => b.id === "mine")?.kind).toBe("shape");
  });
});

describe("insertShape field/TOC/SDT guard (A2)", () => {
  const tocPara = (id: string, text = "1. Intro"): Paragraph => ({
    ...para(id, text), style: { ...PARA, tocEntry: { targetId: "h1", level: 1 } },
  });
  const applyAll = (base: Document, trn: { ops: import("@cw/shared").Op[] }): Document => {
    let doc = base;
    for (const op of trn.ops) doc = applyOp(doc, op).doc;
    return doc;
  };

  it("redirects the insert to AFTER a TOC instead of splitting it", () => {
    const base: Document = { section: SECTION, blocks: [tocPara("t1"), tocPara("t2"), para("p3", "body")] };
    // Caret inside the first TOC entry.
    const trn = trnOf({ doc: base, selection: caret("t1", 2) }, insertShape("rect"));
    // A single insertBlock op — no splitParagraph (the TOC entry is left intact).
    expect(trn.ops.every((o) => o.type !== "splitParagraph")).toBe(true);
    const doc = applyAll(base, trn);
    const ids = doc.blocks.map((b) => b.id);
    // Shape landed after the whole TOC region (t1, t2), before the body paragraph.
    expect(ids.slice(0, 2)).toEqual(["t1", "t2"]);
    expect(doc.blocks[2]?.kind).toBe("shape");
    expect(ids[3]).toBe("p3");
    // The TOC entries survive unchanged.
    expect((doc.blocks[0] as Paragraph).style.tocEntry).toBeDefined();
  });

  it("redirects out of a block-level field", () => {
    const field: Paragraph = { ...para("f", "field text"), fieldId: "fld1" };
    const base: Document = { section: SECTION, blocks: [field, para("after")] };
    const trn = trnOf({ doc: base, selection: caret("f", 3) }, insertShape("rect"));
    expect(trn.ops.every((o) => o.type !== "splitParagraph")).toBe(true);
    const doc = applyAll(base, trn);
    expect(doc.blocks.map((b) => b.id)[0]).toBe("f");
    expect(doc.blocks[1]?.kind).toBe("shape");
    expect(doc.blocks[2]?.id).toBe("after");
  });

  it("redirects when the caret is strictly inside an inline field run", () => {
    const p: Paragraph = {
      kind: "paragraph", id: "p1", revision: 0, style: { ...PARA },
      runs: [
        { text: "a", style: { ...CHAR } },
        { text: "PAGE", style: { ...CHAR, fieldId: "pg" } },
        { text: "b", style: { ...CHAR } },
      ],
    };
    const base: Document = { section: SECTION, blocks: [p, para("q")] };
    // offset 3 is inside "PAGE" (chars 1..4) — both sides carry fieldId "pg".
    const trn = trnOf({ doc: base, selection: caret("p1", 3) }, insertShape("rect"));
    expect(trn.ops.every((o) => o.type !== "splitParagraph")).toBe(true);
    const doc = applyAll(base, trn);
    expect(doc.blocks[0]?.id).toBe("p1");
    expect(doc.blocks[1]?.kind).toBe("shape"); // after the paragraph, not splitting the field
  });

  it("still splits a normal body paragraph (caret at a field's edge is safe)", () => {
    const base: Document = { section: SECTION, blocks: [para("p1", "hello")] };
    const trn = trnOf({ doc: base, selection: caret("p1", 2) }, insertShape("rect"));
    expect(trn.ops.some((o) => o.type === "splitParagraph")).toBe(true);
  });

  it("keeps the original selection when the block after the region isn't a paragraph", () => {
    // A TOC entry immediately followed by an image (a non-paragraph block can't hold a
    // text caret): the redirect must not land an invalid caret on it — fall back to the
    // caret's original position.
    const img = { kind: "image" as const, id: "im1", revision: 0, src: "blob:x", widthPx: 10, heightPx: 10, align: "left" as const };
    const base: Document = { section: SECTION, blocks: [tocPara("t1"), img] };
    const trn = trnOf({ doc: base, selection: caret("t1", 2) }, insertShape("rect"));
    expect(trn.ops.every((o) => o.type !== "splitParagraph")).toBe(true);
    // Selection stays on the TOC entry (offset 2), NOT on the image block.
    expect(trn.selectionAfter?.focus.blockId).toBe("t1");
    expect(trn.selectionAfter?.focus.offset).toBe(2);
  });
});

describe("clampAnchorOffsetToPage (F2)", () => {
  const page = { width: 800, height: 1000 };
  const size = { width: 100, height: 100 };

  it("leaves an on-page offset untouched", () => {
    expect(clampAnchorOffsetToPage({ x: 0, y: 0 }, { x: 50, y: 60 }, size, page)).toEqual({ x: 50, y: 60 });
  });

  it("clamps a far-right / far-down drag so a sliver stays on-page", () => {
    // keep = 24 → placed x maxes at 800-24=776, y at 1000-24=976.
    expect(clampAnchorOffsetToPage({ x: 0, y: 0 }, { x: 5000, y: 5000 }, size, page)).toEqual({ x: 776, y: 976 });
  });

  it("clamps a far-left / far-up drag so the box can't vanish past the origin edge", () => {
    // placed x mins at keep - width = 24-100 = -76.
    expect(clampAnchorOffsetToPage({ x: 0, y: 0 }, { x: -5000, y: -5000 }, size, page)).toEqual({ x: -76, y: -76 });
  });

  it("accounts for a non-zero anchor origin", () => {
    // origin 100,50 + offset 5000 → placed 5100 clamps to 776 → offset 676; y 5050→976→926.
    expect(clampAnchorOffsetToPage({ x: 100, y: 50 }, { x: 5000, y: 5000 }, size, page)).toEqual({ x: 676, y: 926 });
  });
});

describe("insertTextBox (issue #235)", () => {
  it("inserts a rectangle carrying an empty, editable text body seeded with the given ids", () => {
    const base: Document = { section: SECTION, blocks: [para("p1")] };
    const trn = trnOf({ doc: base, selection: caret("p1", 2) }, insertTextBox("box1", "boxpar1"));
    let doc = base;
    for (const op of trn.ops) doc = applyOp(doc, op).doc;
    const s = doc.blocks.find((b): b is ShapeBlock => b.kind === "shape")!;
    expect(s.id).toBe("box1");
    expect(s.geometry.preset).toBe("rect");
    // A text body with exactly one empty paragraph — the shape the wps:txbx
    // round-trip already handles, and enough for the caret to land inside.
    expect(s.text?.blocks.length).toBe(1);
    expect(s.text?.blocks[0]?.id).toBe("boxpar1");
    expect(s.text?.blocks[0]?.runs.map((r) => r.text).join("")).toBe("");
    expect(s.fill).toBeDefined();
    expect(s.stroke).toBeDefined();
  });
});

describe("addShapeText (issue #235)", () => {
  const docOf = (...blocks: ShapeBlock[]): Document => ({ section: SECTION, blocks });

  it("starts an empty text body on a text-less top-level shape; the op inverse clears it", () => {
    const base = docOf(shape("a"));
    const trn = trnOf({ doc: base, selection: null }, addShapeText("a", "np"));
    const res = applyOp(base, trn.ops[0]!);
    const s = shapeOf(res.doc, "a");
    expect(s.text?.blocks.length).toBe(1);
    expect(s.text?.blocks[0]?.id).toBe("np");
    expect(s.text?.blocks[0]?.runs.map((r) => r.text).join("")).toBe("");
    // Undo removes the whole body (single-step).
    expect(shapeOf(applyOp(res.doc, res.inverse).doc, "a").text).toBeUndefined();
  });

  it("is a no-op for a shape that already has a text body", () => {
    const base = docOf(shape("a", { text: { blocks: [para("boxpar", "hi")] } }));
    expect(addShapeText("a", "np")({ doc: base, selection: null })).toBeNull();
  });

  it("is a no-op for a missing shape id", () => {
    expect(addShapeText("nope", "np")({ doc: docOf(shape("a")), selection: null })).toBeNull();
  });

  it("refuses a cell-nested shape (its text box stays read-only, #219/#230)", () => {
    const cellShape = shape("cellbox");
    const table = { kind: "table" as const, id: "tbl", revision: 0, rows: [{ cells: [{ id: "c0", blocks: [cellShape] }] }] };
    const doc: Document = { section: SECTION, blocks: [table] };
    expect(addShapeText("cellbox", "np")({ doc, selection: null })).toBeNull();
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

  it("sets rotation and geometry (preset + adjust); the op inverse restores them", () => {
    const base = docOf(shape("a"));
    const trn = trnOf({ doc: base, selection: null }, setShapeProps("a", { geometry: { preset: "roundRect", adjust: { adj: 18000 } }, rotation: 20 }));
    const res = applyOp(base, trn.ops[0]!);
    expect(shapeOf(res.doc, "a").geometry).toEqual({ preset: "roundRect", adjust: { adj: 18000 } });
    expect(shapeOf(res.doc, "a").rotation).toBe(20);
    const undone = applyOp(res.doc, res.inverse).doc;
    expect(shapeOf(undone, "a").geometry).toEqual({ preset: "rect" });
    expect(shapeOf(undone, "a").rotation).toBeUndefined();
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
