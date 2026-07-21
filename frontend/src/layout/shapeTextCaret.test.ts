// Issue #219 — the caret must reach a drawing shape's text box body. The box's
// nested paragraphs are laid out in the shape's LOCAL frame; a ShapeScope index
// shifts them to absolute page coords so hitTest / caretRect work through the
// SAME geometry queries — while the box stays INVISIBLE to unscoped body queries.
import "./test-canvas-setup";

import { describe, it, expect } from "vitest";
import type { CharStyle, Document, ParaStyle, Paragraph, SectionProps, ShapeBlock } from "@cw/shared";
import { createLayoutEngine } from "./engine";
import { caretRect, hitTest } from "./geometry";
import type { PlacedBlock } from "./layoutTree";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA: ParaStyle = { align: "left", lineHeight: 1, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let id = 0;
const fresh = () => `t${id++}`;
const para = (text: string): Paragraph => ({ kind: "paragraph", id: fresh(), revision: 0, runs: [{ text, style: CHAR }], style: PARA });
const box = (paras: Paragraph[]): ShapeBlock => ({
  kind: "shape", id: "shp", revision: 0, geometry: { preset: "rect" },
  widthPx: 300, heightPx: 120, align: "left", text: { blocks: paras },
});
const layout = (d: Document) => createLayoutEngine().layout(d);

const placedShape = (tree: ReturnType<typeof layout>): { page: number; pb: PlacedBlock } => {
  for (const pg of tree.pages) for (const pb of pg.blocks) if (pb.blockId === "shp" && pb.shape?.text) return { page: pg.index, pb };
  throw new Error("shape not placed with text");
};

describe("shape text box caret geometry (issue #219)", () => {
  it("hit-tests a point inside the box (scoped) into one of the box paragraphs", () => {
    const p0 = para("First line");
    const p1 = para("Second line");
    const tree = layout({ section: SECTION, blocks: [para("body"), box([p0, p1])] });
    const { page, pb } = placedShape(tree);
    const t = pb.shape!.text!;
    // A point inside the box, over the first text line.
    const x = pb.x + t.offsetX + 4;
    const y = pb.y + t.offsetY + 4;
    const hit = hitTest(tree, page, x, y, { shape: "shp", pageIndex: page });
    expect(hit).not.toBeNull();
    expect([p0.id, p1.id]).toContain(hit!.blockId);
  });

  it("caretRect (scoped) lands inside the box; unscoped it is invisible", () => {
    const p0 = para("Boxed");
    const tree = layout({ section: SECTION, blocks: [para("body"), box([p0])] });
    const { page, pb } = placedShape(tree);
    const scoped = caretRect(tree, { blockId: p0.id, offset: 0 }, { shape: "shp", pageIndex: page });
    expect(scoped).not.toBeNull();
    // Caret sits within the shape's box, in absolute page coords.
    expect(scoped!.pageIndex).toBe(page);
    expect(scoped!.x).toBeGreaterThanOrEqual(pb.x);
    expect(scoped!.x).toBeLessThanOrEqual(pb.x + pb.shape!.width);
    expect(scoped!.y).toBeGreaterThanOrEqual(pb.y);
    expect(scoped!.y).toBeLessThanOrEqual(pb.y + pb.shape!.height);
    // The box paragraph is a CLOSED sub-flow: an UNSCOPED query can't see it.
    expect(caretRect(tree, { blockId: p0.id, offset: 0 })).toBeNull();
  });

  // Issue #235 — Insert Text Box / "Add text" seed an EMPTY body (one empty
  // paragraph). The caret must still land inside so typing starts immediately.
  it("places a caret at offset 0 of an empty text body (Insert Text Box)", () => {
    const empty = para(""); // one empty paragraph, as the authoring gestures seed
    const tree = layout({ section: SECTION, blocks: [para("body"), box([empty])] });
    const { page, pb } = placedShape(tree);
    const scoped = caretRect(tree, { blockId: empty.id, offset: 0 }, { shape: "shp", pageIndex: page });
    expect(scoped).not.toBeNull();
    expect(scoped!.pageIndex).toBe(page);
    expect(scoped!.x).toBeGreaterThanOrEqual(pb.x);
    expect(scoped!.y).toBeGreaterThanOrEqual(pb.y);
  });
});
