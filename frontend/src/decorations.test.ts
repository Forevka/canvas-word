import { describe, it, expect } from "vitest";
import { resolveDecorations, type DecorationGeometry, type DecorationSpec } from "./decorations";
import type { LayoutTree } from "./layout/layoutTree";
import type { DocSelection } from "@cw/shared";

// A fake layout tree + geometry so the routing/defaulting logic is testable
// without building a real LayoutTree. selectionRects returns one rect per range;
// caretRect returns a fixed point (or null for a "missing" block id).
const tree = {} as LayoutTree;
const range = (): DocSelection => ({ anchor: { blockId: "b", offset: 0 }, focus: { blockId: "b", offset: 3 } });

const geo: DecorationGeometry = {
  selectionRects: () => [{ pageIndex: 0, x: 10, y: 20, width: 30, height: 12 }],
  caretRect: (_t, pos) => (pos.blockId === "missing" ? null : { pageIndex: 0, x: 5, y: 6, height: 14 }),
};

const resolve = (specs: DecorationSpec[]) => resolveDecorations(specs, tree, undefined, geo);

describe("resolveDecorations", () => {
  it("routes each spec type into its bucket", () => {
    const out = resolve([
      { type: "highlight", range: range(), color: "yellow" },
      { type: "underline", range: range(), color: "red" },
      { type: "box", range: range(), color: "blue" },
      { type: "badge", at: { blockId: "b", offset: 0 }, color: "green", label: "1" },
    ]);
    expect(out.highlights).toHaveLength(1);
    expect(out.underlines).toHaveLength(1);
    expect(out.boxes).toHaveLength(1);
    expect(out.badges).toEqual([{ pageIndex: 0, x: 5, y: 6, color: "green", label: "1" }]);
  });

  it("applies default opacity/thickness and honors overrides", () => {
    const out = resolve([
      { type: "highlight", range: range(), color: "yellow" },
      { type: "highlight", range: range(), color: "yellow", opacity: 0.9 },
      { type: "underline", range: range(), color: "red" },
      { type: "underline", range: range(), color: "red", thickness: 3 },
      { type: "box", range: range(), color: "blue" },
    ]);
    expect(out.highlights[0]!.opacity).toBe(0.4);
    expect(out.highlights[1]!.opacity).toBe(0.9);
    expect(out.underlines[0]!.thickness).toBe(1.6);
    expect(out.underlines[1]!.thickness).toBe(3);
    expect(out.boxes[0]!.thickness).toBe(1);
  });

  it("carries the resolved page-local rect through", () => {
    const out = resolve([{ type: "highlight", range: range(), color: "yellow" }]);
    expect(out.highlights[0]).toMatchObject({ pageIndex: 0, x: 10, y: 20, width: 30, height: 12, color: "yellow" });
  });

  it("skips specs with an empty color", () => {
    const out = resolve([{ type: "highlight", range: range(), color: "" }]);
    expect(out.highlights).toHaveLength(0);
  });

  it("skips a badge whose position can't be resolved", () => {
    const out = resolve([{ type: "badge", at: { blockId: "missing", offset: 0 }, color: "green" }]);
    expect(out.badges).toHaveLength(0);
  });

  it("omits label when not provided (plain dot)", () => {
    const out = resolve([{ type: "badge", at: { blockId: "b", offset: 0 }, color: "green" }]);
    expect(out.badges[0]).toEqual({ pageIndex: 0, x: 5, y: 6, color: "green" });
    expect("label" in out.badges[0]!).toBe(false);
  });
});
