// Issue #244 (F5) — group / ungroup authoring commands. Multi-select drag isn't
// scriptable (SDLC known constraints), so the group/ungroup COMMIT logic is verified
// here on the ops: N shapes collapse into one wpg:wgp container (identity child map),
// ungroup restores them at their page positions, and the round-trip is exact.
import { describe, expect, it } from "vitest";
import type { Document, SectionProps, ShapeBlock } from "@cw/shared";
import { applyOp } from "@cw/shared";
import { groupShapes, ungroupShape } from "./commands";
import type { Command, EditorState } from "./state";

const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };
const shape = (id: string, patch: Partial<ShapeBlock> = {}): ShapeBlock => ({
  kind: "shape", id, revision: 0, geometry: { preset: "rect" }, widthPx: 100, heightPx: 80, align: "left", ...patch,
});
const docOf = (...blocks: ShapeBlock[]): Document => ({ section: SECTION, blocks });
const shapeOf = (doc: Document, id: string): ShapeBlock | undefined => doc.blocks.find((b) => b.id === id) as ShapeBlock | undefined;
const applyAll = (doc: Document, cmd: Command, sel: EditorState["selection"] = null): { doc: Document; inverses: import("@cw/shared").Op[] } => {
  const trn = cmd({ doc, selection: sel });
  if (!trn) throw new Error("no transaction");
  let d = doc;
  const inverses: import("@cw/shared").Op[] = [];
  for (const op of trn.ops) {
    const res = applyOp(d, op);
    d = res.doc;
    inverses.push(res.inverse);
  }
  return { doc: d, inverses };
};
const undoAll = (doc: Document, inverses: import("@cw/shared").Op[]): Document => {
  let d = doc;
  for (let i = inverses.length - 1; i >= 0; i--) d = applyOp(d, inverses[i]!).doc;
  return d;
};

describe("groupShapes", () => {
  it("collapses two body shapes into one anchored group with an identity child map", () => {
    // a at page (200,300) 100×80 ; b at (350,420) 100×80. bbox = (200,300)..(450,500).
    const base = docOf(shape("a"), shape("b"));
    const { doc } = applyAll(base, groupShapes([{ blockId: "a", xPx: 200, yPx: 300 }, { blockId: "b", xPx: 350, yPx: 420 }], "g1"));
    expect(shapeOf(doc, "a")).toBeUndefined();
    expect(shapeOf(doc, "b")).toBeUndefined();
    const g = shapeOf(doc, "g1")!;
    expect(g.group).toBeDefined();
    expect(g.widthPx).toBe(250); // 450-200
    expect(g.heightPx).toBe(200); // 500-300
    expect(g.anchor).toEqual({ behind: false, offsetXPx: 200, offsetYPx: 300, relFromH: "page", relFromV: "page" });
    // Identity map: childExtent == group extent, childOffset 0.
    expect(g.group!.childOffsetXPx).toBe(0);
    expect(g.group!.childExtentXPx).toBe(250);
    expect(g.group!.childExtentYPx).toBe(200);
    // Children positioned relative to the bbox origin.
    expect(g.group!.children.map((c) => ({ x: c.xPx, y: c.yPx, id: c.shape.id }))).toEqual([
      { x: 0, y: 0, id: "a" },
      { x: 150, y: 120, id: "b" },
    ]);
  });

  it("strips a member's anchor/wrap when it becomes a child", () => {
    const base = docOf(shape("a", { anchor: { behind: true, offsetXPx: 10, offsetYPx: 20, relFromH: "margin", relFromV: "paragraph" } }), shape("b", { wrap: "square" }));
    const { doc } = applyAll(base, groupShapes([{ blockId: "a", xPx: 0, yPx: 0 }, { blockId: "b", xPx: 100, yPx: 0 }], "g1"));
    const g = shapeOf(doc, "g1")!;
    expect(g.group!.children[0]!.shape.anchor).toBeUndefined();
    expect(g.group!.children[1]!.shape.wrap).toBeUndefined();
  });

  it("group → ungroup is an exact round-trip (positions + sizes restored)", () => {
    const base = docOf(shape("a", { fill: { color: "#112233" } }), shape("b", { rotation: 30 }));
    const grouped = applyAll(base, groupShapes([{ blockId: "a", xPx: 200, yPx: 300 }, { blockId: "b", xPx: 350, yPx: 420 }], "g1"));
    // The host resolves the group's page origin (its anchor offset) for ungroup.
    const g = shapeOf(grouped.doc, "g1")!;
    const { doc } = applyAll(grouped.doc, ungroupShape("g1", g.anchor!.offsetXPx, g.anchor!.offsetYPx));
    expect(shapeOf(doc, "g1")).toBeUndefined();
    const a = shapeOf(doc, "a")!;
    const b = shapeOf(doc, "b")!;
    expect(a.widthPx).toBe(100);
    expect(a.heightPx).toBe(80);
    expect(a.anchor!.offsetXPx).toBe(200);
    expect(a.anchor!.offsetYPx).toBe(300);
    expect(a.fill).toEqual({ color: "#112233" });
    expect(b.anchor!.offsetXPx).toBe(350);
    expect(b.anchor!.offsetYPx).toBe(420);
    expect(b.rotation).toBe(30);
  });

  it("ungroup applies the group's child→box scale (imported non-identity group)", () => {
    // A group whose box is 2× its child coordinate space: children double in size.
    const child: ShapeBlock = shape("c", { widthPx: 50, heightPx: 40 });
    const grp = shape("g1", {
      widthPx: 200, heightPx: 160,
      anchor: { behind: false, offsetXPx: 100, offsetYPx: 100, relFromH: "page", relFromV: "page" },
      group: { children: [{ xPx: 10, yPx: 20, shape: child }], childOffsetXPx: 0, childOffsetYPx: 0, childExtentXPx: 100, childExtentYPx: 80 },
    });
    const base = docOf(grp);
    const { doc } = applyAll(base, ungroupShape("g1", 100, 100));
    const c = shapeOf(doc, "c")!;
    expect(c.widthPx).toBe(100); // 50 * (200/100)
    expect(c.heightPx).toBe(80); // 40 * (160/80)
    expect(c.anchor!.offsetXPx).toBe(120); // 100 + 10*2
    expect(c.anchor!.offsetYPx).toBe(140); // 100 + 20*2
  });

  it("is a single undo step (grouping is fully reversible)", () => {
    const base = docOf(shape("a"), shape("b"));
    const { doc, inverses } = applyAll(base, groupShapes([{ blockId: "a", xPx: 0, yPx: 0 }, { blockId: "b", xPx: 200, yPx: 0 }], "g1"));
    const undone = undoAll(doc, inverses);
    expect(undone.blocks.map((b) => b.id)).toEqual(["a", "b"]);
    expect(shapeOf(undone, "g1")).toBeUndefined();
  });

  it("guards: <2 members, duplicate ids, missing/cell-nested shapes all return null", () => {
    const base = docOf(shape("a"), shape("b"));
    expect(groupShapes([{ blockId: "a", xPx: 0, yPx: 0 }], "g1")({ doc: base, selection: null })).toBeNull();
    expect(groupShapes([{ blockId: "a", xPx: 0, yPx: 0 }, { blockId: "a", xPx: 1, yPx: 1 }], "g1")({ doc: base, selection: null })).toBeNull();
    expect(groupShapes([{ blockId: "a", xPx: 0, yPx: 0 }, { blockId: "nope", xPx: 1, yPx: 1 }], "g1")({ doc: base, selection: null })).toBeNull();
    const cellShape = shape("cellbox");
    const table = { kind: "table" as const, id: "tbl", revision: 0, rows: [{ cells: [{ id: "c0", blocks: [cellShape] }] }] };
    const withCell: Document = { section: SECTION, blocks: [shape("a"), table] };
    expect(groupShapes([{ blockId: "a", xPx: 0, yPx: 0 }, { blockId: "cellbox", xPx: 1, yPx: 1 }], "g1")({ doc: withCell, selection: null })).toBeNull();
  });
});

describe("ungroupShape", () => {
  it("returns null for a non-group shape or a missing id", () => {
    const base = docOf(shape("a"));
    expect(ungroupShape("a", 0, 0)({ doc: base, selection: null })).toBeNull();
    expect(ungroupShape("nope", 0, 0)({ doc: base, selection: null })).toBeNull();
  });
});
