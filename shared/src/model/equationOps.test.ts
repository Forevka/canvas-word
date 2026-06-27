// The setEquation op: replaces a display equation's MathML and is exactly
// invertible (undo restores the prior equation).

import { describe, expect, it } from "vitest";
import { applyOp } from "./ops";
import type { Document, EquationBlock } from "./document";
import type { MathEquation } from "./math";

const eq = (text: string): MathEquation => ({ root: { type: "row", children: [{ type: "number", text }] }, display: true });

const docWith = (equation: MathEquation): Document => ({
  section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
  blocks: [{ kind: "equation", id: "e1", revision: 0, equation, align: "center" } satisfies EquationBlock],
});

describe("setEquation op", () => {
  it("replaces the equation and inverts cleanly", () => {
    const doc = docWith(eq("1"));
    const r = applyOp(doc, { type: "setEquation", blockId: "e1", equation: eq("2") });
    const after = r.doc.blocks[0] as EquationBlock;
    expect(after.equation.root.children[0]).toMatchObject({ type: "number", text: "2" });
    expect(after.revision).toBe(1);

    // Inverse restores the original.
    const back = applyOp(r.doc, r.inverse).doc.blocks[0] as EquationBlock;
    expect(back.equation.root.children[0]).toMatchObject({ type: "number", text: "1" });
  });

  it("throws for a non-equation block id", () => {
    const doc = docWith(eq("1"));
    expect(() => applyOp(doc, { type: "setEquation", blockId: "missing", equation: eq("9") })).toThrow();
  });
});

describe("setEquationAlign op", () => {
  it("sets the alignment and inverts to the prior value", () => {
    const doc = docWith(eq("1")); // default align "center"
    const r = applyOp(doc, { type: "setEquationAlign", blockId: "e1", align: "right" });
    expect((r.doc.blocks[0] as EquationBlock).align).toBe("right");
    const back = applyOp(r.doc, r.inverse).doc.blocks[0] as EquationBlock;
    expect(back.align).toBe("center");
  });
});
