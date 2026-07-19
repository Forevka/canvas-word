// The canvas stub (deterministic text metrics) must load before the engine.
import "./layout/test-canvas-setup";

import { describe, it, expect, beforeEach } from "vitest";
import type { Block, CustomBlock, Document, Paragraph, SectionProps } from "@cw/shared";
import { registerBlockType, getBlockType, hasBlockType, _clearBlockTypes } from "./blockRegistry";
import { createLayoutEngine, CUSTOM_BLOCK_PLACEHOLDER_H } from "./layout/engine";
import type { PlacedBlock } from "./layout/layoutTree";

const SECTION: SectionProps = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
};

let n = 0;
const customBlock = (customType: string, data: unknown = {}): CustomBlock => ({
  kind: "custom",
  id: `c${n++}`,
  revision: 0,
  customType,
  data,
});

const doc = (blocks: Block[]): Document => ({ section: SECTION, blocks });

const find = (tree: ReturnType<ReturnType<typeof createLayoutEngine>["layout"]>, id: string): PlacedBlock | undefined => {
  for (const pg of tree.pages) for (const pb of pg.blocks) if (pb.blockId === id) return pb;
  return undefined;
};

beforeEach(() => _clearBlockTypes());

describe("block registry (pure)", () => {
  it("registers, looks up, and reports presence", () => {
    expect(hasBlockType("x")).toBe(false);
    const def = { type: "x", measure: () => ({ height: 10 }), paint: () => {} };
    registerBlockType(def);
    expect(hasBlockType("x")).toBe(true);
    expect(getBlockType("x")).toBe(def);
  });

  it("dispose unregisters", () => {
    const off = registerBlockType({ type: "x", measure: () => ({ height: 10 }), paint: () => {} });
    off();
    expect(hasBlockType("x")).toBe(false);
  });

  it("dispose is a no-op after the type was overwritten", () => {
    const off1 = registerBlockType({ type: "x", measure: () => ({ height: 1 }), paint: () => {} });
    const def2 = { type: "x", measure: () => ({ height: 2 }), paint: () => {} };
    registerBlockType(def2); // overwrites
    off1(); // must NOT remove def2
    expect(getBlockType("x")).toBe(def2);
  });
});

describe("block registry (layout integration)", () => {
  it("measures + places a registered custom block as an atomic block", () => {
    registerBlockType({ type: "box", measure: (d) => ({ height: (d as { h: number }).h }), paint: () => {} });
    const cb = customBlock("box", { h: 120 });
    const tree = createLayoutEngine().layout(doc([cb]));
    const placed = find(tree, cb.id);
    expect(placed).toBeDefined();
    expect(placed!.custom).toBeDefined();
    expect(placed!.custom!.customType).toBe("box");
    expect(placed!.custom!.height).toBe(120);
    expect(placed!.custom!.width).toBe(624); // 816 - 96*2 content width
    expect(placed!.lines).toHaveLength(0); // atomic — carries no lines (no caret)
  });

  it("measure receives the content width", () => {
    let seenWidth = -1;
    registerBlockType({
      type: "wprobe",
      measure: (_d, ctx) => { seenWidth = ctx.width; return { height: 20 }; },
      paint: () => {},
    });
    createLayoutEngine().layout(doc([customBlock("wprobe")]));
    expect(seenWidth).toBe(624);
  });

  it("falls back to a placeholder height when the type isn't registered", () => {
    const cb = customBlock("missing");
    const tree = createLayoutEngine().layout(doc([cb]));
    const placed = find(tree, cb.id);
    expect(placed!.custom!.height).toBe(CUSTOM_BLOCK_PLACEHOLDER_H);
  });

  it("a throwing measure() falls back to the placeholder height (never crashes layout)", () => {
    registerBlockType({ type: "boom", measure: () => { throw new Error("bad measure"); }, paint: () => {} });
    const cb = customBlock("boom");
    // Must not throw — the engine catches it.
    const tree = createLayoutEngine().layout(doc([cb]));
    expect(find(tree, cb.id)!.custom!.height).toBe(CUSTOM_BLOCK_PLACEHOLDER_H);
  });

  it("a following paragraph is pushed below the custom block", () => {
    registerBlockType({ type: "tall", measure: () => ({ height: 300 }), paint: () => {} });
    const cb = customBlock("tall");
    const para: Paragraph = {
      kind: "paragraph",
      id: "p1",
      revision: 0,
      runs: [{ text: "after", style: { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" } }],
      style: { align: "left", lineHeight: 1, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 },
    };
    const tree = createLayoutEngine().layout(doc([cb, para]));
    const c = find(tree, cb.id)!;
    const p = find(tree, "p1")!;
    expect(p.y).toBeGreaterThanOrEqual(c.y + c.custom!.height);
  });
});
