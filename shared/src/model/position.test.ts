import { describe, expect, it } from "vitest";
import type { DocSelection } from "./position";
import { isCollapsed } from "./position";

const sel = (a: { blockId: string; offset: number }, f: { blockId: string; offset: number }): DocSelection => ({
  anchor: a,
  focus: f,
});

describe("isCollapsed", () => {
  it("is true when anchor and focus share block and offset", () => {
    expect(isCollapsed(sel({ blockId: "p1", offset: 3 }, { blockId: "p1", offset: 3 }))).toBe(true);
  });

  it("is false when offsets differ in the same block", () => {
    expect(isCollapsed(sel({ blockId: "p1", offset: 0 }, { blockId: "p1", offset: 3 }))).toBe(false);
  });

  it("is false when blocks differ even at the same offset", () => {
    expect(isCollapsed(sel({ blockId: "p1", offset: 2 }, { blockId: "p2", offset: 2 }))).toBe(false);
  });

  it("ignores goalX when judging collapse", () => {
    expect(isCollapsed({ anchor: { blockId: "p1", offset: 5 }, focus: { blockId: "p1", offset: 5 }, goalX: 120 })).toBe(true);
  });
});
