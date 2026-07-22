// E4 (docs/UX_ANALYSIS.md) — the shape-scoped accessible-name + announcement
// helpers. Pure string builders; no DOM, so unit-tested directly.

import { describe, it, expect } from "vitest";
import type { Paragraph, ShapeBlock } from "@cw/shared";
import {
  shapeGeometryLabel,
  shapeText,
  shapeAccessibleName,
  describeShapeSelected,
  describeShapeMoved,
  describeShapeRotated,
} from "./shapeAnnounce";

const CHAR = { fontFamily: "Arial", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" };
const PARA = { align: "left" as const, lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const para = (text: string): Paragraph => ({
  kind: "paragraph",
  id: "p",
  revision: 0,
  runs: text ? [{ text, style: { ...CHAR } }] : [],
  style: { ...PARA },
});
const shape = (patch: Partial<ShapeBlock> = {}): ShapeBlock => ({
  kind: "shape",
  id: "s",
  revision: 0,
  geometry: { preset: "rect" },
  widthPx: 180,
  heightPx: 120,
  align: "left",
  ...patch,
});

describe("shapeGeometryLabel", () => {
  it("maps presets to friendly names", () => {
    expect(shapeGeometryLabel(shape({ geometry: { preset: "rect" } }))).toBe("Rectangle");
    expect(shapeGeometryLabel(shape({ geometry: { preset: "roundRect" } }))).toBe("Rounded rectangle");
    expect(shapeGeometryLabel(shape({ geometry: { preset: "rightArrow" } }))).toBe("Right arrow");
    expect(shapeGeometryLabel(shape({ geometry: { preset: "line" } }))).toBe("Line");
  });

  it("reports a group by its member count (singular/plural)", () => {
    const child = (): { xPx: number; yPx: number; shape: ShapeBlock } => ({ xPx: 0, yPx: 0, shape: shape() });
    const grp = (n: number): ShapeBlock =>
      shape({
        group: { children: Array.from({ length: n }, child), childOffsetXPx: 0, childOffsetYPx: 0, childExtentXPx: 100, childExtentYPx: 100 },
      });
    expect(shapeGeometryLabel(grp(1))).toBe("Group of 1 shape");
    expect(shapeGeometryLabel(grp(3))).toBe("Group of 3 shapes");
  });

  it("calls a freeform custom geometry a generic Shape", () => {
    expect(shapeGeometryLabel(shape({ geometry: { preset: "rect", custom: { segments: [] } } }))).toBe("Shape");
  });
});

describe("shapeText", () => {
  it("is empty for a text-less shape", () => {
    expect(shapeText(shape())).toBe("");
  });

  it("flattens paragraphs and collapses whitespace", () => {
    const s = shape({ text: { blocks: [para("Hello"), para("  world  ")] } });
    expect(shapeText(s)).toBe("Hello world");
  });

  it("treats a whitespace-only body as empty", () => {
    expect(shapeText(shape({ text: { blocks: [para("   ")] } }))).toBe("");
  });
});

describe("shapeAccessibleName", () => {
  it("is just the label with no text", () => {
    expect(shapeAccessibleName(shape())).toBe("Rectangle");
  });

  it("appends the text when present", () => {
    expect(shapeAccessibleName(shape({ text: { blocks: [para("Title")] } }))).toBe("Rectangle: Title");
  });

  it("truncates a long text body", () => {
    const long = "a".repeat(200);
    const name = shapeAccessibleName(shape({ text: { blocks: [para(long)] } }));
    expect(name.startsWith("Rectangle: ")).toBe(true);
    expect(name.endsWith("…")).toBe(true);
    expect(name.length).toBeLessThan(long.length);
  });
});

describe("announcements", () => {
  it("announces selection with the accessible name", () => {
    expect(describeShapeSelected(shape({ text: { blocks: [para("Note")] } }))).toBe("Rectangle: Note, selected");
  });

  it("announces a move with rounded offsets", () => {
    expect(describeShapeMoved(shape(), 120.4, 79.6)).toBe("Rectangle moved to 120, 80 pixels");
  });

  it("announces a rotation, and rotation removal at 0", () => {
    expect(describeShapeRotated(shape(), 45)).toBe("Rectangle rotated to 45 degrees");
    expect(describeShapeRotated(shape(), 0)).toBe("Rectangle, rotation removed");
  });
});
