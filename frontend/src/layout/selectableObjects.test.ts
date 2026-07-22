// E3 — listSelectableObjects enumerates every object-selectable block (image /
// shape / equation / custom) in page + reading order, the order the keyboard
// Tab / F6 object-focus cycle steps through.
//
// geometry.ts creates a measuring <canvas> at module load, so we stub a minimal
// `document` (this repo's vitest runs in node, no jsdom) and dynamic-import after.

import { beforeAll, describe, expect, it } from "vitest";
import type { Document, ImageBlock, Paragraph, ShapeBlock } from "@cw/shared";

beforeAll(() => {
  (globalThis as unknown as { document: unknown }).document = {
    createElement: () => ({ getContext: () => ({ measureText: () => ({ width: 0 }), font: "" }) }),
  };
});

const PARA = { align: "left" as const, lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const CHAR = { fontFamily: "Arial", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" };
const para = (id: string, text: string): Paragraph => ({
  kind: "paragraph", id, revision: 0, runs: [{ text, style: { ...CHAR } }], style: { ...PARA },
});
const image = (id: string): ImageBlock => ({
  kind: "image", id, revision: 0, src: "blob:test", widthPx: 120, heightPx: 80, align: "center",
});
const shape = (id: string): ShapeBlock => ({
  kind: "shape", id, revision: 0, geometry: { preset: "rect" }, widthPx: 120, heightPx: 80, align: "left",
});

describe("listSelectableObjects", () => {
  it("returns image/shape/equation ids in document order, excluding paragraphs", async () => {
    const { installMeasureHost } = await import("../export/shared/measureHost");
    const { createLayoutEngine } = await import("./engine");
    const { listSelectableObjects } = await import("./geometry");
    const { parseMathml } = await import("../mathml/parse");
    await installMeasureHost();

    const eq = { kind: "equation" as const, id: "eq1", revision: 0, align: "center" as const,
      equation: { ...parseMathml("<math><mn>1</mn></math>"), display: true } };
    const doc: Document = {
      section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [para("p0", "before"), image("img1"), shape("shp1"), eq, para("p1", "after")],
    };
    const tree = createLayoutEngine().layout(doc);
    expect(listSelectableObjects(tree)).toEqual(["img1", "shp1", "eq1"]);
  });

  it("is empty for a document with no objects", async () => {
    const { installMeasureHost } = await import("../export/shared/measureHost");
    const { createLayoutEngine } = await import("./engine");
    const { listSelectableObjects } = await import("./geometry");
    await installMeasureHost();

    const doc: Document = {
      section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
      blocks: [para("p0", "just text")],
    };
    expect(listSelectableObjects(createLayoutEngine().layout(doc))).toEqual([]);
  });
});
