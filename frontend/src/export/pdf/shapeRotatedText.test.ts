// Issue #232 — a rotated shape's text box body must rotate WITH the geometry (Word
// rotates both). The PDF painter applies the same rotation-about-center transform to
// the text sub-flow that it applies to the fill/stroke, so a rotated text box shows
// tilted text instead of an upright label inside a tilted box. A recording pdfkit
// stand-in captures the rotate() calls so we can assert both the geometry and the
// text body rotate by the same angle about the same box centre.
import { describe, expect, it } from "vitest";
import type { PlacedBlock } from "../../layout/layoutTree";
import { paintBlock, type PaintCtx } from "./paintBlock";

/** A pdfkit stand-in that records every op (and the args to rotate). Any method is
 *  chainable and a no-op — only rotate is inspected. */
function makeRecordingDoc(): { doc: PaintCtx["doc"]; ops: string[]; rotations: { angle: unknown; origin: unknown }[] } {
  const ops: string[] = [];
  const rotations: { angle: unknown; origin: unknown }[] = [];
  const doc: unknown = new Proxy(
    {},
    {
      get(_t, prop) {
        return (...args: unknown[]) => {
          const name = String(prop);
          ops.push(name);
          if (name === "rotate") rotations.push({ angle: args[0], origin: (args[1] as { origin?: unknown } | undefined)?.origin });
          return doc;
        };
      },
    },
  );
  return { doc: doc as PaintCtx["doc"], ops, rotations };
}

/** A placed leaf shape rotated 30°, carrying a (minimal, empty-lines) text body. */
const rotatedTextShape = (): PlacedBlock =>
  ({
    blockId: "sh0",
    x: 100,
    y: 200,
    firstLineIndex: 0,
    lines: [],
    shape: {
      preset: "rect",
      width: 160,
      height: 80,
      rotation: 30,
      text: { offsetX: 6, offsetY: 6, blocks: [{ blockId: "t0", x: 0, y: 0, firstLineIndex: 0, lines: [] }] },
    },
  }) as unknown as PlacedBlock;

describe("PDF shape paint — rotated text box rotates WITH the geometry (#232)", () => {
  it("rotates both the geometry and the text body by the same angle about the box centre", () => {
    const { doc, ops, rotations } = makeRecordingDoc();
    const ctx: PaintCtx = { doc, font: () => "Arimo-Regular.ttf", image: () => new Uint8Array(), warn: () => {} };

    paintBlock(ctx, rotatedTextShape());

    // Two rotations: one for the fill/stroke geometry, one for the text body — both
    // 30° about the box centre (100 + 160/2, 200 + 80/2).
    expect(rotations.length).toBe(2);
    for (const r of rotations) {
      expect(r.angle).toBe(30);
      expect(r.origin).toEqual([180, 240]);
    }
    // The text clip never leaks: every save is matched by a restore.
    expect(ops.filter((o) => o === "save").length).toBe(ops.filter((o) => o === "restore").length);
  });
});
