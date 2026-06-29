// Case transforms (w:caps / w:smallCaps) are baked into the runs the engine
// measures and paints: letters render UPPERCASED (caps) and the originally-
// lowercase letters render smaller (smallCaps), while the model offsets stay 1:1
// so the caret/measurement land on the transformed glyphs. The deterministic
// canvas stub measures width = len x 1/2 size, so a reduced-size fragment is both
// narrower and carries a smaller fontSizePx.
import "./test-canvas-setup";

import { describe, it, expect } from "vitest";
import type { CharStyle, Document, Paragraph, ParaStyle, SectionProps } from "@cw/shared";
import { createLayoutEngine } from "./engine";
import { SMALL_CAPS_SCALE } from "../paint/paintStyle";
import type { InlineFragment, PlacedBlock } from "./layoutTree";

const CHAR: CharStyle = {
  fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false,
  underline: false, strikethrough: false, color: "#000000",
};
const PARA: ParaStyle = {
  align: "left", lineHeight: 1, spaceBeforePx: 0, spaceAfterPx: 0,
  indentFirstLinePx: 0, indentLeftPx: 0,
};
const SECTION: SectionProps = {
  pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
};

const para = (runs: { text: string; style: Partial<CharStyle> }[]): Paragraph => ({
  kind: "paragraph", id: "p1", revision: 0, style: { ...PARA },
  runs: runs.map((r) => ({ text: r.text, style: { ...CHAR, ...r.style } })),
});

const fragsOf = (p: Paragraph): InlineFragment[] => {
  const doc: Document = { section: SECTION, blocks: [p] };
  const tree = createLayoutEngine().layout(doc);
  let pb: PlacedBlock | undefined;
  for (const pg of tree.pages) for (const b of pg.blocks) if (b.blockId === p.id) pb = b;
  return pb!.lines.flatMap((l) => l.fragments);
};

describe("caps / smallCaps layout transforms", () => {
  it("caps uppercases the painted glyphs while preserving model offsets", () => {
    const frags = fragsOf(para([{ text: "Hello", style: { caps: true } }]));
    expect(frags).toHaveLength(1);
    const f = frags[0]!;
    expect(f.text).toBe("HELLO"); // painters draw the uppercased glyphs
    expect(f.style.fontSizePx).toBe(16); // caps does NOT shrink anything
    expect(f.style.caps).toBeUndefined(); // flag baked away on the display run
    // Offset-transparent: the model run "Hello" is 5 UTF-16 units, so the
    // fragment still spans offsets 0..5 (caret/measurement stay aligned).
    expect(f.startOffset).toBe(0);
    expect(f.endOffset).toBe(5);
  });

  it("smallCaps uppercases all letters and shrinks the originally-lowercase ones", () => {
    const frags = fragsOf(para([{ text: "Ab", style: { smallCaps: true } }]));
    // "A" (already upper) keeps full size; "b" (lowercase) is uppercased + reduced.
    const full = frags.find((f) => f.text === "A")!;
    const small = frags.find((f) => f.text === "B")!;
    expect(full).toBeTruthy();
    expect(small).toBeTruthy();
    expect(full.style.fontSizePx).toBe(16);
    expect(small.style.fontSizePx).toBeCloseTo(16 * SMALL_CAPS_SCALE, 2);
    // Narrower because measured at the reduced size (stub: width = len x 1/2 size).
    expect(small.width).toBeLessThan(full.width);
    // Offsets remain 1:1 with the model "Ab": "A" -> 0..1, "B" -> 1..2.
    expect(full.startOffset).toBe(0);
    expect(full.endOffset).toBe(1);
    expect(small.startOffset).toBe(1);
    expect(small.endOffset).toBe(2);
    expect(small.style.smallCaps).toBeUndefined();
  });

  it("keeps a combining mark attached to its reduced small-caps base glyph", () => {
    // Decomposed base+mark built from char codes (source stays ASCII): "a" + the
    // combining diaeresis U+0308, then "b". All reduce into ONE fragment; the mark
    // must ride with its uppercased base ("A" + U+0308 + "B"), never split off into
    // a separate full-size fragment (which would mis-shape the accent).
    const mark = String.fromCharCode(0x0308);
    const input = "a" + mark + "b";
    const expected = input.toUpperCase(); // "A" + U+0308 + "B"
    const frags = fragsOf(para([{ text: input, style: { smallCaps: true } }]));
    const reduced = frags.find((fr) => fr.text === expected);
    expect(reduced).toBeTruthy();
    expect(reduced!.style.fontSizePx).toBeCloseTo(16 * SMALL_CAPS_SCALE, 2);
    expect(reduced!.startOffset).toBe(0);
    expect(reduced!.endOffset).toBe(3); // a, U+0308, b -> 3 UTF-16 units preserved
  });

  it("leaves runs without case flags untouched", () => {
    const frags = fragsOf(para([{ text: "Plain", style: {} }]));
    expect(frags).toHaveLength(1);
    expect(frags[0]!.text).toBe("Plain");
  });
});
