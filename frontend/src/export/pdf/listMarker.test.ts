// Issue #99: nested bullet markers (◦ U+25E6, ▪ U+25AA) were painted as glyphs in
// the PDF, but those code points are absent from the bundled Latin font subset, so
// they rendered as .notdef/tofu. The PDF painter now draws the three default
// bullets as vector shapes (disc / ring / square). This test drives paintBlock with
// a recording fake PDFKit document and asserts a bullet marker emits vector ops
// (never glyph text), while a numbered marker still paints as text.
import { describe, expect, it } from "vitest";
import type { CharStyle } from "@cw/shared";
import type { PlacedBlock } from "../../layout/layoutTree";
import { paintBlock, type PaintCtx } from "./paintBlock";

const CHAR: CharStyle = {
  fontFamily: "Calibri, sans-serif",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#202124",
};

/** A chainable fake of the PDFKit drawing surface that records the method names
 *  paintBlock invokes (every method returns `this` so chains like
 *  `.circle(...).fill(...)` work). */
function recordingDoc(): { doc: PDFKit.PDFDocument; ops: string[] } {
  const ops: string[] = [];
  const handler: ProxyHandler<object> = {
    get(_t, prop) {
      return (..._args: unknown[]) => {
        ops.push(String(prop));
        return proxy;
      };
    },
  };
  const proxy = new Proxy({}, handler) as unknown as PDFKit.PDFDocument;
  return { doc: proxy, ops };
}

function markerBlock(text: string): PlacedBlock {
  return {
    blockId: "b",
    x: 50,
    y: 100,
    firstLineIndex: 0,
    lines: [{ y: 0, height: 20, ascent: 16, fragments: [] }],
    marker: { text, style: CHAR, x: 40 },
  };
}

function paint(text: string): string[] {
  const { doc, ops } = recordingDoc();
  const ctx: PaintCtx = { doc, font: () => "F", image: () => undefined, warn: () => {} };
  paintBlock(ctx, markerBlock(text));
  return ops;
}

describe("PDF list marker — default bullets paint as vector shapes (issue #99)", () => {
  it("draws • as a filled circle, not glyph text", () => {
    const ops = paint("•");
    expect(ops).toContain("circle");
    expect(ops).toContain("fill");
    expect(ops).not.toContain("text");
  });

  it("draws ◦ (U+25E6, missing from the font subset) as a stroked ring, not tofu", () => {
    const ops = paint("◦");
    expect(ops).toContain("circle");
    expect(ops).toContain("stroke");
    expect(ops).not.toContain("text");
  });

  it("draws ▪ (U+25AA, missing from the font subset) as a filled square, not tofu", () => {
    const ops = paint("▪");
    expect(ops).toContain("rect");
    expect(ops).toContain("fill");
    expect(ops).not.toContain("text");
  });

  it("still paints a numbered marker as glyph text (no vector shape)", () => {
    const ops = paint("1.");
    expect(ops).toContain("text");
    expect(ops).not.toContain("circle");
  });
});
