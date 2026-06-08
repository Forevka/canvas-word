// M0 — the riskiest piece de-risked: prove the layout engine runs with NO DOM
// in Node, driven by the fontkit measure host, and that the host's widths agree
// with pdfkit's own measurement of the SAME embedded font (so a PDF's painted
// glyphs land where the engine broke its lines).

import PDFDocument from "pdfkit";
import { describe, expect, it, beforeAll } from "vitest";
import { sampleDoc } from "../../model/sampleDoc";
import { charStyleToFont, measureTextWidth } from "../../layout/metrics";
import { resolveFont } from "./fontRegistry";
import { installMeasureHost } from "./measureHost";

beforeAll(async () => {
  await installMeasureHost();
});

describe("measure host", () => {
  it("lays out the sample document in Node without any DOM", async () => {
    // Dynamic import AFTER the host is installed — the engine only measures at
    // layout() time, but this keeps the ordering contract explicit.
    const { createLayoutEngine } = await import("../../layout/engine");
    const tree = createLayoutEngine().layout(sampleDoc());
    expect(tree.pages.length).toBeGreaterThan(0);
    // Every page carries placed, positioned content.
    const placed = tree.pages.flatMap((p) => p.blocks);
    expect(placed.length).toBeGreaterThan(0);
    const withText = placed.some((b) => b.lines.some((l) => l.fragments.length > 0));
    expect(withText).toBe(true);
  });

  it("agrees with pdfkit.widthOfString for the same font and size", () => {
    const pdf = new PDFDocument();
    const cases = [
      { family: "Calibri", bold: false, italic: false, size: 16, text: "The quick brown fox" },
      { family: "Times New Roman", bold: true, italic: false, size: 24, text: "Heading Sample" },
      { family: "Arial", bold: false, italic: true, size: 11, text: "italic body text 12345" },
    ];
    for (const c of cases) {
      const resolved = resolveFont(c.family, c.bold, c.italic);
      pdf.registerFont(resolved.file, resolved.bytes as unknown as Buffer);
      const pdfWidth = pdf.font(resolved.file).fontSize(c.size).widthOfString(c.text);

      const font = charStyleToFont({
        fontFamily: `${c.family}, serif`,
        fontSizePx: c.size,
        bold: c.bold,
        italic: c.italic,
      } as never);
      const hostWidth = measureTextWidth(c.text, font);

      // Same fontkit advances on both sides → within sub-pixel rounding.
      expect(Math.abs(hostWidth - pdfWidth)).toBeLessThan(0.5);
    }
  });
});
