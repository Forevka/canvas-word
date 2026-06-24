// Custom fonts must behave EXACTLY like the bundled clones at export time: the
// fontkit measure host and pdfkit must agree on widths for the custom face (so its
// glyphs land where the engine broke lines), a missing optional style must fall
// back to the custom Regular (deterministic parity with the editor), and line-height
// must come from the caller-supplied sizing. We register a custom family backed by a
// real bundled TTF so the bytes are valid.

import { readFileSync } from "node:fs";
import PDFDocument from "pdfkit";
import { beforeAll, describe, expect, it } from "vitest";
import { charStyleToFont, fontMetrics, measureTextWidth } from "../../layout/metrics";
import {
  __resetCustomFonts,
  customFontFileName,
  registerCustomFonts,
  type CustomFontDef,
} from "../../fonts/customRegistry";
import { registerCustomFontBytes, resolveFont } from "./fontRegistry";
import { installMeasureHost } from "./measureHost";

const FAMILY = "FakeBrand";
const SIZING = { ascent: 0.95, descent: 0.3 };
const def: CustomFontDef = {
  family: FAMILY,
  faces: { regular: "https://x/FakeBrand-Regular.ttf" }, // regular only — no bold
  sizing: SIZING,
};

let regularBytes: Uint8Array;

beforeAll(async () => {
  await installMeasureHost();
  regularBytes = new Uint8Array(readFileSync(new URL("./fonts/Arimo-Regular.ttf", import.meta.url)));
  __resetCustomFonts();
  registerCustomFonts({ fonts: [def] });
  registerCustomFontBytes([{ family: FAMILY, style: "Regular", bytes: regularBytes }]);
});

describe("custom font export parity", () => {
  it("resolves to its synthetic file and embeds without substitution", () => {
    const r = resolveFont(FAMILY, false, false);
    expect(r.file).toBe(customFontFileName(FAMILY, "Regular"));
    expect(r.substituted).toBe(false);
    expect(r.bytes.byteLength).toBe(regularBytes.byteLength);
  });

  it("falls back to the custom Regular for a missing bold face (not Arimo)", () => {
    const r = resolveFont(FAMILY, true, false);
    expect(r.file).toBe(customFontFileName(FAMILY, "Regular"));
    expect(r.substituted).toBe(false);
  });

  it("a registered custom family with NO loaded bytes resolves to Arimo, marked substituted", () => {
    // Registered as a def (so cloneFamilyFor keeps it custom) but its bytes never
    // loaded — e.g. the required Regular face failed to fetch. resolveFont must NOT
    // emit Arimo bytes under a custom file name with substituted:false (a metadata
    // lie); it falls back honestly so a font-substituted warning can fire.
    registerCustomFonts({
      fonts: [{ family: "GhostBrand", faces: { regular: "https://x/Ghost.ttf" }, sizing: { ascent: 0.9, descent: 0.25 } }],
    });
    const r = resolveFont("GhostBrand", false, false);
    expect(r.file).toBe("Arimo-Regular.ttf");
    expect(r.substituted).toBe(true);
  });

  it("measure host agrees with pdfkit.widthOfString for the custom face", () => {
    const pdf = new PDFDocument();
    const resolved = resolveFont(FAMILY, false, false);
    pdf.registerFont(resolved.file, resolved.bytes as unknown as Buffer);
    const pdfWidth = pdf.font(resolved.file).fontSize(16).widthOfString("The quick brown fox");

    const font = charStyleToFont({ fontFamily: FAMILY, fontSizePx: 16, bold: false, italic: false } as never);
    const hostWidth = measureTextWidth("The quick brown fox", font);

    expect(Math.abs(hostWidth - pdfWidth)).toBeLessThan(0.5);
  });

  it("line-height metrics come from the supplied sizing", () => {
    const font = charStyleToFont({ fontFamily: FAMILY, fontSizePx: 20, bold: false, italic: false } as never);
    expect(fontMetrics(font)).toEqual({
      ascent: Math.round(SIZING.ascent * 20),
      descent: Math.round(SIZING.descent * 20),
    });
  });
});
