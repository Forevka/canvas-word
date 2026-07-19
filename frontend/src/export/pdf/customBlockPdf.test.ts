// Integration test: a REGISTERED custom block renders into a real PDF via the
// Canvas2D→pdfkit shim (this is exactly the headless path — Node/ClearScript run
// the same renderPdf). Uses real pdfkit + the bundled fonts.
import { describe, expect, it, beforeAll, afterEach } from "vitest";
import type { CustomBlock, Document } from "@cw/shared";
import { installMeasureHost } from "../shared/measureHost";
import { renderPdf } from "./renderPdf";
import { registerBlockType, _clearBlockTypes } from "../../blockRegistry";

beforeAll(async () => {
  await installMeasureHost();
});
afterEach(() => _clearBlockTypes());

const customDoc = (data: unknown = { title: "Q" }): Document => ({
  section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } },
  blocks: [{ kind: "custom", id: "c1", revision: 0, customType: "chart", data } as CustomBlock],
});

const isPdf = (b: Uint8Array): boolean => b[0] === 0x25 && b[1] === 0x50 && b[2] === 0x44 && b[3] === 0x46;
const unrendered = (ws: { code: string }[]): boolean => ws.some((w) => w.code === "custom-block-not-rendered");

describe("custom block PDF export (vector shim)", () => {
  it("renders a REGISTERED block — valid PDF, no 'not-rendered' warning, larger than the placeholder", async () => {
    let painted = false;
    registerBlockType({
      type: "chart",
      measure: () => ({ height: 120 }),
      paint: (ctx, box) => {
        painted = true;
        ctx.fillStyle = "#eef2ff";
        ctx.fillRect(0, 0, box.width, box.height);
        ctx.strokeStyle = "#c7d2fe";
        ctx.strokeRect(0.5, 0.5, box.width - 1, box.height - 1);
        ctx.font = "bold 14px Arial";
        ctx.fillStyle = "#1e3a8a";
        ctx.fillText("Quarterly", 12, 22);
        ctx.beginPath();
        ctx.arc(40, 60, 12, 0, Math.PI * 2);
        ctx.fillStyle = "#2563eb";
        ctx.fill();
        ctx.fillRect(80, 40, 30, 60); // a "bar"
      },
    });
    const rendered = await renderPdf(customDoc());
    expect(painted).toBe(true); // paint() actually ran during export
    expect(isPdf(rendered.bytes)).toBe(true);
    expect(unrendered(rendered.warnings)).toBe(false);

    // Compare against the same doc with the type NOT registered (placeholder only).
    _clearBlockTypes();
    const placeholder = await renderPdf(customDoc());
    expect(unrendered(placeholder.warnings)).toBe(true);
    // The rendered block adds vector content, so its PDF is larger.
    expect(rendered.bytes.length).toBeGreaterThan(placeholder.bytes.length);
  });

  it("reserves + warns for an UNREGISTERED custom type", async () => {
    const { bytes, warnings } = await renderPdf(customDoc());
    expect(isPdf(bytes)).toBe(true);
    expect(unrendered(warnings)).toBe(true);
  });

  it("a pdf:'raster' block reserves + warns until the raster path ships", async () => {
    registerBlockType({ type: "chart", pdf: "raster", measure: () => ({ height: 100 }), paint: () => {} });
    const { warnings } = await renderPdf(customDoc());
    expect(unrendered(warnings)).toBe(true);
  });

  it("a throwing paint() doesn't break the export (valid PDF still produced)", async () => {
    registerBlockType({ type: "chart", measure: () => ({ height: 100 }), paint: () => { throw new Error("bad paint"); } });
    const { bytes, warnings } = await renderPdf(customDoc());
    expect(isPdf(bytes)).toBe(true);
    expect(warnings.some((w) => w.code === "custom-block-paint-failed")).toBe(true);
  });
});
