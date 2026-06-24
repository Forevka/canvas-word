// Regression: registering custom fonts BEFORE installMeasureHost (the worker order
// — runExport registers custom faces, then renderPdf installs the host) must NOT
// trick the host into skipping the bundled-font load. Custom bytes populate the
// fontkit registry, so a naive "any font loaded?" guard would skip the clones and
// every built-in family would fail to resolve. The guard keys on the bundled Arimo
// sentinel (builtinsRegistered) instead. This file has its own module registry
// (vitest isolates per file), so the order here is the real first-load order.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { registerCustomFonts } from "../../fonts/customRegistry";
import { registerCustomFontBytes, resolveFont } from "./fontRegistry";
import { installMeasureHost } from "./measureHost";

describe("measure host load order", () => {
  it("loads bundled fonts even when custom fonts were registered first", async () => {
    const bytes = new Uint8Array(readFileSync(new URL("./fonts/Arimo-Regular.ttf", import.meta.url)));
    // Custom first — this is what runExport does before renderPdf installs the host.
    registerCustomFonts({ fonts: [{ family: "Brandish", faces: { regular: "x" }, sizing: { ascent: 0.9, descent: 0.25 } }] });
    registerCustomFontBytes([{ family: "Brandish", style: "Regular", bytes }]);

    await installMeasureHost();

    // A built-in family must still resolve (bundled clones were loaded).
    const arial = resolveFont("Arial", false, false);
    expect(arial.file).toBe("Arimo-Regular.ttf");
    // And the custom family resolves to its own face.
    const custom = resolveFont("Brandish", false, false);
    expect(custom.file).toBe("__custom__brandish-Regular.ttf");
    expect(custom.substituted).toBe(false);
  });
});
