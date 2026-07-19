// Regression: on the browser main thread (the inline path taken for PDF export
// of a REGISTERED custom block) and inside the StackBlitz WebContainer worker,
// `globalThis.document` is a native read-only accessor with no setter. The
// measure host must NOT throw "Cannot set property document of #<Window> which
// has only a getter" when it installs the DOM-free measurement shim — it should
// tolerate the read-only document and leave it in place (pretext still measures
// widths through the fontkit OffscreenCanvas shim).

import { describe, expect, it, afterEach } from "vitest";

describe("installMeasureHost — read-only document", () => {
  afterEach(() => {
    delete (globalThis as { document?: unknown }).document;
    delete (globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas;
  });

  it("resolves without throwing when document is a getter-only accessor", async () => {
    // Simulate window.document: an accessor with a getter and NO setter, so any
    // assignment throws in strict mode (ES modules are always strict).
    const realDocument = { createElement: () => ({ getContext: () => null }), body: {} };
    Object.defineProperty(globalThis, "document", {
      configurable: true,
      get: () => realDocument,
    });

    // Import AFTER defining document so the module's singleton install runs against it.
    const { installMeasureHost } = await import("./measureHost");
    await expect(installMeasureHost()).resolves.toBeUndefined();

    // The real (read-only) document was left untouched — not clobbered by the stub.
    expect((globalThis as { document?: unknown }).document).toBe(realDocument);
    // The fontkit OffscreenCanvas shim (pretext's width source) still installed.
    expect(typeof (globalThis as { OffscreenCanvas?: unknown }).OffscreenCanvas).toBe("function");
  });
});
