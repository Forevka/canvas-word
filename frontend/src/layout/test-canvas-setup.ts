// Deterministic canvas measurement for engine tests under vitest (node, no DOM).
// pretext (rich-inline) and metrics.ts both measure text through a 2d canvas
// context — pretext via OffscreenCanvas, metrics via document.createElement.
// We give them length-proportional widths (chars × ½ font-size) so layout is
// deterministic and assertable without a real font engine.
//
// IMPORT THIS FIRST in any engine test, before the engine itself: ESM evaluates
// imports in order, so the globals are installed before metrics.ts (which makes
// its context at module-load time) runs.

const sizeOf = (font: string): number => {
  const m = /(\d+(?:\.\d+)?)px/.exec(font);
  return m ? parseFloat(m[1]!) : 16;
};

const makeCtx = (): unknown => {
  let font = "16px sans-serif";
  return {
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    measureText(text: string) {
      const size = sizeOf(font);
      // 1F600 (pretext's emoji probe) is two UTF-16 units → width === size,
      // which keeps `width > fontSize + 0.5` false so the DOM-correction path
      // (which would need document.body) never runs.
      return {
        width: text.length * size * 0.5,
        fontBoundingBoxAscent: size * 0.8,
        fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8,
        actualBoundingBoxDescent: size * 0.2,
      };
    },
  };
};

const g = globalThis as unknown as {
  OffscreenCanvas?: unknown;
  document?: unknown;
};

if (g.OffscreenCanvas === undefined) {
  g.OffscreenCanvas = class {
    constructor(_w: number, _h: number) {}
    getContext(): unknown {
      return makeCtx();
    }
  };
}

if (g.document === undefined) {
  g.document = {
    createElement: () => ({ getContext: () => makeCtx() }),
    body: null,
  };
}
