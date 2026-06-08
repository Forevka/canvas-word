// Installs DOM-free text measurement for the layout engine so it runs in the
// export worker and on Node backends. Measurement always goes through the
// fontkit-backed shim over the BUNDLED CLONE FONTS — the exact same path in the
// browser worker and on Node — so an exported PDF is byte-identical across
// environments (no system fonts anywhere). The editor renders the same clones
// (via charStyleToFont + FontFace, see src/export/shared/editorFonts.ts), so the
// on-screen document and the export agree too.
//
// After this resolves, BOTH measurement paths use the shim:
//   - pretext (rich-inline): via globalThis.OffscreenCanvas / document
//   - metrics.ts: via setMeasureContext
// MUST be awaited before the engine measures any text.

import { setMeasureContext } from "../../layout/metrics";
import { FontkitMeasureContext } from "./fontkitContext";
import { fontsLoaded, registerFont } from "./fontRegistry";
import { FONT_FILES } from "../../fonts/clones";

async function readFontBytes(file: string): Promise<Uint8Array> {
  const url = new URL(`./fonts/${file}`, import.meta.url);
  // Prefer Node's fs (backend + vitest); in the browser the node: import is
  // externalized to a stub whose readFile is undefined, so we fall through to
  // fetch (the bundler emits the font as an asset).
  try {
    const fs = await import(/* @vite-ignore */ "node:fs/promises");
    if (typeof fs?.readFile === "function") return new Uint8Array(await fs.readFile(url));
  } catch {
    // not Node — fall back to fetch
  }
  const res = await fetch(url);
  return new Uint8Array(await res.arrayBuffer());
}

let installPromise: Promise<void> | null = null;

export function installMeasureHost(): Promise<void> {
  if (installPromise) return installPromise;
  installPromise = (async () => {
    if (!fontsLoaded()) {
      await Promise.all(FONT_FILES.map(async (f) => registerFont(f, await readFontBytes(f))));
    }
    // Route pretext + metrics through the fontkit shim over the bundled clones.
    // Override the globals unconditionally — this host runs only off the main
    // thread (export worker / Node), never where it could clobber the editor.
    const g = globalThis as unknown as { OffscreenCanvas?: unknown; document?: unknown };
    g.OffscreenCanvas = class {
      constructor(_w?: number, _h?: number) {}
      getContext(): unknown {
        return new FontkitMeasureContext();
      }
    };
    // body:null disables pretext's emoji DOM-correction probe; createElement
    // covers pretext's DOM fallback.
    g.document = { createElement: () => ({ getContext: () => new FontkitMeasureContext() }), body: null };
    setMeasureContext(new FontkitMeasureContext());
  })();
  return installPromise;
}
