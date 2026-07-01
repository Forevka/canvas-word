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
import { builtinsRegistered, registerFont } from "./fontRegistry";
import { ARABIC_FONT_FILE, CJK_FONT_FILE, FONT_FILES, HEBREW_FONT_FILE, MATH_FONT_FILE } from "../../fonts/clones";

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
    if (!builtinsRegistered()) {
      await Promise.all(FONT_FILES.map(async (f) => registerFont(f, await readFontBytes(f))));
      // The bundled math font (single Regular face) — registered alongside the
      // clones so equations measure + embed identically across editor/export.
      try {
        registerFont(MATH_FONT_FILE, await readFontBytes(MATH_FONT_FILE));
      } catch (e) {
        // Math font missing — equations fall back to the default clone (tofu for
        // math-alphanumeric glyphs), but the rest of the document is unaffected.
        // Surfaced (not silent) so a broken bundle is diagnosable.
        console.warn(`[wordcanvas] math font ${MATH_FONT_FILE} failed to load; equations will not typeset correctly`, e);
      }
      // The bundled CJK fallback (subset of Noto Sans SC) — registered alongside the
      // clones so script-split CJK runs measure + embed identically across editor/
      // export instead of rendering as .notdef/tofu.
      try {
        registerFont(CJK_FONT_FILE, await readFontBytes(CJK_FONT_FILE));
      } catch (e) {
        // CJK font missing — CJK text falls back to a Latin clone (tofu), but the
        // rest of the document is unaffected. Surfaced so a broken bundle is visible.
        console.warn(`[wordcanvas] CJK font ${CJK_FONT_FILE} failed to load; Chinese text will not render correctly`, e);
      }
      // The bundled Arabic fallback (Noto Sans Arabic) — registered so script-split
      // Arabic runs measure + embed identically, with correct joining-form shaping
      // via fontkit's GSUB, instead of rendering as .notdef/tofu.
      try {
        registerFont(ARABIC_FONT_FILE, await readFontBytes(ARABIC_FONT_FILE));
      } catch (e) {
        console.warn(`[wordcanvas] Arabic font ${ARABIC_FONT_FILE} failed to load; Arabic text will not render correctly`, e);
      }
      // The bundled Hebrew fallback (Noto Sans Hebrew) — registered so script-split
      // Hebrew runs measure + embed identically instead of rendering as .notdef/tofu.
      try {
        registerFont(HEBREW_FONT_FILE, await readFontBytes(HEBREW_FONT_FILE));
      } catch (e) {
        console.warn(`[wordcanvas] Hebrew font ${HEBREW_FONT_FILE} failed to load; Hebrew text will not render correctly`, e);
      }
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
