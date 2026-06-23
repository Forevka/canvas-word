// Font registry for export — loads the bundled metric clones and resolves a
// requested family to a parsed fontkit face. The SAME face both measures (layout
// widths) and is embedded into the PDF, so painted glyphs land where measured.
// Family→clone mapping is shared with the editor via src/fonts/clones.ts.

import * as fontkit from "fontkit";
import type { Font } from "fontkit";
import { cloneFamilyFor, FONT_FILES } from "../../fonts/clones";
import { customFontFileName, customFontFor, type CustomFontFaceBytes } from "../../fonts/customRegistry";

export { FONT_FILES };

interface LoadedFont {
  bytes: Uint8Array;
  font: Font;
}

const loaded = new Map<string, LoadedFont>();

function toBuffer(bytes: Uint8Array): Buffer {
  return Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

/** Register a face's bytes by key (filename for clones, synthetic name for custom
 *  fonts). Built-in registration is first-writer-wins; pass `overwrite` for custom
 *  fonts so an export honors the config it was called with. */
export function registerFont(file: string, bytes: Uint8Array, overwrite = false): void {
  if (loaded.has(file) && !overwrite) return;
  loaded.set(file, { bytes, font: fontkit.create(toBuffer(bytes)) as Font });
}

/** Register the fetched bytes of custom faces (keyed by the synthetic filename so
 *  resolveFont/pdfkit find them). Called from runExport before any layout. */
export function registerCustomFontBytes(faces: CustomFontFaceBytes[]): void {
  for (const f of faces) registerFont(customFontFileName(f.family, f.style), f.bytes, true);
}

export function fontsLoaded(): boolean {
  return loaded.size > 0;
}

/** True once the BUNDLED clones are registered. Distinct from fontsLoaded() so that
 *  registering custom fonts first doesn't trick installMeasureHost into skipping the
 *  bundled-font read (the fallback Arimo face is the sentinel). */
export function builtinsRegistered(): boolean {
  return loaded.has("Arimo-Regular.ttf");
}

export interface ResolvedFont {
  /** Bundled filename — also the pdfkit font registration key. */
  file: string;
  font: Font;
  bytes: Uint8Array;
  /** True when the requested family had no mapping and the fallback was used. */
  substituted: boolean;
}

/** Resolve a requested family+style to a loaded face. A registered custom family
 *  resolves to ITS OWN faces (with a missing style falling back to its Regular),
 *  never substituted; otherwise it maps to the bundled clone (or Arimo). */
export function resolveFont(family: string, bold: boolean, italic: boolean): ResolvedFont {
  const { clone, substituted } = cloneFamilyFor(family);
  const style = bold && italic ? "BoldItalic" : bold ? "Bold" : italic ? "Italic" : "Regular";

  // Custom font: prefer the exact style, fall back to its own Regular. If its bytes
  // somehow aren't registered, fall through to the clone path below.
  if (customFontFor(clone)) {
    const want = customFontFileName(clone, style);
    const reg = customFontFileName(clone, "Regular");
    const hit = loaded.get(want) ?? loaded.get(reg);
    if (hit) {
      const file = loaded.has(want) ? want : reg;
      return { file, font: hit.font, bytes: hit.bytes, substituted: false };
    }
  }

  const file = `${clone}-${style}.ttf`;
  const hit = loaded.get(file) ?? loaded.get("Arimo-Regular.ttf");
  if (!hit) {
    throw new Error("fontRegistry: no fonts loaded — call installMeasureHost() before measuring/exporting");
  }
  return { file, font: hit.font, bytes: hit.bytes, substituted };
}
