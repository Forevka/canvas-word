// Font registry for export — loads the bundled metric clones and resolves a
// requested family to a parsed fontkit face. The SAME face both measures (layout
// widths) and is embedded into the PDF, so painted glyphs land where measured.
// Family→clone mapping is shared with the editor via src/fonts/clones.ts.

import * as fontkit from "fontkit";
import type { Font } from "fontkit";
import { cloneFamilyFor, FONT_FILES, SINGLE_FACE_FILES } from "../../fonts/clones";
import {
  activeFontRegistry,
  customFontFor,
  type CustomFontFaceBytes,
  type CustomFontRegistry,
} from "../../fonts/customRegistry";

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

/** Register the fetched bytes of custom faces (keyed by the ACTIVE registry's
 *  synthetic filename so resolveFont/pdfkit find them). Called from runExport
 *  inside the synchronous span after the job's registry is made active. */
export function registerCustomFontBytes(faces: CustomFontFaceBytes[]): void {
  const reg = activeFontRegistry();
  for (const f of faces) registerFont(reg.faceKey(f.family, f.style), f.bytes, true);
}

/** Drop a job registry's custom faces from the loaded map after its export, so a
 *  long-running process (Node backend / shared worker) doesn't accumulate every
 *  job's font bytes forever. No-op for the default registry (empty prefix). */
export function clearCustomFontBytes(reg: CustomFontRegistry): void {
  if (!reg.faceKeyPrefix) return;
  const tag = `__custom__${reg.faceKeyPrefix}`;
  for (const key of [...loaded.keys()]) if (key.startsWith(tag)) loaded.delete(key);
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

  // Custom font: prefer the exact style, fall back to its own Regular.
  if (customFontFor(clone)) {
    const reg = activeFontRegistry();
    const wantKey = reg.faceKey(clone, style);
    const regularKey = reg.faceKey(clone, "Regular");
    const hit = loaded.get(wantKey) ?? loaded.get(regularKey);
    if (hit) {
      const file = loaded.has(wantKey) ? wantKey : regularKey;
      return { file, font: hit.font, bytes: hit.bytes, substituted: false };
    }
    // Registered as a custom family but its bytes never loaded (e.g. the required
    // Regular face failed to fetch). Don't emit bundled bytes under a custom file
    // name with substituted:false — that would make the exported face metadata lie.
    // Fall back to the visible substitution path so a warning fires and the
    // reported file/flag are honest.
    const fb = loaded.get("Arimo-Regular.ttf");
    if (!fb) throw new Error("fontRegistry: no fonts loaded — call installMeasureHost() before measuring/exporting");
    return { file: "Arimo-Regular.ttf", font: fb.font, bytes: fb.bytes, substituted: true };
  }

  // Single-face built-ins (math, CJK fallback) ship only a Regular face: every
  // style resolves to it, so bold/italic CJK don't fall through to Arimo (tofu).
  const file = SINGLE_FACE_FILES[clone] ?? `${clone}-${style}.ttf`;
  const hit = loaded.get(file);
  if (hit) return { file, font: hit.font, bytes: hit.bytes, substituted };
  // The requested face isn't loaded (e.g. a broken bundle, or the CJK font failed
  // to register — installMeasureHost warns-and-continues). Fall back to Arimo, but
  // report ITS file + substituted:true so the returned metadata describes the face
  // actually used (never the requested filename with someone else's bytes).
  const fb = loaded.get("Arimo-Regular.ttf");
  if (!fb) {
    throw new Error("fontRegistry: no fonts loaded — call installMeasureHost() before measuring/exporting");
  }
  return { file: "Arimo-Regular.ttf", font: fb.font, bytes: fb.bytes, substituted: true };
}
