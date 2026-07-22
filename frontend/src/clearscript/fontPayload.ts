// Assembles the CustomFontPayload the export pipeline consumes from the flat,
// host-marshalled specs the ClearScript C# side ships (family + sizing + per-style
// byte arrays). Kept in its own leaf module — separate from the bundle entry — so it
// is unit-testable without instantiating the whole IIFE, while the payload SHAPE
// (defs + faces) stays owned next to the customRegistry types.

import {
  isRegistrableFontDef,
  type CustomFontDef,
  type CustomFontFaceBytes,
  type CustomFontFaces,
  type CustomFontPayload,
  type FontStyleName,
} from "../fonts/customRegistry";

/** One host-supplied custom font: the family, its REQUIRED vertical metrics
 *  (fractions of em — see CustomFontDef.sizing), and the raw TTF/OTF bytes per style.
 *  Only `regular` is required; a missing bold/italic/boldItalic resolves to regular
 *  (deterministic parity with the editor and the Node/browser paths). This is the
 *  bare-V8 analogue of the browser/Node `fonts` config: there the faces are URLs the
 *  host fetches, here the host hands the bytes straight in. */
export interface CustomFontSpec {
  family: string;
  ascent: number;
  descent: number;
  label?: string;
  regular: Uint8Array;
  bold?: Uint8Array;
  italic?: Uint8Array;
  boldItalic?: Uint8Array;
}

const STYLE_BYTES: [FontStyleName, keyof CustomFontSpec][] = [
  ["Regular", "regular"],
  ["Bold", "bold"],
  ["Italic", "italic"],
  ["BoldItalic", "boldItalic"],
];

/** Build the {@link CustomFontPayload} from host specs. `defs.fonts[].faces` carries
 *  synthetic (never-fetched) markers just so each def validates + resolves as a custom
 *  family; the actual bytes travel in `faces`, keyed by family+style exactly as the
 *  Node backend hands them over. Invalid specs (bad sizing, no regular face) throw —
 *  they'd otherwise silently substitute to a bundled clone. */
export function buildFontPayload(specs: CustomFontSpec[]): CustomFontPayload {
  const fonts: CustomFontDef[] = [];
  const faces: CustomFontFaceBytes[] = [];
  for (const spec of specs) {
    // The regular face's bytes are required. The synthetic markers below are always
    // truthy strings, so isRegistrableFontDef can't see a missing regular face — guard
    // it here (matches the C# host's own validation) or a bytes-less font would slip
    // through and then silently substitute at resolve time.
    if (!(spec.regular instanceof Uint8Array) || spec.regular.length === 0) {
      throw new Error(`custom font "${spec.family}" requires the regular face bytes (TTF/OTF)`);
    }
    // Synthetic marker per present face — non-empty + non-WOFF so isRegistrableFontDef
    // accepts it; never dereferenced (bare V8 doesn't fetch), the bytes are in `faces`.
    const marker = (style: FontStyleName): string => `wc-clearscript://${spec.family}/${style}.ttf`;
    const faceUrls: CustomFontFaces = { regular: marker("Regular") };
    if (spec.bold) faceUrls.bold = marker("Bold");
    if (spec.italic) faceUrls.italic = marker("Italic");
    if (spec.boldItalic) faceUrls.boldItalic = marker("BoldItalic");

    const def: CustomFontDef = {
      family: spec.family,
      faces: faceUrls,
      sizing: { ascent: spec.ascent, descent: spec.descent },
      ...(spec.label ? { label: spec.label } : {}),
    };
    if (!isRegistrableFontDef(def, true)) {
      throw new Error(
        `custom font "${spec.family}" is not registrable (need family, ascent > 0, descent >= 0, a regular TTF/OTF face)`,
      );
    }
    fonts.push(def);
    for (const [style, key] of STYLE_BYTES) {
      const bytes = spec[key] as Uint8Array | undefined;
      if (bytes) faces.push({ family: spec.family, style, bytes });
    }
  }
  return { defs: { disableBuiltin: [], fonts }, faces };
}
