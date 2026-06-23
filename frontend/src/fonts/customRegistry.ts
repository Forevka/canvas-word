// Runtime overlay over the built-in clone tables (clones.ts) that lets an embedder
// register their OWN fonts at runtime via the WordCanvas `fonts` option. The
// built-in metric clones stay the source of truth; this module is the additive
// layer the clones.ts resolvers consult so a custom family renders/measures/embeds
// as ITSELF (not substituted to a clone), with caller-supplied vertical metrics so
// the editor and both exporters paginate identically (see FONTS.md).
//
// This module deliberately imports nothing from clones.ts (clones.ts imports IT, so
// keeping this leaf-level avoids an import cycle) and pulls in no heavy deps, so the
// Node export bundle (pipeline.ts → this) stays small.

/** The four faces every font is resolved against. A custom font need only supply
 *  `Regular`; the others fall back to it (same in editor and export). */
export type FontStyleName = "Regular" | "Bold" | "Italic" | "BoldItalic";

export const CUSTOM_FONT_STYLES: FontStyleName[] = ["Regular", "Bold", "Italic", "BoldItalic"];

/** Per-style face URLs for a custom font. `regular` is required; a missing
 *  bold/italic/boldItalic resolves to `regular` in both the editor and exporters
 *  (deterministic parity over faux-synthesis). */
export interface CustomFontFaces {
  regular: string;
  bold?: string;
  italic?: string;
  boldItalic?: string;
}

/** One custom font supplied through configuration. */
export interface CustomFontDef {
  /** Family name — stored in the model, shown in the toolbar, AND the render name
   *  (a custom font is never substituted to a clone). */
  family: string;
  /** Per-style face URLs (TTF/OTF; WOFF2 is not supported). */
  faces: CustomFontFaces;
  /** Vertical metrics as fractions of em — same shape/role as the built-in
   *  CLONE_METRICS. REQUIRED so the editor and exporters compute identical line
   *  heights (and therefore identical pagination). */
  sizing: { ascent: number; descent: number };
  /** Optional toolbar label (defaults to `family`). */
  label?: string;
}

/** Public `fonts` option shape. */
export interface FontsConfig {
  /** Original family names (e.g. "Calibri") to hide from the toolbar. Hidden
   *  built-ins stay LOADED and RESOLVABLE so a loaded .docx that references them
   *  still renders — this only trims the font dropdown. */
  disableBuiltin?: string[];
  /** Custom fonts to register. */
  fonts?: CustomFontDef[];
}

/** Fully-populated `fonts` config (no optional arrays) — what the app threads down. */
export interface ResolvedFontsConfig {
  disableBuiltin: string[];
  fonts: CustomFontDef[];
}

/** A single fetched custom face's bytes, ready to register with fontkit/pdfkit.
 *  Produced on the main thread / backend (where URLs are fetchable) and handed to
 *  the export pipeline. */
export interface CustomFontFaceBytes {
  family: string;
  style: FontStyleName;
  bytes: Uint8Array;
}

/** Transport payload carrying custom fonts into the export pipeline (worker /
 *  backend): the defs (family→sizing, for resolution/metrics) plus the face bytes
 *  (for fontkit measuring + pdfkit embedding). */
export interface CustomFontPayload {
  defs: ResolvedFontsConfig;
  faces: CustomFontFaceBytes[];
}

// ---- global additive registry ----------------------------------------------
// The font subsystem is already a process-global singleton (clones.ts consts, the
// editor FontFace loader, the fontkit `loaded` map). Custom fonts join it as a
// global merge-by-family registry; the toolbar stays per-instance (computed from
// each mount's own config — see toolbarFonts in clones.ts).

const customFonts = new Map<string, CustomFontDef>(); // key = normalized family

/** Normalize a family token the same way clones.ts does (trim, lowercase, strip
 *  surrounding quotes) so lookups agree across the editor and exporters. */
export function normalizeFamily(family: string): string {
  return family.trim().toLowerCase().replace(/^["']|["']$/g, "");
}

/** fontkit-registry + pdfkit registration key for a custom face. Custom fonts have
 *  no on-disk file, so this synthetic name plays the role the bundled filename
 *  plays for clones. */
export function customFontFileName(family: string, style: FontStyleName): string {
  return `__custom__${normalizeFamily(family)}-${style}.ttf`;
}

/** The face URL for a style, falling back to `regular` when absent. */
export function faceUrlForStyle(faces: CustomFontFaces, style: FontStyleName): string {
  switch (style) {
    case "Bold":
      return faces.bold ?? faces.regular;
    case "Italic":
      return faces.italic ?? faces.regular;
    case "BoldItalic":
      return faces.boldItalic ?? faces.regular;
    default:
      return faces.regular;
  }
}

function isWoff2(url: string): boolean {
  return /\.woff2(\?|#|$)/i.test(url);
}

function sameDef(a: CustomFontDef, b: CustomFontDef): boolean {
  return (
    a.faces.regular === b.faces.regular &&
    a.faces.bold === b.faces.bold &&
    a.faces.italic === b.faces.italic &&
    a.faces.boldItalic === b.faces.boldItalic &&
    a.sizing.ascent === b.sizing.ascent &&
    a.sizing.descent === b.sizing.descent
  );
}

/** Register custom fonts into the global overlay (idempotent). Validates `sizing`
 *  and rejects WOFF2 (fontkit can't parse it and the exporter must parse the exact
 *  bytes it embeds). A family redefined with different faces/sizing warns and the
 *  latest wins, so an export always honors the config it was called with. */
export function registerCustomFonts(cfg?: { fonts?: CustomFontDef[] } | undefined): void {
  if (!cfg?.fonts) return;
  for (const def of cfg.fonts) {
    const key = normalizeFamily(def.family);
    if (!key) continue;
    const { ascent, descent } = def.sizing ?? ({} as CustomFontDef["sizing"]);
    if (!(typeof ascent === "number" && ascent > 0 && typeof descent === "number" && descent >= 0)) {
      console.warn(`[wordcanvas] custom font "${def.family}" has invalid sizing (need ascent > 0, descent >= 0); skipped`);
      continue;
    }
    if (!def.faces?.regular) {
      console.warn(`[wordcanvas] custom font "${def.family}" is missing the required regular face; skipped`);
      continue;
    }
    if (
      isWoff2(def.faces.regular) ||
      (def.faces.bold && isWoff2(def.faces.bold)) ||
      (def.faces.italic && isWoff2(def.faces.italic)) ||
      (def.faces.boldItalic && isWoff2(def.faces.boldItalic))
    ) {
      console.warn(`[wordcanvas] custom font "${def.family}": WOFF2 is not supported — use TTF/OTF; skipped`);
      continue;
    }
    const existing = customFonts.get(key);
    if (existing && !sameDef(existing, def)) {
      console.warn(`[wordcanvas] custom font "${def.family}" redefined with different faces/sizing; using the latest`);
    }
    customFonts.set(key, def);
  }
}

/** The custom font registered for a family token, or undefined. */
export function customFontFor(family: string): CustomFontDef | undefined {
  return customFonts.get(normalizeFamily(family));
}

/** Caller-supplied vertical metrics for a custom family, or undefined. */
export function customMetrics(family: string): { ascent: number; descent: number } | undefined {
  return customFonts.get(normalizeFamily(family))?.sizing;
}

/** All registered custom fonts (testing / diagnostics). */
export function allCustomFonts(): CustomFontDef[] {
  return [...customFonts.values()];
}

/** Clear the registry — TEST ONLY (the registry is otherwise process-lifetime). */
export function __resetCustomFonts(): void {
  customFonts.clear();
}
