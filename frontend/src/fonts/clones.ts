// Single source of truth for font substitution. The editor and both exporters
// render & measure with these freely-redistributable metric clones instead of
// the (Windows-only) originals, so layout is identical across the editor, the
// browser export, and the Node backend — no system fonts anywhere. The model
// keeps the original family names; only rendering/measuring/embedding map to a
// clone. See src/export/shared/fonts/LICENSES.md.

import { customFontFor, customMetrics, isRegistrableFontDef, type CustomFontDef } from "./customRegistry";

// "TimesNewRoman" is the genuine Microsoft Times New Roman (bundled at the user's
// request), not a clone — see FONTS.md / fonts/LICENSES.md for the licensing note.
export const CLONE_FAMILIES = ["Carlito", "Caladea", "Gelasio", "Arimo", "TimesNewRoman", "Cousine"] as const;
export const FALLBACK_CLONE = "Arimo";

// Requested family (lowercased, first CSS-stack token) -> bundled clone family.
// Includes identity entries for the clone names themselves (so already-mapped
// strings resolve to themselves) and the generic CSS families.
export const CLONE_OF: Record<string, string> = {
  // metric clones of common Office/Windows fonts
  calibri: "Carlito",
  cambria: "Caladea",
  georgia: "Gelasio",
  arial: "Arimo",
  helvetica: "Arimo",
  "helvetica neue": "Arimo",
  "times new roman": "TimesNewRoman",
  times: "TimesNewRoman",
  "courier new": "Cousine",
  courier: "Cousine",
  // reasonable substitutes for fonts we don't ship
  verdana: "Arimo",
  tahoma: "Arimo",
  "segoe ui": "Arimo",
  "trebuchet ms": "Arimo",
  consolas: "Cousine",
  monaco: "Cousine",
  "lucida console": "Cousine",
  garamond: "TimesNewRoman",
  "book antiqua": "Caladea",
  palatino: "Caladea",
  // generic CSS families
  serif: "TimesNewRoman",
  "sans-serif": "Arimo",
  sans: "Arimo",
  monospace: "Cousine",
  mono: "Cousine",
  // identity (already a clone)
  carlito: "Carlito",
  caladea: "Caladea",
  gelasio: "Gelasio",
  arimo: "Arimo",
  timesnewroman: "TimesNewRoman",
  cousine: "Cousine",
};

// Per-clone vertical metrics as a fraction of font size (ascent, descent above/
// below the baseline). Used by fontMetrics so line heights are computed the SAME
// way in the editor and both exporters — otherwise the browser's canvas
// fontBoundingBox (editor) and fontkit's tables (export) disagree by a hair and
// pagination drifts by a page. Values measured from the browser's canvas
// fontBoundingBox at 16px (an exact multiple here), so the editor's appearance is
// unchanged while the export now matches it on every platform.
export const CLONE_METRICS: Record<string, { ascent: number; descent: number }> = {
  Carlito: { ascent: 0.9375, descent: 0.25 },
  Caladea: { ascent: 0.9375, descent: 0.25 },
  Gelasio: { ascent: 0.9375, descent: 0.3125 },
  Arimo: { ascent: 0.875, descent: 0.1875 },
  TimesNewRoman: { ascent: 0.875, descent: 0.1875 },
  Cousine: { ascent: 0.625, descent: 0.1875 },
};

/** Resolve one CSS-stack family token to the face family the editor/exporters
 *  render with. A registered custom font resolves to ITSELF (no clone, no
 *  substitution) and shadows a built-in of the same name; otherwise it maps to the
 *  bundled clone (or the Arimo fallback). */
export function cloneFamilyFor(family: string): { clone: string; substituted: boolean } {
  const custom = customFontFor(family);
  if (custom) return { clone: custom.family, substituted: false };
  const key = family.trim().toLowerCase().replace(/^["']|["']$/g, "");
  const clone = CLONE_OF[key];
  return clone ? { clone, substituted: false } : { clone: FALLBACK_CLONE, substituted: true };
}

/** Vertical metrics (fractions of em) for a resolved face name — caller-supplied
 *  for custom fonts, baked for clones, with a safe default otherwise. Consumed by
 *  layout/metrics.ts so editor and exporters agree on line height. */
export function metricsFor(name: string): { ascent: number; descent: number } {
  return customMetrics(name) ?? CLONE_METRICS[name] ?? { ascent: 0.9, descent: 0.25 };
}

/** First token of a CSS font-family stack ("Calibri, serif" -> "Calibri"). */
export const firstFamilyToken = (stack: string): string => (stack.split(",")[0] ?? "").trim();

export const FONT_STYLES = ["Regular", "Bold", "Italic", "BoldItalic"] as const;

/** Every bundled face filename. */
export const FONT_FILES: string[] = CLONE_FAMILIES.flatMap((fam) =>
  FONT_STYLES.map((style) => `${fam}-${style}.ttf`),
);

/** Toolbar font list: value is the original family (stored in the model), label
 *  shows the clone we actually render — e.g. "Calibri (Carlito)". */
export const TOOLBAR_FONTS: { value: string; label: string }[] = [
  { value: "Calibri, sans-serif", label: "Calibri (Carlito)" },
  { value: "Cambria, serif", label: "Cambria (Caladea)" },
  { value: "Georgia, serif", label: "Georgia (Gelasio)" },
  { value: "Arial, sans-serif", label: "Arial (Arimo)" },
  { value: "Times New Roman, serif", label: "Times New Roman" },
  { value: "Verdana, sans-serif", label: "Verdana (Arimo)" },
  { value: "Consolas, monospace", label: "Consolas (Cousine)" },
  { value: "Courier New, monospace", label: "Courier New (Cousine)" },
];

/** The font-dropdown list for one editor instance: the built-ins minus any hidden
 *  by `disableBuiltin` (matched on the original family name), plus the instance's
 *  custom fonts. Computed PER MOUNT from that instance's config, so the toolbar is
 *  per-instance even though the underlying custom-font registry is global. */
export function toolbarFonts(cfg?: { disableBuiltin?: string[]; fonts?: CustomFontDef[] }): { value: string; label: string }[] {
  const disabled = new Set((cfg?.disableBuiltin ?? []).map((s) => firstFamilyToken(s).trim().toLowerCase()));
  const builtins = TOOLBAR_FONTS.filter((f) => !disabled.has(firstFamilyToken(f.value).trim().toLowerCase()));
  // Only advertise custom families that actually register (valid sizing, a regular
  // face, no WOFF2) — otherwise the dropdown lists a font cloneFamilyFor can never
  // resolve, and picking it silently falls back to the default clone.
  const custom = (cfg?.fonts ?? [])
    .filter((f) => isRegistrableFontDef(f))
    .map((f) => ({ value: f.family, label: f.label ?? f.family }));
  return [...builtins, ...custom];
}
