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

// ---- per-instance registry ---------------------------------------------------
// Custom fonts used to live in ONE process-global map, so two WordCanvas editors
// (or two concurrent export jobs) on the same page/process could clobber each
// other's family→sizing definitions and silently change a live document's
// pagination. They now live in CustomFontRegistry *instances*: each editor mount
// and each export job owns one. The resolvers (cloneFamilyFor / metricsFor /
// resolveFont) read whichever registry is currently ACTIVE; the editor engine and
// paint layer, and the export pipeline, assert their own registry active at the
// top of each SYNCHRONOUS layout/paint span (no `await` between the assertion and
// font resolution), which makes concurrent editors and concurrent export jobs
// safe without a lock. A process-wide DEFAULT registry preserves the old global
// API for tests and any caller that doesn't opt into instances.

/** Normalize a family token the same way clones.ts does (trim, lowercase, strip
 *  surrounding quotes) so lookups agree across the editor and exporters. */
export function normalizeFamily(family: string): string {
  return family.trim().toLowerCase().replace(/^["']|["']$/g, "");
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

/** WOFF and WOFF2 are compressed SFNT wrappers (zlib / Brotli). The export pipeline
 *  feeds the SAME raw bytes to fontkit (to measure glyph widths) and to pdfkit (to
 *  subset-embed), and both need an uncompressed TTF/OTF — so both `.woff` and
 *  `.woff2` are rejected at registration. */
function isWoffUrl(url: string): boolean {
  return /\.woff2?(\?|#|$)/i.test(url);
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

/** True if a def is registrable: valid sizing, a regular face, no WOFF2. Logs the
 *  reason (once per call) when it isn't, so the toolbar and registry agree on which
 *  custom families actually resolve. */
export function isRegistrableFontDef(def: CustomFontDef, warn = false): boolean {
  if (!normalizeFamily(def.family)) return false;
  const { ascent, descent } = def.sizing ?? ({} as CustomFontDef["sizing"]);
  if (!(typeof ascent === "number" && ascent > 0 && typeof descent === "number" && descent >= 0)) {
    if (warn) console.warn(`[wordcanvas] custom font "${def.family}" has invalid sizing (need ascent > 0, descent >= 0); skipped`);
    return false;
  }
  if (!def.faces?.regular) {
    if (warn) console.warn(`[wordcanvas] custom font "${def.family}" is missing the required regular face; skipped`);
    return false;
  }
  if (
    isWoffUrl(def.faces.regular) ||
    (def.faces.bold && isWoffUrl(def.faces.bold)) ||
    (def.faces.italic && isWoffUrl(def.faces.italic)) ||
    (def.faces.boldItalic && isWoffUrl(def.faces.boldItalic))
  ) {
    if (warn) console.warn(`[wordcanvas] custom font "${def.family}": WOFF/WOFF2 is not supported — use TTF/OTF; skipped`);
    return false;
  }
  return true;
}

/** An instance-scoped overlay of custom families. The editor mount and each export
 *  job own one; resolution reads the currently-active instance (see active below). */
export class CustomFontRegistry {
  /** Disambiguates this instance's synthetic face keys in the export-side fontkit
   *  `loaded` map, so two concurrent jobs registering the same family with
   *  different bytes don't collide. Empty for the default registry so its keys
   *  match the legacy `__custom__<family>-<style>.ttf` shape. */
  readonly faceKeyPrefix: string;
  private readonly fonts = new Map<string, CustomFontDef>(); // key = normalized family
  /** layout/metrics.ts memoizes computed FontMetrics here — per registry, so two
   *  instances with the same family string but different sizing never share a cache
   *  entry. */
  readonly metricsCache = new Map<string, { ascent: number; descent: number }>();
  /** layout/metrics.ts memoizes the canvas font string per CharStyle object here —
   *  per registry, because a registered custom font shadows a built-in clone of the
   *  same family name. Keyed by object identity (model styles are immutable; edits
   *  mint new objects). Replaced wholesale whenever the family set changes, since
   *  adding a family can re-route tokens that previously fell back to a clone. */
  charFontCache = new WeakMap<object, string>();

  constructor(faceKeyPrefix = "") {
    this.faceKeyPrefix = faceKeyPrefix;
  }

  /** Merge defs into this registry (idempotent). Invalid defs are skipped (see
   *  isRegistrableFontDef); a family redefined with different faces/sizing warns and
   *  the latest wins. Clears the metrics cache when a def actually changes so stale
   *  line heights can't survive a re-register. */
  register(cfg?: { fonts?: CustomFontDef[] } | undefined): void {
    if (!cfg?.fonts) return;
    let familiesChanged = false;
    for (const def of cfg.fonts) {
      if (!isRegistrableFontDef(def, true)) continue;
      const key = normalizeFamily(def.family);
      const existing = this.fonts.get(key);
      if (existing && !sameDef(existing, def)) {
        console.warn(`[wordcanvas] custom font "${def.family}" redefined with different faces/sizing; using the latest`);
        this.metricsCache.clear();
      }
      if (!existing) familiesChanged = true;
      this.fonts.set(key, def);
    }
    // A new family can shadow a clone that cached font strings already resolved
    // to — drop the memo so affected styles re-resolve. (Redefining an EXISTING
    // family keeps the same render name, so the font strings stay valid.)
    if (familiesChanged) this.charFontCache = new WeakMap();
  }

  fontFor(family: string): CustomFontDef | undefined {
    return this.fonts.get(normalizeFamily(family));
  }

  metricsFor(family: string): { ascent: number; descent: number } | undefined {
    return this.fonts.get(normalizeFamily(family))?.sizing;
  }

  all(): CustomFontDef[] {
    return [...this.fonts.values()];
  }

  /** fontkit-registry + pdfkit registration key for a custom face — the synthetic
   *  name that plays the role the bundled filename plays for clones. */
  faceKey(family: string, style: FontStyleName): string {
    return `__custom__${this.faceKeyPrefix}${normalizeFamily(family)}-${style}.ttf`;
  }

  clear(): void {
    this.fonts.clear();
    this.metricsCache.clear();
    this.charFontCache = new WeakMap();
  }
}

// The default (process-global) registry backs the legacy free-function API; its
// empty prefix keeps face keys byte-identical to the old `__custom__…` names.
const defaultRegistry = new CustomFontRegistry("");
let active: CustomFontRegistry = defaultRegistry;
let registrySeq = 0;

/** A fresh instance-scoped registry with a unique face-key prefix (editor mount /
 *  export job). */
export function createFontRegistry(): CustomFontRegistry {
  return new CustomFontRegistry(`r${++registrySeq}__`);
}

/** The registry resolvers consult right now. Set at the top of each synchronous
 *  layout/paint span by the engine, paint layer, and export pipeline. */
export function activeFontRegistry(): CustomFontRegistry {
  return active;
}

/** Make `reg` the active registry for subsequent resolution. MUST be followed by
 *  the font-resolving work with no intervening `await` (see the module header). */
export function setActiveFontRegistry(reg: CustomFontRegistry): void {
  active = reg;
}

/** The process-wide default registry (export pipeline restores it after a job). */
export function defaultFontRegistry(): CustomFontRegistry {
  return defaultRegistry;
}

/** Legacy synthetic face key on the DEFAULT registry (kept for back-compat). */
export function customFontFileName(family: string, style: FontStyleName): string {
  return defaultRegistry.faceKey(family, style);
}

/** Register custom fonts into the DEFAULT registry (legacy global API). The editor
 *  and export job paths use their own CustomFontRegistry instance instead. */
export function registerCustomFonts(cfg?: { fonts?: CustomFontDef[] } | undefined): void {
  defaultRegistry.register(cfg);
}

/** The custom font registered for a family token in the ACTIVE registry. */
export function customFontFor(family: string): CustomFontDef | undefined {
  return active.fontFor(family);
}

/** Caller-supplied vertical metrics for a custom family in the ACTIVE registry. */
export function customMetrics(family: string): { ascent: number; descent: number } | undefined {
  return active.metricsFor(family);
}

/** All custom fonts in the DEFAULT registry (testing / diagnostics). */
export function allCustomFonts(): CustomFontDef[] {
  return defaultRegistry.all();
}

/** Clear the DEFAULT registry — TEST ONLY. */
export function __resetCustomFonts(): void {
  defaultRegistry.clear();
}
