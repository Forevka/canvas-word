// Loads fonts as FontFaces on the editor's main thread so the canvas measures &
// paints with the exact faces the exporters embed — identical layout in the editor,
// the browser export, and the Node backend. Two sources:
//   - the BUNDLED metric clones (Carlito, Arimo, …), registered under their clone
//     names; charStyleToFont maps every document family to a clone name.
//   - the embedder's CUSTOM fonts, registered under their own family name (which
//     charStyleToFont also emits verbatim for a registered custom family).
// Call (and await) before the first layout.

import { ARABIC_FONT_FAMILY, ARABIC_FONT_FILE, CJK_FONT_FAMILY, CJK_FONT_FILE, CLONE_FAMILIES, FONT_STYLES, HEBREW_FONT_FAMILY, HEBREW_FONT_FILE } from "../../fonts/clones";
import {
  CUSTOM_FONT_STYLES,
  faceUrlForStyle,
  normalizeFamily,
  type FontStyleName,
  type ResolvedFontsConfig,
} from "../../fonts/customRegistry";

const STYLE_ATTRS: Record<string, { weight: string; style: string }> = {
  Regular: { weight: "400", style: "normal" },
  Bold: { weight: "700", style: "normal" },
  Italic: { weight: "400", style: "italic" },
  BoldItalic: { weight: "700", style: "italic" },
};

let builtinStarted: Promise<void> | null = null;
// Per-face memoization (keyed by family∥style∥url) so a SECOND editor instance with
// DIFFERENT custom fonts still loads its faces — a single shared promise would skip
// them. FontFace.add is effectively idempotent at the document level. Resolves to
// whether the face actually loaded.
const customLoaded = new Map<string, Promise<boolean>>();

/** Fetch + register the bundled clones and any custom fonts (idempotent).
 *  `onProgress(loaded, total)` reports the bundled fonts' streamed bytes (the
 *  dominant ~9 MB) for a smooth bar; a memoized call fires once at 100%. Resolves to
 *  the set of normalized custom family names whose REQUIRED Regular face loaded — the
 *  caller registers only those, so a family that failed to load (CORS/404) falls back
 *  to a clone consistently in the editor (matching what the exporter does). */
export function loadEditorFonts(
  onProgress?: (loaded: number, total: number) => void,
  customFonts?: ResolvedFontsConfig,
): Promise<Set<string>> {
  if (typeof document === "undefined" || !("fonts" in document)) return Promise.resolve(new Set());
  // The math font is awaited alongside the clones so it's in document.fonts before
  // the first paint (no math tofu flash). loadMathFont never rejects (it catches its
  // own errors), so a CDN/asset failure can't break the clone loading.
  return Promise.all([loadBuiltins(onProgress), loadCustom(customFonts), loadMathFont()]).then(([, loaded]) => loaded);
}

let mathStarted: Promise<void> | null = null;

/** Load the bundled STIX Two Math face (single Regular) under family "StixTwoMath".
 *  Defensive: catches its own errors so it can't reject the editor's font loading. */
function loadMathFont(): Promise<void> {
  if (mathStarted) return mathStarted;
  mathStarted = (async () => {
    try {
      const url = new URL("./fonts/StixTwoMath-Regular.ttf", import.meta.url);
      const buf = await (await fetch(url)).arrayBuffer();
      const face = new FontFace("StixTwoMath", buf, { weight: "400", style: "normal" });
      await face.load();
      (document as Document & { fonts: FontFaceSet }).fonts.add(face);
    } catch (e) {
      console.warn("[wordcanvas] failed to load the math font (StixTwoMath); equations fall back to a clone", e);
    }
  })();
  return mathStarted;
}

let cjkStarted: Promise<boolean> | null = null;

/** Lazily load the bundled CJK fallback face (subset of Noto Sans SC) under family
 *  "NotoSansSC". Deliberately NOT awaited by loadEditorFonts: the face is multi-MB,
 *  so blocking first paint on it would penalize the common Latin-only document. The
 *  caller starts it AFTER mount and calls editor.refreshFonts() when it resolves so
 *  CJK text re-measures against this face (matching the exporters). Registered for
 *  every weight/style from the one buffer (no faux bold/italic) so the canvas paints
 *  the same Regular outlines the single-face exporter embeds. Resolves to whether the
 *  face actually loaded. Defensive: never rejects. */
export function loadCjkFallbackFont(): Promise<boolean> {
  if (cjkStarted) return cjkStarted;
  if (typeof document === "undefined" || !("fonts" in document)) return Promise.resolve(false);
  cjkStarted = (async () => {
    try {
      const url = new URL(`./fonts/${CJK_FONT_FILE}`, import.meta.url);
      const buf = await (await fetch(url)).arrayBuffer();
      const faces = Object.values(STYLE_ATTRS).map(
        (a) => new FontFace(CJK_FONT_FAMILY, buf, { weight: a.weight, style: a.style }),
      );
      await Promise.all(faces.map((f) => f.load()));
      for (const f of faces) (document as Document & { fonts: FontFaceSet }).fonts.add(f);
      return true;
    } catch (e) {
      // The browser still falls back to a system CJK face on screen; only export
      // parity (and re-measurement) is lost. Surfaced, not fatal.
      console.warn(`[wordcanvas] failed to load the bundled CJK fallback font (${CJK_FONT_FAMILY})`, e);
      return false;
    }
  })();
  return cjkStarted;
}

let arabicStarted: Promise<boolean> | null = null;

/** Lazily load the bundled Arabic fallback face (Noto Sans Arabic) under family
 *  "NotoSansArabic". Mirrors {@link loadCjkFallbackFont}: deliberately NOT awaited
 *  at mount so it doesn't block the first paint of Latin-only documents. The caller
 *  starts it after mount and calls editor.refreshFonts() when it resolves so Arabic
 *  text re-measures against the bundled face (contextual shaping via GSUB, matching
 *  the exporters). Registered for every weight/style from the one buffer (single
 *  Regular face; no faux bold/italic). Resolves to whether the face loaded.
 *  Defensive: never rejects. */
export function loadArabicFallbackFont(): Promise<boolean> {
  if (arabicStarted) return arabicStarted;
  if (typeof document === "undefined" || !("fonts" in document)) return Promise.resolve(false);
  arabicStarted = (async () => {
    try {
      const url = new URL(`./fonts/${ARABIC_FONT_FILE}`, import.meta.url);
      const buf = await (await fetch(url)).arrayBuffer();
      const faces = Object.values(STYLE_ATTRS).map(
        (a) => new FontFace(ARABIC_FONT_FAMILY, buf, { weight: a.weight, style: a.style }),
      );
      await Promise.all(faces.map((f) => f.load()));
      for (const f of faces) (document as Document & { fonts: FontFaceSet }).fonts.add(f);
      return true;
    } catch (e) {
      console.warn(`[wordcanvas] failed to load the bundled Arabic fallback font (${ARABIC_FONT_FAMILY})`, e);
      return false;
    }
  })();
  return arabicStarted;
}

let hebrewStarted: Promise<boolean> | null = null;

/** Lazily load the bundled Hebrew fallback face (Noto Sans Hebrew) under family
 *  "NotoSansHebrew". Mirrors {@link loadArabicFallbackFont}: deliberately NOT awaited
 *  at mount so it doesn't block the first paint of Latin-only documents. The caller
 *  starts it after mount and calls editor.refreshFonts() when it resolves so Hebrew
 *  text re-measures against the bundled face (matching the exporters). Registered for
 *  every weight/style from the one buffer (single Regular face; no faux bold/italic).
 *  Resolves to whether the face loaded. Defensive: never rejects. */
export function loadHebrewFallbackFont(): Promise<boolean> {
  if (hebrewStarted) return hebrewStarted;
  if (typeof document === "undefined" || !("fonts" in document)) return Promise.resolve(false);
  hebrewStarted = (async () => {
    try {
      const url = new URL(`./fonts/${HEBREW_FONT_FILE}`, import.meta.url);
      const buf = await (await fetch(url)).arrayBuffer();
      const faces = Object.values(STYLE_ATTRS).map(
        (a) => new FontFace(HEBREW_FONT_FAMILY, buf, { weight: a.weight, style: a.style }),
      );
      await Promise.all(faces.map((f) => f.load()));
      for (const f of faces) (document as Document & { fonts: FontFaceSet }).fonts.add(f);
      return true;
    } catch (e) {
      console.warn(`[wordcanvas] failed to load the bundled Hebrew fallback font (${HEBREW_FONT_FAMILY})`, e);
      return false;
    }
  })();
  return hebrewStarted;
}

function loadBuiltins(onProgress?: (loaded: number, total: number) => void): Promise<void> {
  if (builtinStarted) {
    if (onProgress) void builtinStarted.then(() => onProgress(1, 1));
    return builtinStarted;
  }
  builtinStarted = (async () => {
    const specs = CLONE_FAMILIES.flatMap((fam) =>
      FONT_STYLES.map((style) => ({ fam, style, url: new URL(`./fonts/${fam}-${style}.ttf`, import.meta.url) })),
    );
    // Kick off every fetch, then read content-length up front so `total` (and thus
    // the percentage) is known before any body finishes streaming.
    const responses = await Promise.all(specs.map((s) => fetch(s.url)));
    const total = responses.reduce((sum, r) => sum + (Number(r.headers.get("content-length")) || 0), 0);
    let loaded = 0;
    const buffers = await Promise.all(
      responses.map(async (r) => {
        // No streamable body or unknown size: fall back to a single chunk so the
        // bar still advances (just less smoothly) by the buffer's byte length.
        if (!r.body || total === 0) {
          const buf = await r.arrayBuffer();
          loaded += buf.byteLength;
          onProgress?.(loaded, total || loaded);
          return buf;
        }
        const reader = r.body.getReader();
        const chunks: Uint8Array[] = [];
        let size = 0;
        for (;;) {
          const { done, value } = await reader.read();
          if (done) break;
          chunks.push(value);
          size += value.byteLength;
          loaded += value.byteLength;
          onProgress?.(loaded, total);
        }
        const buf = new Uint8Array(size);
        let off = 0;
        for (const c of chunks) {
          buf.set(c, off);
          off += c.byteLength;
        }
        return buf.buffer;
      }),
    );
    const faces = specs.map((s, i) => {
      const a = STYLE_ATTRS[s.style]!;
      return new FontFace(s.fam, buffers[i]!, { weight: a.weight, style: a.style });
    });
    await Promise.all(faces.map((f) => f.load()));
    for (const f of faces) (document as Document & { fonts: FontFaceSet }).fonts.add(f);
  })();
  return builtinStarted;
}

function loadCustom(cfg?: ResolvedFontsConfig): Promise<Set<string>> {
  if (!cfg || cfg.fonts.length === 0) return Promise.resolve(new Set());
  const loaded = new Set<string>();
  return Promise.all(
    cfg.fonts.flatMap((f) =>
      CUSTOM_FONT_STYLES.map(async (style) => {
        const ok = await loadCustomFace(f.family, style, faceUrlForStyle(f.faces, style));
        // A family is "loaded" once its required Regular face is in document.fonts.
        if (ok && style === "Regular") loaded.add(normalizeFamily(f.family));
      }),
    ),
  ).then(() => loaded);
}

function loadCustomFace(family: string, style: FontStyleName, url: string): Promise<boolean> {
  const key = `${family} ${style} ${url}`;
  let p = customLoaded.get(key);
  if (!p) {
    p = (async () => {
      try {
        const buf = await (await fetch(url)).arrayBuffer();
        const a = STYLE_ATTRS[style]!;
        // Register under the custom FAMILY name (what charStyleToFont emits) at the
        // style's weight/style. A missing optional face reuses the regular URL, so
        // the browser paints regular glyphs for it — matching the exporter, which
        // also resolves a missing style to the custom Regular (deterministic parity).
        const face = new FontFace(family, buf, { weight: a.weight, style: a.style });
        await face.load();
        (document as Document & { fonts: FontFaceSet }).fonts.add(face);
        return true;
      } catch (e) {
        // Don't hard-fail the editor for one bad URL — the canvas falls back to the
        // browser's default and resolveFont falls back to a clone in export.
        console.warn(`[wordcanvas] failed to load custom font "${family}" ${style} from ${url}`, e);
        return false;
      }
    })();
    customLoaded.set(key, p);
  }
  return p;
}
