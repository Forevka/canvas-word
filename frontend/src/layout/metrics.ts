// Vertical font metrics (pretext measures widths; ascent/descent are ours).
// One hidden canvas, results cached per font string.

import type { CharStyle } from "@cw/shared";
import { cloneFamilyFor, firstFamilyToken, metricsFor } from "../fonts/clones";
import { activeFontRegistry } from "../fonts/customRegistry";
import { SUB_SUPER_SCALE } from "../paint/paintStyle";

export interface FontMetrics {
  ascent: number;
  descent: number;
}

/** The slice of CanvasRenderingContext2D the layout actually measures through.
 *  DOM-free hosts (the export worker, Node backends) inject their own via
 *  {@link setMeasureContext}; the browser falls back to a real canvas. */
export interface MeasureContext {
  font: string;
  measureText(text: string): {
    width: number;
    fontBoundingBoxAscent: number;
    fontBoundingBoxDescent: number;
  };
}

let injectedCtx: MeasureContext | null = null;
let browserCtx: MeasureContext | null = null;

/** Install a measurement context (export/Node hosts). Must be called before the
 *  engine measures any text. No-op for the browser, which uses a real canvas. */
export function setMeasureContext(ctx: MeasureContext): void {
  injectedCtx = ctx;
}

// Lazily created so importing this module never touches the DOM — the engine can
// run under Node once a context is injected (browser path stays byte-identical).
function measureContext(): MeasureContext {
  if (injectedCtx) return injectedCtx;
  if (!browserCtx) {
    browserCtx = document.createElement("canvas").getContext("2d")! as unknown as MeasureContext;
  }
  return browserCtx;
}

export function charStyleToFont(s: CharStyle): string {
  // Memoized per (active registry, style object) — this runs per fragment on
  // every layout, paint, and hit-test pass, and the family resolution (registry
  // lookup + normalize + clone map) dominates it. Object identity is a sound key:
  // model CharStyles are immutable (edits mint new objects — the same contract
  // the engine's revision-keyed line caches rely on). The registry keys the cache
  // because a registered custom font shadows a built-in clone of the same name.
  const cache = activeFontRegistry().charFontCache;
  const hit = cache.get(s);
  if (hit !== undefined) return hit;
  const italic = s.italic ? "italic " : "";
  const weight = s.bold ? "700" : "400";
  // Sub/superscript runs are MEASURED at the scaled size (paint shifts the
  // baseline) — pretext must measure exactly what paint draws.
  const size = s.verticalAlign ? Math.round(s.fontSizePx * SUB_SUPER_SCALE) : s.fontSizePx;
  // Render & measure with the bundled metric clone (e.g. Calibri -> Carlito), so
  // the editor, the browser export, and the Node backend all lay out identically
  // — no dependency on system fonts. The model keeps the original family.
  const family = cloneFamilyFor(firstFamilyToken(s.fontFamily)).clone;
  const font = `${italic}${weight} ${size}px ${family}`;
  cache.set(s, font);
  return font;
}

/** One-off text width at a font — LAYOUT-time only (paint never measures).
 *  Used for paint-only decorations the engine positions (TOC page numbers). */
export function measureTextWidth(text: string, font: string): number {
  const ctx = measureContext();
  ctx.font = font;
  return ctx.measureText(text).width;
}

// Vertical metrics come from baked per-clone ratios (NOT the canvas/fontkit
// context), so the editor and the exporters compute identical line heights and
// therefore identical pagination on every platform. charStyleToFont already
// emits the clone family, so the last token of the font string is it.
export function fontMetrics(font: string): FontMetrics {
  // Per-registry cache: a custom family's sizing is instance-scoped, so keying on
  // the active registry's cache keeps two editors with the same family name but
  // different metrics from sharing a (wrong) line height.
  const metricsCache = activeFontRegistry().metricsCache;
  let m = metricsCache.get(font);
  if (!m) {
    const sizeMatch = /(\d+(?:\.\d+)?)px/.exec(font);
    const size = sizeMatch ? parseFloat(sizeMatch[1]!) : 16;
    const family = font.slice(font.indexOf("px ") + 3).trim();
    const r = metricsFor(cloneFamilyFor(family).clone);
    m = { ascent: Math.round(r.ascent * size), descent: Math.round(r.descent * size) };
    metricsCache.set(font, m);
  }
  return m;
}
