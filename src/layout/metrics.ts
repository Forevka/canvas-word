// Vertical font metrics (pretext measures widths; ascent/descent are ours).
// One hidden canvas, results cached per font string.

import type { CharStyle } from "../model/document";

export interface FontMetrics {
  ascent: number;
  descent: number;
}

const measureCtx = document.createElement("canvas").getContext("2d")!;
const metricsCache = new Map<string, FontMetrics>();

export function charStyleToFont(s: CharStyle): string {
  const italic = s.italic ? "italic " : "";
  const weight = s.bold ? "700" : "400";
  // Sub/superscript runs are MEASURED at the scaled size (paint shifts the
  // baseline) — pretext must measure exactly what paint draws.
  const size = s.verticalAlign ? Math.round(s.fontSizePx * 0.65) : s.fontSizePx;
  return `${italic}${weight} ${size}px ${s.fontFamily}`;
}

/** One-off text width at a font — LAYOUT-time only (paint never measures).
 *  Used for paint-only decorations the engine positions (TOC page numbers). */
export function measureTextWidth(text: string, font: string): number {
  measureCtx.font = font;
  return measureCtx.measureText(text).width;
}

export function fontMetrics(font: string): FontMetrics {
  let m = metricsCache.get(font);
  if (!m) {
    measureCtx.font = font;
    const t = measureCtx.measureText("Mg");
    m = {
      ascent: t.fontBoundingBoxAscent,
      descent: t.fontBoundingBoxDescent,
    };
    metricsCache.set(font, m);
  }
  return m;
}
