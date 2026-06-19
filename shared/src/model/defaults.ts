// The editor's default ("Normal") run + paragraph formatting, as CONCRETE
// values. The model keeps runs/paragraphs concrete (every CharStyle/ParaStyle
// field present), so every "no style yet" path — text insertion fallback,
// paste, the sample document, the builder baselines, comment runs — needs a
// full default to spread from. These are that single source of truth; keep them
// in sync with the "Normal" style declared in makeDefaultStylesheet()
// (stylesheet.ts), which is the same look expressed as a style delta.
//
// Embedders can OVERRIDE the library's built-in defaults (see EditorTypography
// + the WordCanvas `overrideDefaultStyles` option). This only affects NEW/blank
// documents and the fallback stylesheet — a loaded .docx keeps its own
// w:docDefaults / Normal style (the importer bakes those in at import time).

import type { CharStyle, ParaStyle } from "./document";

/** Per-instance typography override for the library's built-in defaults. Every
 *  field is optional; omit to keep the built-in value. Drives both the concrete
 *  default run/paragraph styles and the generated default stylesheet. */
export interface EditorTypography {
  /** Body (Normal) font family, e.g. "Times New Roman, serif". */
  fontFamily?: string;
  /** Body font size in px (96 dpi). */
  fontSizePx?: number;
  /** Body text color (any CSS color string). */
  color?: string;
  /** Default paragraph line height (multiplier). */
  lineHeight?: number;
  /** Heading font family (Title/Heading1/Heading2). Omit to keep the built-in. */
  headingFontFamily?: string;
}

/** Built-in body run formatting — Georgia 16px on near-black (#202124). */
export function makeDefaultCharStyle(t: EditorTypography = {}): CharStyle {
  return {
    fontFamily: t.fontFamily ?? "Georgia, serif",
    fontSizePx: t.fontSizePx ?? 16,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    color: t.color ?? "#202124",
  };
}

/** Built-in paragraph formatting — left, 1.5 line height, 12px after. */
export function makeDefaultParaStyle(t: EditorTypography = {}): ParaStyle {
  return {
    align: "left",
    lineHeight: t.lineHeight ?? 1.5,
    spaceBeforePx: 0,
    spaceAfterPx: 12,
    indentFirstLinePx: 0,
    indentLeftPx: 0,
  };
}

/** Default body run formatting (built-in, no overrides). */
export const DEFAULT_CHAR_STYLE: CharStyle = makeDefaultCharStyle();

/** Default paragraph formatting (built-in, no overrides). */
export const DEFAULT_PARA_STYLE: ParaStyle = makeDefaultParaStyle();
