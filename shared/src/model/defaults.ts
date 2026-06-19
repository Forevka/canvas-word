// The editor's default ("Normal") run + paragraph formatting, as CONCRETE
// values. The model keeps runs/paragraphs concrete (every CharStyle/ParaStyle
// field present), so every "no style yet" path — text insertion fallback,
// paste, the sample document, the builder baselines, comment runs — needs a
// full default to spread from. These are that single source of truth; keep them
// in sync with the "Normal" style declared in defaultStylesheet() (stylesheet.ts),
// which is the same look expressed as a style delta.

import type { CharStyle, ParaStyle } from "./document";

/** Default body run formatting — Georgia 16px on near-black (#202124). */
export const DEFAULT_CHAR_STYLE: CharStyle = {
  fontFamily: "Georgia, serif",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#202124",
};

/** Default paragraph formatting — left-aligned, 1.5 line height, 12px after. */
export const DEFAULT_PARA_STYLE: ParaStyle = {
  align: "left",
  lineHeight: 1.5,
  spaceBeforePx: 0,
  spaceAfterPx: 12,
  indentFirstLinePx: 0,
  indentLeftPx: 0,
};
