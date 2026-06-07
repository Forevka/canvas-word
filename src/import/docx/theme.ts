// theme1.xml → font scheme + color scheme. Styles and direct formatting can
// reference these indirectly (w:asciiTheme="minorHAnsi", w:themeColor="accent1");
// the StyleResolver translates the references using this data.

import { attr, children, el, parseXml, rootEl, type XmlNode } from "./xml";

export interface Theme {
  /** a:majorFont latin typeface — headings ("majorHAnsi" etc.). */
  majorLatin?: string;
  /** a:minorFont latin typeface — body text ("minorHAnsi" etc.). */
  minorLatin?: string;
  /** Scheme color slot → hex (no '#'), e.g. "accent1" → "4472C4". */
  colors: Map<string, string>;
}

export const EMPTY_THEME: Theme = { colors: new Map() };

/** w:themeColor names → clrScheme element names. text/background map onto
 *  dk/lt (the standard clrMap; we assume Word's default mapping). */
const THEME_COLOR_ALIASES: Record<string, string> = {
  dark1: "dk1",
  light1: "lt1",
  dark2: "dk2",
  light2: "lt2",
  text1: "dk1",
  background1: "lt1",
  text2: "dk2",
  background2: "lt2",
  hyperlink: "hlink",
  followedHyperlink: "folHlink",
};

export function themeColor(theme: Theme, name: string): string | undefined {
  return theme.colors.get(THEME_COLOR_ALIASES[name] ?? name);
}

export function themeFont(theme: Theme, slot: string): string | undefined {
  if (slot.startsWith("major")) return theme.majorLatin;
  if (slot.startsWith("minor")) return theme.minorLatin;
  return undefined;
}

export function parseThemeXml(xmlText: string): Theme {
  const theme: Theme = { colors: new Map() };
  const root = rootEl(parseXml(xmlText, "theme"), "a:theme");
  const elements = root && el(root, "a:themeElements");
  if (!elements) return theme;

  const fontScheme = el(elements, "a:fontScheme");
  if (fontScheme) {
    const major = latinTypeface(el(fontScheme, "a:majorFont"));
    if (major) theme.majorLatin = major;
    const minor = latinTypeface(el(fontScheme, "a:minorFont"));
    if (minor) theme.minorLatin = minor;
  }

  const clrScheme = el(elements, "a:clrScheme");
  if (clrScheme) {
    for (const slot of children(clrScheme)) {
      const name = slot.tagName.replace(/^a:/, "");
      const srgb = el(slot, "a:srgbClr");
      const sys = el(slot, "a:sysClr");
      const hex = (srgb && attr(srgb, "val")) ?? (sys && attr(sys, "lastClr"));
      if (hex) theme.colors.set(name, hex);
    }
  }
  return theme;
}

function latinTypeface(font: XmlNode | undefined): string | undefined {
  const latin = font && el(font, "a:latin");
  const typeface = latin && attr(latin, "typeface");
  return typeface || undefined; // empty typeface="" → undefined
}
