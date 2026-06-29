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

/** Apply OOXML w:themeTint / w:themeShade to a base sRGB hex (no '#'). themeTint
 *  lightens the color toward white, themeShade darkens it toward black; each is a
 *  hex byte ("00".."FF") read as a fraction = value/255 and applied per RGB
 *  channel — the linear interpretation Word uses for w:color tints/shades. At most
 *  one applies (Word never emits both). Returns hex without '#'; unparseable input
 *  is returned unchanged. */
export function applyTintShade(hex: string, tint?: string, shade?: string): string {
  const rgb = parseHex(hex);
  if (!rgb) return hex;
  let [r, g, b] = rgb;
  const tintF = parseByteFraction(tint);
  const shadeF = parseByteFraction(shade);
  if (tintF !== undefined) {
    const lift = (c: number): number => Math.round(c * tintF + 255 * (1 - tintF));
    [r, g, b] = [lift(r), lift(g), lift(b)];
  } else if (shadeF !== undefined) {
    [r, g, b] = [Math.round(r * shadeF), Math.round(g * shadeF), Math.round(b * shadeF)];
  }
  return toHexByte(r) + toHexByte(g) + toHexByte(b);
}

function parseHex(hex: string): [number, number, number] | undefined {
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return undefined;
  return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
}

/** A hex byte ("00".."FF") as a 0..1 fraction, or undefined if absent/invalid. */
function parseByteFraction(v: string | undefined): number | undefined {
  if (!v || !/^[0-9a-fA-F]{1,2}$/.test(v)) return undefined;
  return parseInt(v, 16) / 255;
}

const toHexByte = (c: number): string => Math.max(0, Math.min(255, c)).toString(16).padStart(2, "0");

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
