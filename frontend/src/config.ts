// Per-instance editor configuration — themes, default typography, and behavior
// tuning that an embedder passes to the WordCanvas constructor. Everything here
// is RESOLVED once (public partial → fully-populated internal object) and then
// threaded down through createEditor → the paint layer, rulers, zoom/indent, and
// the initial stylesheet. Nothing is a mutable module singleton, so several
// editors with different configs coexist on one page.
//
// Defaults below ARE the library's built-in look; omit any field to keep it.

import type { EditorTypography, Stylesheet } from "@cw/shared";
import { makeDefaultStylesheet } from "@cw/shared";
import type { FontsConfig, ResolvedFontsConfig } from "./fonts/customRegistry";
import { ARABIC_FONT_FAMILY, CJK_FONT_FAMILY, HEBREW_FONT_FAMILY } from "./fonts/clones";
import {
  ACCENT_BLUE,
  COLUMN_SEPARATOR_COLOR,
  DEFAULT_GRID_COLOR,
  EXTERNAL_LINK_COLOR,
  FOOTNOTE_RULE_COLOR,
  FORMATTING_MARK_COLOR,
  IMAGE_PLACEHOLDER_COLOR,
  TOC_LEADER_COLOR,
} from "./paint/paintStyle";
import { INDENT_STEP_PX, ZOOM_MAX, ZOOM_MIN, ZOOM_STEP } from "./uiConstants";

/** The type behind the WordCanvas `overrideDefaultStyles` option. Overrides the
 *  LIBRARY's built-in default run/paragraph styles for NEW/blank documents and
 *  the fallback stylesheet — NOT a loaded .docx's own w:docDefaults / Normal. */
export type DefaultStyleOverrides = EditorTypography;

export type { CustomFontDef, CustomFontFaces, FontsConfig, ResolvedFontsConfig } from "./fonts/customRegistry";

/** Ruler band styling (the strip the horizontal + vertical rulers paint). */
export interface RulerTheme {
  /** Band fill behind the ticks. */
  bg?: string;
  /** The lighter "content area" portion of the band (within the margins). */
  content?: string;
  /** Tick + boundary line color. */
  line?: string;
  /** Number label color. */
  label?: string;
  /** CSS font for the number labels. */
  font?: string;
}

/** Editor color theme — every field optional; omit to keep the built-in value.
 *  Pass a (partial) object to `WordCanvas({ theme })`. Affects the on-screen
 *  editor only (exported PDFs keep the built-in look). */
export interface EditorTheme {
  /** Gray gutter behind the pages (the scroll area + ruler troughs). */
  canvasBackground?: string;
  /** Native table gridlines when a cell carries no explicit border. */
  grid?: string;
  /** Drawing-grid overlay mesh (the dotted snap grid). */
  gridMesh?: string;
  /** External hyperlink color. */
  externalLink?: string;
  /** UI accent (band-edit boundary and similar chrome affordances). */
  accent?: string;
  /** Footnote separator rule. */
  footnoteRule?: string;
  /** Non-printing formatting marks (space dots, tab arrows, pilcrows). */
  formattingMark?: string;
  /** TOC dot-leader color. */
  tocLeader?: string;
  /** Placeholder fill for an image whose bitmap hasn't loaded. */
  imagePlaceholder?: string;
  /** Rule between newspaper columns. */
  columnSeparator?: string;
  /** Blinking text caret color. */
  caret?: string;
  /** Find-highlight color. */
  searchHighlight?: string;
  /** Comment pin fill once its thread is resolved. */
  reviewPinResolved?: string;
  /** Comment pin stroke. */
  reviewPinStroke?: string;
  /** Gap (px) between stacked pages. */
  pageGapPx?: number;
  /** Ruler band styling. */
  ruler?: RulerTheme;
}

/** East-Asian / CJK + Arabic tuning. Line-breaking between CJK characters and the
 *  kinsoku rules work out of the box (pretext); these only tune the edge cases and
 *  the fonts used for non-Latin scripts. */
export interface CjkConfig {
  /** BCP-47 locale passed to pretext's analyzer (e.g. "ja", "ko", "zh"). Tunes
   *  locale-specific breaking such as Korean word-keep. Absent = language-neutral.
   *  NOTE: pretext's analyzer locale is process-global, so the LAST editor to
   *  mount wins when several use different locales on one page. */
  locale?: string;
  /** Family name of the font to use for CJK runs. Defaults to the bundled CJK
   *  fallback (`NotoSansSC`), so Chinese text measures, renders, and embeds with a
   *  known face out of the box instead of an arbitrary (and server-absent) system
   *  one. Set to a registered custom font's family (see `fonts`) to override it, or
   *  to `""` to opt out and keep the browser's on-screen system fallback (CJK then
   *  renders as tofu in PDF export). */
  fallbackFont?: string;
  /** Family name of the font to use for Arabic runs. Defaults to the bundled Arabic
   *  fallback (`NotoSansArabic`), so Arabic text measures, renders (with correct
   *  contextual joining forms via GSUB), and embeds with a known face out of the box.
   *  Set to a registered custom font's family to override, or `""` to opt out. */
  arabicFallbackFont?: string;
  /** Family name of the font to use for Hebrew runs. Defaults to the bundled Hebrew
   *  fallback (`NotoSansHebrew`), so Hebrew text measures, renders, and embeds with a
   *  known face out of the box instead of rendering as tofu ("x") in PDF export.
   *  Set to a registered custom font's family to override, or `""` to opt out. */
  hebrewFallbackFont?: string;
}

/** Editor behavior tuning — every field optional. Pass to `WordCanvas({ behavior })`. */
export interface EditorBehavior {
  /** Multiplicative zoom step for +/- and Ctrl+wheel. Default 1.1. */
  zoomStep?: number;
  /** Absolute minimum zoom. Default 0.25. */
  zoomMin?: number;
  /** Absolute maximum zoom. Default 5. */
  zoomMax?: number;
  /** Indent / outdent step in px per toolbar press. Default 36. */
  indentStepPx?: number;
  /** Default drawing-grid spacing in px (view.gridSpacingPx still overrides). Default 24. */
  gridSpacingPx?: number;
}

/** Fully-populated theme (no optional fields) — what the renderer/rulers read. */
export interface ResolvedTheme {
  canvasBackground: string;
  grid: string;
  gridMesh: string;
  externalLink: string;
  accent: string;
  footnoteRule: string;
  formattingMark: string;
  tocLeader: string;
  imagePlaceholder: string;
  columnSeparator: string;
  caret: string;
  searchHighlight: string;
  reviewPinResolved: string;
  reviewPinStroke: string;
  pageGapPx: number;
  ruler: Required<RulerTheme>;
}

/** Fully-populated behavior. */
export type ResolvedBehavior = Required<EditorBehavior>;

/** The library's built-in theme. Seeded from paintStyle.ts (kept in lockstep
 *  with the PDF painter) plus the editor-chrome colors. */
export const DEFAULT_THEME: ResolvedTheme = {
  canvasBackground: "#e8eaed",
  grid: DEFAULT_GRID_COLOR,
  gridMesh: "rgba(0,0,0,0.08)",
  externalLink: EXTERNAL_LINK_COLOR,
  accent: ACCENT_BLUE,
  footnoteRule: FOOTNOTE_RULE_COLOR,
  formattingMark: FORMATTING_MARK_COLOR,
  tocLeader: TOC_LEADER_COLOR,
  imagePlaceholder: IMAGE_PLACEHOLDER_COLOR,
  columnSeparator: COLUMN_SEPARATOR_COLOR,
  caret: "#1a1a2e",
  searchHighlight: "rgba(251, 188, 4, 0.45)",
  reviewPinResolved: "#9aa0a6",
  reviewPinStroke: "#fff",
  pageGapPx: 24,
  ruler: {
    bg: "#c7cdd6",
    content: "#ffffff",
    line: "#8a8f98",
    label: "#605e5c",
    font: "9px 'Segoe UI', sans-serif",
  },
};

/** The library's built-in behavior. */
export const DEFAULT_BEHAVIOR: ResolvedBehavior = {
  zoomStep: ZOOM_STEP,
  zoomMin: ZOOM_MIN,
  zoomMax: ZOOM_MAX,
  indentStepPx: INDENT_STEP_PX,
  gridSpacingPx: 24,
};

/** A ready-made dark-canvas theme (darkens the chrome/gutter; pages stay white
 *  for print fidelity). Spread + tweak it for your own theme. */
export const darkCanvasTheme: EditorTheme = {
  canvasBackground: "#202124",
  ruler: { bg: "#2a2d31", content: "#3c4043", line: "#5f6368", label: "#bdc1c6" },
};

/** Merge a partial theme over the built-in default (nested `ruler` merges too). */
export function resolveTheme(t?: EditorTheme): ResolvedTheme {
  if (!t) return DEFAULT_THEME;
  return {
    ...DEFAULT_THEME,
    ...stripUndefined(t),
    ruler: { ...DEFAULT_THEME.ruler, ...(t.ruler ? stripUndefined(t.ruler) : {}) },
  };
}

/** Merge a partial behavior over the built-in default. */
export function resolveBehavior(b?: EditorBehavior): ResolvedBehavior {
  return b ? { ...DEFAULT_BEHAVIOR, ...stripUndefined(b) } : DEFAULT_BEHAVIOR;
}

/** The library's built-in (empty) fonts config. */
export const DEFAULT_FONTS: ResolvedFontsConfig = { disableBuiltin: [], fonts: [] };

/** Normalize the public partial `fonts` option into the fully-populated form.
 *  Deep-clones nested `faces`/`sizing` (and never returns the shared DEFAULT_FONTS
 *  object) so later mutation of the caller's `fonts` option can't desync the
 *  already-loaded editor fonts from the export-side config. */
export function resolveFonts(f?: FontsConfig): ResolvedFontsConfig {
  if (!f) return { disableBuiltin: [], fonts: [] };
  return {
    disableBuiltin: f.disableBuiltin ? [...f.disableBuiltin] : [],
    fonts: f.fonts ? f.fonts.map((d) => ({ ...d, faces: { ...d.faces }, sizing: { ...d.sizing } })) : [],
  };
}

/** The resolved, instance-scoped configuration the editor app threads downward. */
export interface ResolvedConfig {
  theme: ResolvedTheme;
  behavior: ResolvedBehavior;
  /** Typography overrides (for new/blank docs + the fallback stylesheet). */
  typography: DefaultStyleOverrides;
  /** The default stylesheet generated from `typography` — used for new/blank
   *  documents and as the fallback when a document carries none. */
  stylesheet: Stylesheet;
  /** Custom fonts + toolbar disables (per-instance toolbar; global font registry). */
  fonts: ResolvedFontsConfig;
  /** CJK locale + fallback-font tuning (applied to the global analyzer at mount). */
  cjk: CjkConfig;
  /** Develop mode: when true, reveal the "Developer" ribbon tab whose Document-tree
   *  inspector lets a developer browse the parsed model. Off by default; even when
   *  on, nothing dev-related runs until the inspector is opened from that tab. */
  develop: boolean;
  /** Show the "Organize Pages" ribbon button (Layout tab) that opens the visual
   *  page-reorder overlay. Default true; set false to hide it for an embed. */
  organizePages: boolean;
  /** Show the floating mini-toolbar (quick B/I/U, font, size, colour) above a text
   *  selection, so common formatting is reachable without the ribbon. Default true. */
  floatingToolbar: boolean;
}

export interface EditorConfigInput {
  theme?: EditorTheme | undefined;
  overrideDefaultStyles?: DefaultStyleOverrides | undefined;
  behavior?: EditorBehavior | undefined;
  fonts?: FontsConfig | undefined;
  /** CJK locale + fallback font. */
  cjk?: CjkConfig | undefined;
  /** Reveal the "Developer" ribbon tab + Document-tree inspector. Default false. */
  develop?: boolean | undefined;
  /** Show the "Organize Pages" reorder overlay button. Default true. */
  organizePages?: boolean | undefined;
  /** Show the floating mini-toolbar above a text selection. Default true. */
  floatingToolbar?: boolean | undefined;
}

/** Resolve the public partial options into the fully-populated internal config. */
export function resolveConfig(input: EditorConfigInput = {}): ResolvedConfig {
  const typography = input.overrideDefaultStyles ? stripUndefined(input.overrideDefaultStyles) : {};
  return {
    theme: resolveTheme(input.theme),
    behavior: resolveBehavior(input.behavior),
    typography,
    stylesheet: makeDefaultStylesheet(typography),
    fonts: resolveFonts(input.fonts),
    cjk: resolveCjk(input.cjk),
    develop: input.develop ?? false,
    organizePages: input.organizePages ?? true,
    floatingToolbar: input.floatingToolbar ?? true,
  };
}

/** Resolve the public partial `cjk` option, defaulting the CJK, Arabic, and Hebrew
 *  fallback fonts to their bundled faces so non-Latin text renders out of the box. An
 *  explicit `""` is preserved (the opt-out: keeps the browser's system fallback on screen). */
export function resolveCjk(c?: CjkConfig): CjkConfig {
  const base = c ? stripUndefined(c) : {};
  return {
    ...base,
    fallbackFont: base.fallbackFont ?? CJK_FONT_FAMILY,
    arabicFallbackFont: base.arabicFallbackFont ?? ARABIC_FONT_FAMILY,
    hebrewFallbackFont: base.hebrewFallbackFont ?? HEBREW_FONT_FAMILY,
  };
}

/** Drop keys whose value is `undefined` so they don't clobber a default during
 *  a spread (exactOptionalPropertyTypes-friendly). */
function stripUndefined<T extends object>(o: T): T {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(o)) if (v !== undefined) out[k] = v;
  return out as T;
}
