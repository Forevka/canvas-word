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
import type { DocSelection } from "@cw/shared";
import type { CurrentFormat } from "./index";
import type { RibbonActionContext } from "./ribbon";
import type { AnchorRect } from "./ui/floatingBarPosition";
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

/** Chrome preset (critique Move 1). `'ribbon'` is the classic Word-style tabbed
 *  ribbon; `'minimal'` demotes it to a quiet ~44px command bar (title + save state,
 *  undo/redo, style picker, the six core formatting commands, an insert `＋`, and an
 *  overflow that opens the command palette) with everything else reaching the user
 *  through the contextual bar, the Inspector and the command palette. The ribbon
 *  ships as a switchable skin so enterprise migrations can keep it. */
export type ChromePreset = "ribbon" | "minimal";

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

// ---- floating format toolbar (Word's selection mini-toolbar) ----------------

/** A built-in floating-toolbar control. `"|"` renders a separator. */
export type FloatingToolbarBuiltin =
  | "font"
  | "fontSize"
  | "bold"
  | "italic"
  | "underline"
  | "strikethrough"
  | "color"
  | "highlight"
  | "clearFormat";

/** A custom floating-toolbar button. Its `onClick` receives the same
 *  `RibbonActionContext` a custom ribbon button gets (editor handle + macro
 *  helpers), so the two customization surfaces share button code. */
export interface FloatingToolbarButtonSpec {
  /** Stable id (namespace it to avoid clashes, e.g. "myco.upper"). */
  id: string;
  /** Raw innerHTML for the button face — an SVG string, emoji, or text. */
  icon?: string;
  /** Text label (used when `icon` is omitted). */
  label?: string;
  /** Hover tooltip. */
  tooltip?: string;
  /** Invoked on click with the editor handle + macro helpers. */
  onClick: (ctx: RibbonActionContext) => void;
  /** Optional pressed-state predicate, re-evaluated on every selection change. */
  active?: (fmt: CurrentFormat) => boolean;
}

/** One entry in the floating toolbar: a built-in control id, a `"|"` separator,
 *  or a custom button. */
export type FloatingToolbarItem = FloatingToolbarBuiltin | "|" | FloatingToolbarButtonSpec;

/** The object form of the `floatingToolbar` option (the boolean is shorthand for
 *  `{ enabled }`). */
export interface FloatingToolbarOptions {
  /** Show the toolbar at all. Default true. */
  enabled?: boolean;
  /** Also show it at a collapsed caret (no selection), not only over a range —
   *  Word only shows it on selection, so this is off by default. */
  onCaret?: boolean;
  /** Which controls to show, in order (built-in ids, `"|"` separators, and custom
   *  buttons). Omit for the full built-in set. An empty array falls back to the
   *  default set. */
  buttons?: FloatingToolbarItem[];
}

/** The public `floatingToolbar` option: `false`/`true` toggles the default set,
 *  an object customizes it. */
export type FloatingToolbarConfig = boolean | FloatingToolbarOptions;

/** Fully-populated floating-toolbar config the editor app reads. */
export interface ResolvedFloatingToolbar {
  enabled: boolean;
  onCaret: boolean;
  buttons: FloatingToolbarItem[];
}

/** The built-in control order (Word-like: font · size · B/I/U/S · colour ·
 *  highlight · clear). */
export const DEFAULT_FLOATING_TOOLBAR_BUTTONS: FloatingToolbarItem[] = [
  "font",
  "|",
  "fontSize",
  "|",
  "bold",
  "italic",
  "underline",
  "strikethrough",
  "|",
  "color",
  "highlight",
  "|",
  "clearFormat",
];

/** Normalize the public partial `floatingToolbar` option into the fully-populated
 *  form. Booleans map to enabled/disabled with the default button set; an empty
 *  `buttons` array falls back to the default set (a fully-empty bar is never
 *  useful). */
export function resolveFloatingToolbar(input?: FloatingToolbarConfig): ResolvedFloatingToolbar {
  if (input === undefined || typeof input === "boolean") {
    return { enabled: input ?? true, onCaret: false, buttons: DEFAULT_FLOATING_TOOLBAR_BUTTONS };
  }
  return {
    enabled: input.enabled ?? true,
    onCaret: input.onCaret ?? false,
    buttons:
      input.buttons && input.buttons.length > 0 ? input.buttons : DEFAULT_FLOATING_TOOLBAR_BUTTONS,
  };
}

// ---- contextual floating toolbars (public embedder API) ---------------------

/** What an embedder-registered context toolbar sees to decide whether to show and
 *  where to anchor. Mirrors the signals the built-in bars use. */
export interface ToolbarContext {
  /** Character/paragraph format + context flags at the caret (imageSelected,
   *  inTable, inContentControl, …). */
  format: CurrentFormat;
  /** The live selection, or null when the editor isn't focused. */
  selection: DocSelection | null;
  /** True when a non-empty range is selected (vs a bare caret). */
  hasRange: boolean;
  /** The hyperlink URL at the caret, or null. */
  linkUrl: string | null;
  /** Viewport anchor rect of the selection start line / caret (the default anchor). */
  selectionRect(): AnchorRect | null;
  /** Viewport anchor rect of the selected image, or null. */
  objectRect(): AnchorRect | null;
}

/** A custom contextual floating toolbar an embedder registers via the
 *  `contextToolbars` option. Shown (above the selection) whenever `when(ctx)` is
 *  true and no higher-priority toolbar is active. */
export interface ContextToolbarSpec {
  /** Stable id (namespace it, e.g. "myco.table"). */
  id: string;
  /** Higher wins when several contexts are active. Built-ins use image=30,
   *  hyperlink=25, text=20; default 15 keeps custom bars below them. */
  priority?: number;
  /** Show predicate over the current context. */
  when: (ctx: ToolbarContext) => boolean;
  /** Buttons — the same custom-button shape as `floatingToolbar` / the ribbon. */
  buttons: FloatingToolbarButtonSpec[];
  /** Anchor rect resolver; defaults to the selection anchor (`ctx.selectionRect()`). */
  anchor?: (ctx: ToolbarContext) => AnchorRect | null;
}

/** Normalize the `contextToolbars` option (a plain copy — nothing to default at the
 *  array level; per-spec `priority`/`anchor` defaults are applied at wiring time). */
export function resolveContextToolbars(input?: ContextToolbarSpec[]): ContextToolbarSpec[] {
  return input ? [...input] : [];
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
  /** Chrome preset: the classic tabbed `'ribbon'` or the quiet `'minimal'` command
   *  bar (critique Move 1). Default `'ribbon'`. */
  chrome: ChromePreset;
  /** Floating mini-toolbar (quick formatting above a text selection): whether it's
   *  shown, whether it appears at a bare caret, and which controls it carries. */
  floatingToolbar: ResolvedFloatingToolbar;
  /** Embedder-registered contextual floating toolbars (in addition to the built-in
   *  image / hyperlink / text bars). */
  contextToolbars: ContextToolbarSpec[];
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
  /** Chrome preset — `'ribbon'` (classic tabbed ribbon, default) or `'minimal'`
   *  (quiet ~44px command bar; everything else via the contextual bar, Inspector
   *  and command palette). */
  chrome?: ChromePreset | undefined;
  /** Floating mini-toolbar above a text selection: `true`/`false` to toggle, or an
   *  object to customize which controls appear, their order, and caret behavior. */
  floatingToolbar?: FloatingToolbarConfig | undefined;
  /** Register custom contextual floating toolbars (shown for your own contexts,
   *  alongside the built-in image / hyperlink / text bars). */
  contextToolbars?: ContextToolbarSpec[] | undefined;
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
    chrome: input.chrome ?? "ribbon",
    floatingToolbar: resolveFloatingToolbar(input.floatingToolbar),
    contextToolbars: resolveContextToolbars(input.contextToolbars),
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
