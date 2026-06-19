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

/** The resolved, instance-scoped configuration the editor app threads downward. */
export interface ResolvedConfig {
  theme: ResolvedTheme;
  behavior: ResolvedBehavior;
  /** Typography overrides (for new/blank docs + the fallback stylesheet). */
  typography: DefaultStyleOverrides;
  /** The default stylesheet generated from `typography` — used for new/blank
   *  documents and as the fallback when a document carries none. */
  stylesheet: Stylesheet;
}

export interface EditorConfigInput {
  theme?: EditorTheme | undefined;
  overrideDefaultStyles?: DefaultStyleOverrides | undefined;
  behavior?: EditorBehavior | undefined;
}

/** Resolve the public partial options into the fully-populated internal config. */
export function resolveConfig(input: EditorConfigInput = {}): ResolvedConfig {
  const typography = input.overrideDefaultStyles ? stripUndefined(input.overrideDefaultStyles) : {};
  return {
    theme: resolveTheme(input.theme),
    behavior: resolveBehavior(input.behavior),
    typography,
    stylesheet: makeDefaultStylesheet(typography),
  };
}

/** Drop keys whose value is `undefined` so they don't clobber a default during
 *  a spread (exactOptionalPropertyTypes-friendly). */
function stripUndefined<T extends object>(o: T): T {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(o)) if (v !== undefined) out[k] = v;
  return out as T;
}
