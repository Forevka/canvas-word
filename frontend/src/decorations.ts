// Public decoration/overlay API — types + the pure spec→geometry resolver.
//
// An embedder describes visuals in DOCUMENT space (a range to highlight, a
// position to badge) via `handle.setDecorations(specs)`; the editor resolves them
// against the live layout tree into page-local rects and paints them. Like the
// review overlays (see review/decorate.ts), this reuses geometry.selectionRects /
// caretRect, so it NEVER measures text and NEVER triggers relayout — a pure
// repaint feed. The editor owns DPR, zoom, virtualization, and scroll; authors
// never touch device pixels.
//
// This module is DOM-free. The resolver takes its geometry helpers as an argument
// (defaulting to the real ones) so the routing/defaulting logic is unit-testable
// without a full LayoutTree.

import type { DocPosition, DocSelection } from "@cw/shared";
import type { LayoutTree } from "./layout/layoutTree";
import type { GeoScope, Rect } from "./layout/geometry";

/** Shared fields for every decoration. `id` is for the author's own bookkeeping
 *  (dedup / update); the editor doesn't interpret it. `color` is any CSS color. */
export interface DecorationBase {
  id?: string;
  color: string;
  /** Make the decoration interactive: clicking it invokes this instead of placing
   *  a caret, and the pointer shows a `pointer` cursor while hovering it (like a
   *  comment pin). Omit for a purely visual decoration that never intercepts
   *  clicks. `event` carries the click's viewport coordinates. */
  onClick?: (event: { clientX: number; clientY: number }) => void;
}

/** A decoration spanning a document range. `highlight` fills under the text;
 *  `underline` draws a line at the baseline; `box` strokes a rect around it —
 *  both `underline`/`box` render over the text. */
export interface RangeDecoration extends DecorationBase {
  type: "highlight" | "underline" | "box";
  range: DocSelection;
  /** Stroke thickness in px for `underline`/`box` (default 1.6 / 1). */
  thickness?: number;
  /** Fill opacity for `highlight`, 0..1 (default 0.4). */
  opacity?: number;
}

/** A small marker anchored at a single document position (e.g. a footnote-style
 *  number, a status dot). Rendered over the text. */
export interface BadgeDecoration extends DecorationBase {
  type: "badge";
  at: DocPosition;
  /** Short text inside the badge; omit for a plain dot. */
  label?: string;
}

/** One item passed to `setDecorations`. */
export type DecorationSpec = RangeDecoration | BadgeDecoration;

// ---- resolved (page-local) forms consumed by the paint layer ----------------

/** A page-local colored rect (a resolved highlight/underline/box). `specIndex`
 *  points back at the source spec (for click dispatch); `interactive` is true when
 *  that spec has an `onClick`. */
export interface ResolvedDecoBox extends Rect {
  color: string;
  thickness?: number;
  opacity?: number;
  specIndex: number;
  interactive: boolean;
}

/** A page-local resolved badge. */
export interface ResolvedBadge {
  pageIndex: number;
  x: number;
  y: number;
  color: string;
  label?: string;
  specIndex: number;
  interactive: boolean;
}

/** Everything the renderer needs, bucketed by paint pass (under-text fills vs.
 *  over-text lines/boxes/badges). */
export interface ResolvedDecorations {
  highlights: ResolvedDecoBox[];
  underlines: ResolvedDecoBox[];
  boxes: ResolvedDecoBox[];
  badges: ResolvedBadge[];
}

export const EMPTY_DECORATIONS: ResolvedDecorations = { highlights: [], underlines: [], boxes: [], badges: [] };

/** The geometry surface the resolver depends on. Passed in by the editor (the
 *  real geometry helpers) — and injectable, so the routing/defaulting logic is
 *  unit-testable without a real LayoutTree (and without pulling the DOM-touching
 *  geometry module into a node test). */
export interface DecorationGeometry {
  selectionRects(tree: LayoutTree, sel: DocSelection, scope?: GeoScope): Rect[];
  caretRect(tree: LayoutTree, pos: DocPosition, scope?: GeoScope): { pageIndex: number; x: number; y: number; height: number } | null;
}

const isRange = (s: DecorationSpec): s is RangeDecoration => s.type === "highlight" || s.type === "underline" || s.type === "box";

/** Resolve embedder decoration specs into page-local paint data against the
 *  current layout tree. Metric-neutral (never re-breaks lines). Invalid specs
 *  (empty color, a range/position the layout can't resolve) are skipped. */
export function resolveDecorations(
  specs: readonly DecorationSpec[],
  tree: LayoutTree,
  scope: GeoScope | undefined,
  geo: DecorationGeometry,
): ResolvedDecorations {
  const out: ResolvedDecorations = { highlights: [], underlines: [], boxes: [], badges: [] };
  for (let specIndex = 0; specIndex < specs.length; specIndex++) {
    const spec = specs[specIndex]!;
    if (!spec.color) continue;
    const interactive = typeof spec.onClick === "function";
    if (isRange(spec)) {
      const rects = geo.selectionRects(tree, spec.range, scope);
      if (rects.length === 0) continue;
      if (spec.type === "highlight") {
        const opacity = spec.opacity ?? 0.4;
        for (const r of rects) out.highlights.push({ ...r, color: spec.color, opacity, specIndex, interactive });
      } else if (spec.type === "underline") {
        const thickness = spec.thickness ?? 1.6;
        for (const r of rects) out.underlines.push({ ...r, color: spec.color, thickness, specIndex, interactive });
      } else {
        const thickness = spec.thickness ?? 1;
        for (const r of rects) out.boxes.push({ ...r, color: spec.color, thickness, specIndex, interactive });
      }
    } else {
      const c = geo.caretRect(tree, spec.at, scope);
      if (!c) continue;
      out.badges.push({ pageIndex: c.pageIndex, x: c.x, y: c.y, color: spec.color, specIndex, interactive, ...(spec.label !== undefined ? { label: spec.label } : {}) });
    }
  }
  return out;
}

/** Badge geometry for hit-testing: a labelled badge is a pill of this height sized
 *  to its label; a plain dot is a small disc. The paint layer uses the same
 *  numbers (see renderer's custom-decoration pass). */
export const BADGE_HEIGHT = 14;
export const BADGE_DOT_RADIUS = 3.5;
