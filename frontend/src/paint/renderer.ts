// Layer 3: paint — dumb and fast. Takes LayoutTree + precomputed selection rects,
// draws pixels. NEVER measures text here (that would defeat pretext), and never
// computes selection geometry (that's geometry.ts, fed through the wiring layer).
//
// Page virtualization: the scroll container holds exact-height placeholder divs
// (heights known from layout without painting). An IntersectionObserver with one
// viewport of overscan mounts a live <canvas> per visible page and unmounts the
// rest. Backing store = CSS size * devicePixelRatio.
//
// The caret is a DOM overlay div inside the page placeholder — blinking via CSS
// animation, so it never triggers a canvas repaint.

import type { LayoutTree, LineBox, Page, PlacedBlock, PlacedTableCell } from "../layout/layoutTree";
import type { CaretRect, Rect } from "../layout/geometry";
import { spaceMarkXs } from "../layout/geometry";
import type { CellBorder } from "@cw/shared";
import { DEFAULT_CHAR_STYLE } from "@cw/shared";
import { charStyleToFont } from "../layout/metrics";
import { paintMathBoxCanvas, paintMathBoxCanvasRaw } from "./paintMath";
import { MATH_FONT_FAMILY } from "../fonts/clones";
import { setActiveFontRegistry, type CustomFontRegistry } from "../fonts/customRegistry";
import {
  cellBorderDash,
  cellBorderWidth,
  decorationThickness,
  doubleBorderGap,
  FOOTNOTE_RULE_WIDTH_FRACTION,
  pageBorderSegments,
  leaderDash,
  leaderWidth,
  normalizeLinkBlue,
  runPaint,
  strikeOffset,
  TOC_LEADER_DASH,
  TOC_LEADER_GAP_PX,
  UNDERLINE_OFFSET_PX,
  underlinePlan,
  doubleUnderlineGap,
  underlineWavePoints,
  verticalShift,
} from "./paintStyle";
import type { RunPaint } from "./paintStyle";
import type { UnderlineStyle } from "@cw/shared";
import { DEFAULT_THEME, type ResolvedTheme } from "../config";
import { ZOOM_MAX, ZOOM_MIN } from "../uiConstants";

/** Paint a run's underline. The plain "single" style keeps the historical
 *  baseline-aligned filled rect (so existing renders don't shift); the richer
 *  styles (double/dotted/dashed/dotDash/wave/thick) stroke per their plan. */
function paintUnderline(
  ctx: CanvasRenderingContext2D,
  x: number,
  yTop: number,
  width: number,
  fontSizePx: number,
  rp: RunPaint,
): void {
  const style: UnderlineStyle = rp.underlineStyle;
  const th = decorationThickness(fontSizePx);
  if (style === "single") {
    ctx.fillStyle = rp.underlineColor;
    ctx.fillRect(x, yTop, width, th);
    return;
  }
  const plan = underlinePlan(style, fontSizePx);
  const yCenter = yTop + plan.thickness / 2;
  ctx.save();
  ctx.strokeStyle = rp.underlineColor;
  ctx.lineWidth = plan.thickness;
  ctx.setLineDash(plan.dash);
  ctx.beginPath();
  if (plan.wave) {
    const pts = underlineWavePoints(x, yCenter, width, fontSizePx);
    pts.forEach((p, i) => (i === 0 ? ctx.moveTo(p.x, p.y) : ctx.lineTo(p.x, p.y)));
  } else {
    ctx.moveTo(x, yCenter);
    ctx.lineTo(x + width, yCenter);
    if (plan.double) {
      const dy = doubleUnderlineGap(plan.thickness);
      ctx.moveTo(x, yCenter + dy);
      ctx.lineTo(x + width, yCenter + dy);
    }
  }
  ctx.stroke();
  ctx.restore();
}

export interface PagePoint {
  pageIndex: number;
  x: number;
  y: number;
  /** False when the click fell OUTSIDE the page rect and x/y were clamped to the
   *  nearest edge. Caret placement still uses the clamp (click-below → caret at
   *  end); object selection must NOT (an off-page click should deselect, never
   *  "hit" a background image snapped onto the edge). */
  inside: boolean;
}

/** A colored rect for review overlays (insertion underline, deletion strike,
 *  comment highlight, margin change-bar). x/y/width/height are page-local. */
export interface ReviewDecoBox extends Rect {
  color: string;
}

/** A comment marker drawn in the right margin; threadId drives hit-testing. */
export interface ReviewPin {
  pageIndex: number;
  x: number;
  y: number;
  color: string;
  threadId: string;
  resolved: boolean;
}

/** Paint-only review decorations, computed from the review layer + layout tree.
 *  Metric-neutral (underline/strike/highlight don't change line breaking) → a
 *  repaint, never a relayout, exactly like search highlights. */
export interface ReviewDecorations {
  comments: ReviewDecoBox[];
  inserts: ReviewDecoBox[];
  deletes: ReviewDecoBox[];
  changeBars: ReviewDecoBox[];
  pins: ReviewPin[];
}

const EMPTY_REVIEW_DECOS: ReviewDecorations = { comments: [], inserts: [], deletes: [], changeBars: [], pins: [] };

export interface PaintScheduler {
  setTree(tree: LayoutTree): void;
  /** Story-edit affordance: dim the body and mark the band boundary (Word's
   *  header/footer mode). Null restores normal painting. */
  setBandEditMode(band: "header" | "footer" | null): void;
  /** Precomputed highlight rects (geometry.selectionRects). Repaints affected pages. */
  setSelectionRects(rects: Rect[]): void;
  /** Find-bar match highlights — painted under the selection. */
  setSearchRects(rects: Rect[]): void;
  /** Track-changes + comment overlays (insertion underline, deletion strike,
   *  comment highlight + margin pins, change bars). Replaces the whole set. */
  setReviewDecorations(decos: ReviewDecorations): void;
  /** The comment thread whose margin pin is at a client point, or null. */
  reviewPinAt(clientX: number, clientY: number): string | null;
  /** Content-control adornment as an ordered OUTER→INNER stack of layers, so
   *  nested controls render as concentric frames with a breadcrumb tab. A single
   *  control is just a one-layer stack. Null/empty clears. */
  setSdtAdornment(layers: { rects: Rect[]; label: string }[] | null): void;
  /** Inline/block FIELD focus adornment: gray field-shading fill + a labelled tab
   *  (Word's field highlight). Distinct from the SDT frame. Null clears. */
  setFieldAdornment(adorn: { rects: Rect[]; label: string } | null): void;
  /** Develop-mode Document-tree inspector highlight: a devtools-blue overlay box +
   *  a labelled tab over a hovered node's painted region. Painted ABOVE the other
   *  adornments so it reads as an ephemeral inspection cue. Null clears. */
  setInspectorRects(adorn: { rects: Rect[]; label: string } | null): void;
  /** Develop-mode layout overlay toggle (block/line/fragment boxes, baselines,
   *  margins, cells, page info). Repaints all live pages. */
  setDebugOverlay(kind: string, on: boolean): void;
  /** Highlight the column boundary the pointer is poised to drag (a vertical
   *  accent line + soft grab-zone band over the table chunk). Null clears. */
  setColumnGuide(guide: ColumnGuide | null): void;
  /** Highlight the row boundary the pointer is poised to drag (a horizontal accent
   *  line + soft grab-zone band over the table chunk). Null clears. */
  setRowGuide(guide: RowGuide | null): void;
  setCaret(caret: CaretRect | null): void;
  /** Remote collaborators' carets (DOM overlays with name flags). Replaces the
   *  whole set each call; pass [] to clear. */
  setRemoteCarets(carets: RemoteCaret[]): void;
  /** Coalesce: mark pages dirty now, paint once on the next animation frame. */
  invalidatePages(pageIndexes: number[]): void;
  /** Map a client (viewport) point onto page-local coordinates. Clamps into the
   *  nearest page — drags above/below/between pages still resolve. */
  clientToPage(clientX: number, clientY: number): PagePoint | null;
  /** Container-relative position of a caret rect (for the IME proxy + scrolling). */
  caretToContainer(caret: CaretRect): { left: number; top: number } | null;
  /** Page placeholder element — host for DOM overlays (object selection frame). */
  getPageElement(pageIndex: number): HTMLElement | null;
  /** Scroll the container to reveal the caret. "nearest" (default) scrolls the
   *  minimum amount; "center" vertically centers the target in the viewport. */
  ensureVisible(caret: CaretRect, align?: "nearest" | "center"): void;
  /** Presentational zoom (1 = 100%). Scales pages, canvases, caret — NOT the
   *  layout (document coords are unchanged), so no relayout. Clamped to [.25, 5]. */
  setZoom(zoom: number): void;
  getZoom(): number;
  /** Drawing-grid overlay: a light gridline mesh painted on every page (for
   *  precise object placement). Presentational only — no layout/model change. */
  setShowGrid(show: boolean): void;
  getShowGrid(): boolean;
  /** Whether dragging an anchored object snaps to the grid. Drag-time only, so
   *  toggling it triggers no repaint. */
  setSnapToGrid(snap: boolean): void;
  getSnapToGrid(): boolean;
  /** Grid step in document px (96dpi). Drives both the drawn mesh and snapping. */
  setGridSpacing(px: number): void;
  getGridSpacing(): number;
  /** Non-printing formatting marks (space dots, tab arrows, pilcrows, line-break
   *  arrows). Pure overlay over existing geometry — repaint, never relayout. */
  setShowFormattingMarks(show: boolean): void;
  getShowFormattingMarks(): boolean;
  destroy(): void;
}

// Selection is painted with a "difference" blend over a dark-gray source at
// partial alpha — so it adapts to whatever is beneath it: a soft cool tint on
// the white page, a clearly lighter/inverted band on a shaded cell (e.g. a blue
// table header), always visible without burying the text (it paints UNDER the
// glyphs, above cell fills).
const SELECTION_DIFF_COLOR = "rgb(64, 64, 64)";
const SELECTION_DIFF_ALPHA = 0.5;

// Content-control / field adornment chrome shared geometry. The two adornment
// blocks differ in colour + rect treatment (SDT strokes a frame, fields fill +
// stroke a shade) but share the title-tab height, label font, and text offsets.
const ADORNMENT_TAB_H = 15;
const ADORNMENT_LABEL_FONT = "10px Arial";
/** Horizontal padding added to the measured label width for the tab box. */
const ADORNMENT_LABEL_PAD_X = 10;
/** Label baseline insets within the tab: +x from the tab left, +y from its top. */
const ADORNMENT_LABEL_OFFSET_X = 5;
const ADORNMENT_LABEL_OFFSET_Y = 11;

// Review comment pin (right-margin marker). Radius + vertical offset are shared
// between the painted disc and reviewPinAt's hit-testing so they stay aligned.
const PIN_RADIUS_PX = 5;
const PIN_OFFSET_PX = 6;

/** Stroke a page's w:pgBorders box. Shares geometry with the PDF painter via
 *  pageBorderSegments so the two backends stay in lockstep. */
function paintPageBorders(ctx: CanvasRenderingContext2D, page: Page): void {
  if (!page.pageBorders) return;
  const segs = pageBorderSegments(page.pageBorders, page.widthPx, page.heightPx, page.marginPx);
  for (const s of segs) {
    ctx.strokeStyle = s.color;
    ctx.lineWidth = s.widthPx;
    ctx.setLineDash(s.dash);
    ctx.beginPath();
    ctx.moveTo(s.x1, s.y1);
    ctx.lineTo(s.x2, s.y2);
    ctx.stroke();
  }
  ctx.setLineDash([]);
}

/** A remote collaborator's presence to render: their caret (focus) + any
 *  selection highlight rects, in page-local coords, with their color/name. */
export interface RemoteCaret {
  siteId: string;
  color: string;
  label: string;
  rect: CaretRect | null;
  rects?: Rect[];
}

/** A column-resize hover guide: the boundary x and the table chunk's vertical
 *  span on `pageIndex`, all in page-local document coords. */
export interface ColumnGuide {
  pageIndex: number;
  x: number;
  y: number;
  height: number;
}

/** A row-resize hover guide: the boundary y and the table chunk's horizontal span
 *  on `pageIndex`, all in page-local document coords. */
export interface RowGuide {
  pageIndex: number;
  y: number;
  x: number;
  width: number;
}

let caretCssInjected = false;
function injectCaretCss(): void {
  if (caretCssInjected) return;
  caretCssInjected = true;
  const style = document.createElement("style");
  style.textContent =
    "@keyframes cw-caret-blink{0%,55%{opacity:1}56%,100%{opacity:0}}" +
    // Caret color is set per-instance inline (caretEl.style.background = theme.caret).
    ".cw-caret{position:absolute;width:2px;pointer-events:none;animation:cw-caret-blink 1.06s step-end infinite;}" +
    ".cw-rcaret{position:absolute;width:2px;pointer-events:none;z-index:3;}" +
    ".cw-rcaret .flag{position:absolute;top:-13px;left:-1px;height:13px;display:flex;align-items:center;" +
    "font:600 10px/1 'Segoe UI',Roboto,sans-serif;color:#fff;padding:0 4px;border-radius:3px 3px 3px 0;white-space:nowrap;}" +
    ".cw-rsel{position:absolute;pointer-events:none;z-index:2;opacity:0.24;border-radius:1px;}";
  document.head.appendChild(style);
}

/** Optional render-surface tuning. Defaults reproduce the full editor chrome. */
export interface PaintLayerOptions {
  /** Draw page chrome — the drop shadow, inter-page gap, and wrap padding. Set
   *  false for embedded previews / child documents that should render the bare
   *  page(s) flush in their host element (no shadow, no surrounding gap). */
  chrome?: boolean;
  /** Resolved color theme (omit ⇒ the library default). Per-instance. */
  theme?: ResolvedTheme;
  /** Absolute zoom clamp (omit ⇒ built-in 0.25/5). */
  zoomMin?: number;
  zoomMax?: number;
  /** This editor instance's custom-font registry — asserted active for each paint
   *  pass so charStyleToFont resolves against this instance's fonts (not another
   *  WordCanvas mount's). Omit ⇒ the process-default registry. */
  fontRegistry?: CustomFontRegistry;
}

/** Serialize a page's paint inputs for change detection in setTree. Includes the
 *  whole Page (body blocks, bands, page chrome) — over-inclusive on purpose: a
 *  stray field can only trigger a harmless extra repaint, never a missed one.
 *  Overlays (selection, search, review, adornments) have their own dirty paths,
 *  so they are intentionally NOT part of this signature. */
function pageSignature(p: Page): string {
  return JSON.stringify(p);
}

export function createPaintLayer(container: HTMLElement, opts: PaintLayerOptions = {}): PaintScheduler {
  injectCaretCss();
  const theme = opts.theme ?? DEFAULT_THEME;
  const zoomMin = opts.zoomMin ?? ZOOM_MIN;
  const zoomMax = opts.zoomMax ?? ZOOM_MAX;
  // Page chrome (shadow + gaps) is on by default; previews turn it off so the
  // bare page sits flush in its host. The gap also feeds clientToPage hit-testing.
  const chrome = opts.chrome !== false;
  const fontRegistry = opts.fontRegistry;
  const gap = chrome ? theme.pageGapPx : 0;
  let tree: LayoutTree | null = null;
  let selectionRects: Rect[] = [];
  let searchRects: Rect[] = [];
  let reviewDecos: ReviewDecorations = EMPTY_REVIEW_DECOS;
  let sdtAdorn: { rects: Rect[]; label: string }[] | null = null;
  let fieldAdorn: { rects: Rect[]; label: string } | null = null;
  let inspectorAdorn: { rects: Rect[]; label: string } | null = null;
  // Develop-mode layout overlays (block/line/fragment boxes, baselines, margins,
  // cells, page info) — a Set of enabled kinds, drawn last on every page.
  const debugOverlays = new Set<string>();
  let bandEditMode: "header" | "footer" | null = null;

  const pagesWrap = document.createElement("div");
  pagesWrap.style.cssText = `display:flex;flex-direction:column;align-items:center;gap:${gap}px;padding:${gap}px 0;`;
  container.appendChild(pagesWrap);

  const caretEl = document.createElement("div");
  caretEl.className = "cw-caret";
  caretEl.style.background = theme.caret;
  caretEl.style.display = "none";

  const placeholders: HTMLDivElement[] = [];
  const liveCanvases = new Map<number, HTMLCanvasElement>();
  const dirty = new Set<number>();
  // Per-page paint signature from the last setTree. A relayout rebuilds the whole
  // LayoutTree (fresh PlacedBlock objects), so we can't compare by reference —
  // instead we serialize each LIVE page's paint inputs and repaint only the pages
  // whose content actually changed, rather than every mounted canvas. Serializing
  // is strictly over-dirty-safe (an unrelated field change can only cause an extra
  // repaint, never a missed one) and bounded to the few live pages.
  const pageSigs = new Map<number, string>();
  let rafId: number | null = null;
  let zoom = 1;
  // Drawing-grid view state (mirrors `zoom` — presentational, no model change).
  let showGrid = false;
  let snapToGrid = false;
  let gridSpacingPx = 24; // 1/4 inch @96dpi
  // Show non-printing formatting marks (spaces, tabs, paragraph ends, breaks).
  // Pure overlay over the already-positioned fragments — no relayout.
  let showFormattingMarks = false;
  let lastCaret: CaretRect | null = null;
  let colGuide: ColumnGuide | null = null;
  let rowGuide: RowGuide | null = null;
  // Remote collaborators' presence: per siteId a caret overlay (colored bar +
  // name flag) plus selection-highlight rect overlays, repositioned on zoom.
  const remoteCaretEls = new Map<string, { caret: HTMLDivElement; sels: HTMLDivElement[] }>();
  let lastRemoteCarets: RemoteCaret[] = [];

  const observer = new IntersectionObserver(
    (entries) => {
      for (const entry of entries) {
        const index = Number((entry.target as HTMLElement).dataset["page"]);
        if (entry.isIntersecting) mountPage(index);
        else unmountPage(index);
      }
    },
    { root: container, rootMargin: "100% 0px" }, // one viewport of overscan
  );

  // A document loaded in a hidden/background tab can paint gray image
  // placeholders (rAF is throttled) before its bitmaps decode; the load events
  // fire while throttled. Repaint everything when the tab comes back to the fore.
  const onVisible = (): void => {
    if (!document.hidden) repaintAllLive();
  };
  document.addEventListener("visibilitychange", onVisible);

  function stylePlaceholder(ph: HTMLDivElement, page: Page): void {
    ph.dataset["page"] = String(page.index);
    const shadow = chrome ? "box-shadow:0 1px 4px rgba(0,0,0,.25);" : "";
    ph.style.cssText = `position:relative;width:${page.widthPx * zoom}px;height:${page.heightPx * zoom}px;background:#fff;${shadow}flex-shrink:0;`;
  }

  // Reconcile the placeholder elements against tree.pages IN PLACE — surplus
  // trimmed from the tail, deficit appended — rather than tearing the whole set
  // down. Surviving placeholders keep their DOM identity (and any overlay child:
  // the caret, an active object-selection frame, remote-presence carets). That
  // matters during an image-resize drag: the resize handle holds pointer
  // capture, and removing its host page from the document would implicitly
  // release that capture and abort the drag. Re-rendering the same page set is
  // also cheaper (no canvas remount / observer churn) than a full rebuild.
  function rebuildPlaceholders(): void {
    if (!tree) {
      observer.disconnect();
      liveCanvases.clear();
      caretEl.remove();
      pagesWrap.textContent = "";
      placeholders.length = 0;
      return;
    }
    const want = tree.pages.length;
    // Trim surplus from the tail (page count shrank).
    while (placeholders.length > want) {
      const index = placeholders.length - 1;
      const ph = placeholders.pop()!;
      observer.unobserve(ph);
      unmountPage(index); // drop its canvas from liveCanvases + dirty
      ph.remove();
    }
    // Append new placeholders for added pages.
    while (placeholders.length < want) {
      const ph = document.createElement("div");
      stylePlaceholder(ph, tree.pages[placeholders.length]!);
      pagesWrap.appendChild(ph);
      placeholders.push(ph);
      observer.observe(ph);
    }
    // Refresh geometry on every surviving placeholder — dims can change without
    // the count changing (a section resizing its pages) — and repaint any page
    // whose canvas is mounted so the new dimensions take effect.
    for (let i = 0; i < placeholders.length; i++) {
      stylePlaceholder(placeholders[i]!, tree.pages[i]!);
      if (liveCanvases.has(i)) dirty.add(i);
    }
  }

  function mountPage(index: number): void {
    if (liveCanvases.has(index) || !tree) return;
    const ph = placeholders[index];
    if (!ph) return;
    const canvas = document.createElement("canvas");
    canvas.style.cssText = "position:absolute;inset:0;width:100%;height:100%;";
    ph.prepend(canvas); // under the caret overlay if it lives on this page
    liveCanvases.set(index, canvas);
    dirty.add(index);
    // Record the signature it's about to be painted at, so the next setTree only
    // repaints it if its content actually changes after this mount.
    pageSigs.set(index, pageSignature(tree.pages[index]!));
    schedule();
  }

  function unmountPage(index: number): void {
    const canvas = liveCanvases.get(index);
    if (!canvas) return;
    canvas.remove();
    liveCanvases.delete(index);
    dirty.delete(index);
  }

  function schedule(): void {
    if (rafId !== null) return;
    rafId = requestAnimationFrame(() => {
      rafId = null;
      flush();
    });
  }

  function flush(): void {
    if (!tree) return;
    // Assert this instance's fonts for the whole (synchronous) paint pass so
    // charStyleToFont resolves against them even if another WordCanvas mount last
    // touched the active registry.
    if (fontRegistry) setActiveFontRegistry(fontRegistry);
    for (const index of dirty) {
      const canvas = liveCanvases.get(index);
      const page = tree.pages[index];
      if (canvas && page) paintPage(canvas, page);
    }
    dirty.clear();
  }

  // Reconcile the remote-presence DOM overlays against the current set (and
  // zoom): per siteId a caret (focus) + selection-highlight rects; drop absent.
  function applyRemoteCarets(carets: RemoteCaret[]): void {
    const seen = new Set<string>();
    for (const c of carets) {
      seen.add(c.siteId);
      let entry = remoteCaretEls.get(c.siteId);
      if (!entry) {
        const caret = document.createElement("div");
        caret.className = "cw-rcaret";
        caret.appendChild(Object.assign(document.createElement("div"), { className: "flag" }));
        entry = { caret, sels: [] };
        remoteCaretEls.set(c.siteId, entry);
      }
      // Caret at the focus position.
      const caret = entry.caret;
      const ph = c.rect ? placeholders[c.rect.pageIndex] : undefined;
      if (c.rect && ph) {
        if (caret.parentElement !== ph) ph.appendChild(caret);
        caret.style.display = "block";
        caret.style.background = c.color;
        caret.style.left = `${c.rect.x * zoom - 1}px`;
        caret.style.top = `${c.rect.y * zoom}px`;
        caret.style.height = `${c.rect.height * zoom}px`;
        const flag = caret.firstElementChild as HTMLElement;
        flag.textContent = c.label;
        flag.style.background = c.color;
        flag.style.display = c.label ? "flex" : "none";
      } else {
        caret.style.display = "none";
      }
      // Selection highlight rects (rebuilt each update — usually a handful).
      for (const s of entry.sels) s.remove();
      entry.sels = [];
      for (const r of c.rects ?? []) {
        const rph = placeholders[r.pageIndex];
        if (!rph) continue;
        const sd = document.createElement("div");
        sd.className = "cw-rsel";
        sd.style.background = c.color;
        sd.style.left = `${r.x * zoom}px`;
        sd.style.top = `${r.y * zoom}px`;
        sd.style.width = `${r.width * zoom}px`;
        sd.style.height = `${r.height * zoom}px`;
        rph.appendChild(sd);
        entry.sels.push(sd);
      }
    }
    for (const [siteId, entry] of remoteCaretEls) {
      if (!seen.has(siteId)) {
        entry.caret.remove();
        for (const s of entry.sels) s.remove();
        remoteCaretEls.delete(siteId);
      }
    }
  }

  function paintPage(canvas: HTMLCanvasElement, page: Page): void {
    // Backing store scales with zoom so every document px maps to zoom×dpr device
    // px — crisp at any zoom. The ctx transform folds in zoom too, so the paint
    // code keeps drawing in document coords (0…page.widthPx), unaware of zoom.
    const scale = (window.devicePixelRatio || 1) * zoom;
    const w = Math.round(page.widthPx * scale);
    const h = Math.round(page.heightPx * scale);
    if (canvas.width !== w || canvas.height !== h) {
      canvas.width = w;
      canvas.height = h;
    }
    const ctx = canvas.getContext("2d")!;
    ctx.setTransform(scale, 0, 0, scale, 0, 0);

    // 1. page background (w:background page color, else white)
    ctx.fillStyle = page.pageColorHex ?? "#fff";
    ctx.fillRect(0, 0, page.widthPx, page.heightPx);

    // 1·grid. drawing-grid mesh (under content) for precise object placement.
    // Drawn in document coords from the page top-left; hairlines stay 1 device
    // px crisp at any zoom (ctx is already scaled by `scale`).
    if (showGrid && gridSpacingPx > 0) {
      ctx.save();
      ctx.strokeStyle = theme.gridMesh;
      ctx.lineWidth = 1 / scale;
      ctx.beginPath();
      for (let x = gridSpacingPx; x < page.widthPx; x += gridSpacingPx) {
        const px = Math.round(x) + 0.5 / scale;
        ctx.moveTo(px, 0);
        ctx.lineTo(px, page.heightPx);
      }
      for (let y = gridSpacingPx; y < page.heightPx; y += gridSpacingPx) {
        const py = Math.round(y) + 0.5 / scale;
        ctx.moveTo(0, py);
        ctx.lineTo(page.widthPx, py);
      }
      ctx.stroke();
      ctx.restore();
    }

    // 1a. page borders (w:pgBorders) — behind content, inset per offsetFrom.
    if (page.pageBorders) paintPageBorders(ctx, page);

    // 1b. table cell background fills — the bottom layer, hoisted out of the
    // block painter so the highlight layers below sit ON TOP of them. Otherwise
    // an opaque cell fill (e.g. a blue header) paints over and buries the
    // selection/search highlights drawn before the blocks.
    for (const block of page.blocks) paintCellFills(ctx, block, page.index);
    if (page.header) for (const b of page.header) paintCellFills(ctx, b, page.index);
    if (page.footer) for (const b of page.footer) paintCellFills(ctx, b, page.index);
    // Paragraph shading (w:shd) — same bottom layer as cell fills.
    for (const block of page.blocks) paintParaFills(ctx, block, page.index);
    if (page.header) for (const b of page.header) paintParaFills(ctx, b, page.index);
    if (page.footer) for (const b of page.footer) paintParaFills(ctx, b, page.index);

    // 2a. search-match highlights (over fills, under text)
    ctx.fillStyle = theme.searchHighlight;
    for (const r of searchRects) {
      if (r.pageIndex === page.index) ctx.fillRect(r.x, r.y, r.width, r.height);
    }
    // 2b. selection highlights (over fills, under text) — adaptive "difference"
    // blend so the band stays visible on shaded cells, not just the white page.
    ctx.save();
    ctx.globalCompositeOperation = "difference";
    ctx.globalAlpha = SELECTION_DIFF_ALPHA;
    ctx.fillStyle = SELECTION_DIFF_COLOR;
    for (const r of selectionRects) {
      if (r.pageIndex === page.index) ctx.fillRect(r.x, r.y, r.width, r.height);
    }
    ctx.restore();

    // 2c. comment highlight bands (under text, author-tinted)
    ctx.globalAlpha = 0.16;
    for (const r of reviewDecos.comments) {
      if (r.pageIndex !== page.index) continue;
      ctx.fillStyle = r.color;
      ctx.fillRect(r.x, r.y, r.width, r.height);
    }
    ctx.globalAlpha = 1;

    // 2c. content-control adornment: gray frame + title tab (Word's active
    // control chrome). The tab text is UI chrome, not document text — measuring
    // it here doesn't violate the paint-never-measures-layout invariant.
    if (sdtAdorn) {
      // Concentric frames, OUTER→INNER. Deeper controls draw darker and are inset
      // a touch so nesting reads even when an inner control covers its outer whole.
      sdtAdorn.forEach((layer, depth) => {
        const onPage = layer.rects.filter((r) => r.pageIndex === page.index);
        if (onPage.length === 0) return;
        const pad = Math.max(0, 2.5 - depth * 1.5); // 2.5, 1.0, 0, …
        const shade = Math.max(96, 168 - depth * 32); // 168, 136, 104, …
        ctx.strokeStyle = `rgb(${shade},${shade},${shade})`;
        ctx.lineWidth = 1;
        for (const r of onPage) {
          ctx.strokeRect(r.x - pad, r.y - pad + 1, Math.max(2, r.width + pad * 2), Math.max(2, r.height + pad * 2 - 2));
        }
      });
      // One breadcrumb tab ("Outer › Inner"), anchored to the OUTERMOST layer's
      // first rect, drawn only on the page where that control starts.
      const outer = sdtAdorn[0]!;
      const globalFirst = outer.rects[0];
      if (globalFirst && globalFirst.pageIndex === page.index) {
        const label = sdtAdorn.map((l) => l.label).join(" › ");
        ctx.font = ADORNMENT_LABEL_FONT;
        const w = ctx.measureText(label).width + ADORNMENT_LABEL_PAD_X;
        const tx = globalFirst.x - 2.5;
        const ty = globalFirst.y - 1.5 - ADORNMENT_TAB_H;
        ctx.fillStyle = "#d8d8d8";
        ctx.beginPath();
        ctx.roundRect(tx, ty, w, ADORNMENT_TAB_H, [3, 3, 0, 0]);
        ctx.fill();
        ctx.fillStyle = "#3c4043";
        ctx.fillText(label, tx + ADORNMENT_LABEL_OFFSET_X, ty + ADORNMENT_LABEL_OFFSET_Y);
      }
    }

    // 2d. field adornment: Word's gray field shading + a labelled tab. Distinct
    // from the SDT frame (a filled shade vs a stroked frame) so fields read apart.
    if (fieldAdorn) {
      const onPage = fieldAdorn.rects.filter((r) => r.pageIndex === page.index);
      for (const r of onPage) {
        ctx.fillStyle = "rgba(200,200,200,0.25)";
        ctx.fillRect(r.x - 1, r.y - 1, r.width + 2, r.height + 2);
        ctx.strokeStyle = "#9aa7b8";
        ctx.lineWidth = 1;
        ctx.strokeRect(r.x - 1, r.y - 1, r.width + 2, r.height + 2);
      }
      const first = onPage[0];
      if (first && fieldAdorn.label) {
        ctx.font = ADORNMENT_LABEL_FONT;
        const w = ctx.measureText(fieldAdorn.label).width + ADORNMENT_LABEL_PAD_X;
        const tx = first.x - 1;
        const ty = first.y - 1 - ADORNMENT_TAB_H;
        ctx.fillStyle = "#c7d2e0";
        ctx.beginPath();
        ctx.roundRect(tx, ty, w, ADORNMENT_TAB_H, [3, 3, 0, 0]);
        ctx.fill();
        ctx.fillStyle = "#2b3a4a";
        ctx.fillText(fieldAdorn.label, tx + ADORNMENT_LABEL_OFFSET_X, ty + ADORNMENT_LABEL_OFFSET_Y);
      }
    }

    // 2e. develop-mode inspector highlight: a devtools-blue translucent fill +
    // solid stroke + a labelled tab over the hovered tree node's painted region.
    // Drawn last of the overlays (above selection/sdt/field) so it reads as an
    // ephemeral inspection cue, never as a persistent document decoration.
    if (inspectorAdorn) {
      const onPage = inspectorAdorn.rects.filter((r) => r.pageIndex === page.index);
      for (const r of onPage) {
        ctx.fillStyle = "rgba(56,121,217,0.18)";
        ctx.fillRect(r.x - 1, r.y - 1, r.width + 2, r.height + 2);
        ctx.strokeStyle = "#3879d9";
        ctx.lineWidth = 1;
        ctx.strokeRect(r.x - 1, r.y - 1, r.width + 2, r.height + 2);
      }
      const first = onPage[0];
      if (first && inspectorAdorn.label) {
        ctx.font = ADORNMENT_LABEL_FONT;
        const w = ctx.measureText(inspectorAdorn.label).width + ADORNMENT_LABEL_PAD_X;
        const tx = first.x - 1;
        const ty = first.y - 1 - ADORNMENT_TAB_H;
        ctx.fillStyle = "#3879d9";
        ctx.beginPath();
        ctx.roundRect(tx, ty, w, ADORNMENT_TAB_H, [3, 3, 0, 0]);
        ctx.fill();
        ctx.fillStyle = "#ffffff";
        ctx.fillText(inspectorAdorn.label, tx + ADORNMENT_LABEL_OFFSET_X, ty + ADORNMENT_LABEL_OFFSET_Y);
      }
    }

    // 3. blocks: text fragments (one fillText each), images, table grids
    ctx.textBaseline = "alphabetic";
    for (const block of page.blocks) paintBlock(ctx, block, page.index);

    // 3a. column separator rules (w:cols/@w:sep) — thin vertical lines in gaps.
    if (page.columnSeparatorsX) {
      ctx.strokeStyle = theme.columnSeparator;
      ctx.lineWidth = 1;
      for (const x of page.columnSeparatorsX) {
        const sx = Math.round(x) + 0.5;
        ctx.beginPath();
        ctx.moveTo(sx, page.contentTopPx);
        ctx.lineTo(sx, page.contentBottomPx);
        ctx.stroke();
      }
    }

    // 3b. footnote / endnote separator rule (1/3 content width, Word style)
    for (const ruleY of [page.footnoteRuleY, page.endnoteRuleY]) {
      if (ruleY === undefined) continue;
      const cw = page.widthPx - page.marginPx.left - page.marginPx.right;
      ctx.strokeStyle = theme.footnoteRule;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(page.marginPx.left, ruleY + 0.5);
      ctx.lineTo(page.marginPx.left + cw * FOOTNOTE_RULE_WIDTH_FRACTION, ruleY + 0.5);
      ctx.stroke();
    }

    // 3c. margin line numbers (w:lnNumType) — pre-measured, right-aligned labels.
    if (page.lineNumbers) {
      ctx.textBaseline = "alphabetic";
      for (const ln of page.lineNumbers) {
        ctx.font = charStyleToFont(ln.style);
        ctx.fillStyle = ln.style.color;
        ctx.fillText(ln.text, ln.x, ln.baseline);
      }
    }

    // 4. margin-band stories (headers/footers) — pre-laid-out rich blocks,
    //    painted through the same block painter, outside the selectable tree.
    if (page.header) for (const b of page.header) paintBlock(ctx, b, page.index);
    if (page.footer) for (const b of page.footer) paintBlock(ctx, b, page.index);

    // 4b. review track-changes overlays (over text): insertion underlines,
    //     deletion strikes, margin change-bars + comment pins.
    for (const r of reviewDecos.inserts) {
      if (r.pageIndex !== page.index) continue;
      ctx.fillStyle = r.color;
      ctx.fillRect(r.x, r.y + r.height - 2, r.width, 1.6); // underline at the baseline
    }
    for (const r of reviewDecos.deletes) {
      if (r.pageIndex !== page.index) continue;
      ctx.fillStyle = r.color;
      ctx.fillRect(r.x, r.y + r.height / 2 - 0.8, r.width, 1.6); // strike through the middle
    }
    for (const r of reviewDecos.changeBars) {
      if (r.pageIndex !== page.index) continue;
      ctx.fillStyle = r.color;
      ctx.fillRect(r.x, r.y, r.width, r.height);
    }
    for (const pin of reviewDecos.pins) {
      if (pin.pageIndex !== page.index) continue;
      ctx.beginPath();
      ctx.arc(pin.x, pin.y + PIN_OFFSET_PX, PIN_RADIUS_PX, 0, Math.PI * 2);
      ctx.fillStyle = pin.resolved ? theme.reviewPinResolved : pin.color;
      ctx.fill();
      ctx.lineWidth = 1;
      ctx.strokeStyle = theme.reviewPinStroke;
      ctx.stroke();
    }

    // 5. story-edit affordance: dim the body, dash the band boundary.
    if (bandEditMode) {
      // The body's REAL content box — tall bands push it past the margins, and
      // the boundary dash must sit at the body/band edge, never across band text.
      const contentTop = page.contentTopPx;
      const contentBottom = page.contentBottomPx;
      ctx.fillStyle = "rgba(255,255,255,0.55)";
      ctx.fillRect(0, contentTop, page.widthPx, contentBottom - contentTop);
      const boundaryY = bandEditMode === "header" ? contentTop : contentBottom;
      ctx.strokeStyle = theme.accent;
      ctx.setLineDash([4, 3]);
      ctx.beginPath();
      ctx.moveTo(0, boundaryY + 0.5);
      ctx.lineTo(page.widthPx, boundaryY + 0.5);
      ctx.stroke();
      ctx.setLineDash([]);
    }

    // 6. column-resize hover guide: the vertical boundary the pointer will drag,
    //    drawn as a crisp accent line with a soft grab-zone band behind it.
    if (colGuide && colGuide.pageIndex === page.index) {
      const gx = colGuide.x;
      ctx.save();
      ctx.fillStyle = theme.accent;
      ctx.globalAlpha = 0.16;
      ctx.fillRect(gx - 3, colGuide.y, 6, colGuide.height); // grab zone (matches the 6px grip)
      ctx.globalAlpha = 1;
      ctx.fillRect(gx - 0.75, colGuide.y, 1.5, colGuide.height); // the line itself
      ctx.restore();
    }

    // 6b. row-resize hover guide: the horizontal boundary the pointer will drag,
    //     drawn as a crisp accent line with a soft grab-zone band behind it.
    if (rowGuide && rowGuide.pageIndex === page.index) {
      const gy = rowGuide.y;
      ctx.save();
      ctx.fillStyle = theme.accent;
      ctx.globalAlpha = 0.16;
      ctx.fillRect(rowGuide.x, gy - 3, rowGuide.width, 6); // grab zone (matches the 6px grip)
      ctx.globalAlpha = 1;
      ctx.fillRect(rowGuide.x, gy - 0.75, rowGuide.width, 1.5); // the line itself
      ctx.restore();
    }

    // 7. develop-mode layout overlays — drawn on top of everything so the structure
    //    is visible over the text. Each kind strokes from the live layout tree.
    if (debugOverlays.size > 0) drawDebugOverlays(ctx, page);
  }

  /** Stroke the enabled layout-debug overlays for one page. */
  function drawDebugOverlays(ctx: CanvasRenderingContext2D, page: Page): void {
    const on = (k: string): boolean => debugOverlays.has(k);
    const contentRight = page.widthPx - page.marginPx.right;
    ctx.save();
    ctx.lineWidth = 1;

    // margins / content box
    if (on("margins")) {
      ctx.strokeStyle = "rgba(150,150,165,0.7)";
      ctx.setLineDash([4, 3]);
      ctx.strokeRect(page.marginPx.left + 0.5, page.marginPx.top + 0.5, contentRight - page.marginPx.left - 1, page.heightPx - page.marginPx.top - page.marginPx.bottom - 1);
      ctx.setLineDash([]);
    }

    const drawParaOverlays = (pb: PlacedBlock): void => {
      // line + fragment + baseline boxes
      for (const line of pb.lines) {
        const ly = pb.y + line.y;
        let lx = contentRight, lr = pb.x;
        for (const f of line.fragments) { lx = Math.min(lx, pb.x + f.x); lr = Math.max(lr, pb.x + f.x + f.width); }
        if (line.fragments.length === 0) { lx = pb.x; lr = pb.x + 6; }
        if (on("lineBoxes")) {
          ctx.strokeStyle = "rgba(126,198,153,0.55)";
          ctx.strokeRect(lx + 0.5, ly + 0.5, Math.max(1, lr - lx - 1), Math.max(1, line.height - 1));
        }
        if (on("fragments")) {
          ctx.strokeStyle = "rgba(215,186,125,0.6)";
          for (const f of line.fragments) ctx.strokeRect(pb.x + f.x + 0.5, ly + 0.5, Math.max(1, f.width - 1), Math.max(1, line.height - 1));
        }
        if (on("baselines")) {
          ctx.strokeStyle = "rgba(229,115,115,0.7)";
          ctx.beginPath();
          ctx.moveTo(lx, ly + line.ascent + 0.5);
          ctx.lineTo(lr, ly + line.ascent + 0.5);
          ctx.stroke();
        }
      }
    };

    for (const pb of page.blocks) {
      if (on("blockBoxes")) {
        ctx.strokeStyle = "rgba(91,155,213,0.7)";
        if (pb.table) ctx.strokeRect(pb.table.x + 0.5, pb.table.y + 0.5, pb.table.width - 1, pb.table.height - 1);
        else if (pb.image) ctx.strokeRect(pb.x + 0.5, pb.y + 0.5, pb.image.width - 1, pb.image.height - 1);
        else {
          const top = pb.y + (pb.lines[0]?.y ?? 0);
          const last = pb.lines[pb.lines.length - 1];
          const bot = last ? pb.y + last.y + last.height : top;
          ctx.strokeRect(pb.x + 0.5, top + 0.5, Math.max(1, contentRight - pb.x - 1), Math.max(1, bot - top - 1));
        }
      }
      if (pb.table) {
        if (on("cells")) {
          ctx.strokeStyle = "rgba(197,134,192,0.7)";
          for (const row of pb.table.rows) for (const cell of row.cells) ctx.strokeRect(cell.x + 0.5, cell.y + 0.5, cell.width - 1, cell.height - 1);
        }
        if (on("lineBoxes") || on("fragments") || on("baselines")) {
          for (const row of pb.table.rows) for (const cell of row.cells) for (const cb of cell.blocks) drawParaOverlays(cb);
        }
      } else {
        drawParaOverlays(pb);
      }
    }

    if (on("pageInfo")) {
      const label = `p${page.index} · #${page.number} · ${Math.round(page.widthPx)}×${Math.round(page.heightPx)}`;
      ctx.font = "10px 'Cascadia Code', Consolas, monospace";
      const w = ctx.measureText(label).width + 8;
      ctx.fillStyle = "rgba(40,42,46,0.85)";
      ctx.fillRect(2, 2, w, 14);
      ctx.fillStyle = "#cfe3f5";
      ctx.fillText(label, 6, 12);
    }
    ctx.restore();
  }

  const imageCache = new Map<string, HTMLImageElement>();

  /** Re-dirty and repaint every mounted page. Used when a bitmap arrives late —
   *  we don't track which pages use which image, and a single captured page index
   *  is wrong for shared images and stale after a remount. */
  function repaintAllLive(): void {
    for (const index of liveCanvases.keys()) dirty.add(index);
    schedule();
  }

  function getImage(src: string): HTMLImageElement {
    let img = imageCache.get(src);
    if (!img) {
      img = new Image();
      // A late bitmap (slow network, or a background-tab load where rAF was
      // throttled) must clear its gray placeholder once it decodes.
      img.addEventListener("load", repaintAllLive);
      img.src = src;
      imageCache.set(src, img);
    }
    return img;
  }

  // Paint only the table cell background fills for a block tree (recursing into
  // nested tables). Hoisted ahead of the highlight layers so opaque fills sit
  // beneath the selection/search highlights instead of burying them.
  function paintCellFills(ctx: CanvasRenderingContext2D, block: PlacedBlock, pageIndex: number): void {
    if (!block.table) return;
    for (const row of block.table.rows) {
      for (const cell of row.cells) {
        if (cell.shading) {
          ctx.fillStyle = cell.shading;
          ctx.fillRect(cell.x, cell.y, cell.width, cell.height);
        }
        for (const cb of cell.blocks) paintCellFills(ctx, cb, pageIndex);
      }
    }
  }

  // Paint paragraph background fills (w:shd) for a block tree, recursing into
  // table cells. Like paintCellFills, hoisted ahead of the highlight layers so an
  // opaque paragraph fill sits beneath the selection/search bands, not over them.
  function paintParaFills(ctx: CanvasRenderingContext2D, block: PlacedBlock, pageIndex: number): void {
    const d = block.paraDecor;
    if (d?.shading) {
      ctx.fillStyle = d.shading;
      ctx.fillRect(block.x, block.y, d.width, d.height);
    }
    if (block.table) {
      for (const row of block.table.rows) {
        for (const cell of row.cells) {
          for (const cb of cell.blocks) paintParaFills(ctx, cb, pageIndex);
        }
      }
    }
  }

  function paintBlock(ctx: CanvasRenderingContext2D, block: PlacedBlock, pageIndex: number): void {
    if (block.image) {
      const img = getImage(block.image.src);
      const clip = block.image.clip;
      if (clip) {
        ctx.save();
        ctx.beginPath();
        ctx.rect(clip.x, clip.y, clip.width, clip.height);
        ctx.clip();
      }
      if (img.complete && img.naturalWidth > 0) {
        const crop = block.image.crop;
        if (crop) {
          // a:srcRect crop: draw only the [left,1-right]×[top,1-bottom] source
          // window into the (already cropped-size) destination box.
          const sx = crop.left * img.naturalWidth;
          const sy = crop.top * img.naturalHeight;
          const sw = Math.max(1, (1 - crop.left - crop.right) * img.naturalWidth);
          const sh = Math.max(1, (1 - crop.top - crop.bottom) * img.naturalHeight);
          ctx.drawImage(img, sx, sy, sw, sh, block.x, block.y, block.image.width, block.image.height);
        } else {
          ctx.drawImage(img, block.x, block.y, block.image.width, block.image.height);
        }
      } else {
        ctx.fillStyle = theme.imagePlaceholder;
        ctx.fillRect(block.x, block.y, block.image.width, block.image.height);
      }
      if (clip) ctx.restore();
      return;
    }
    if (block.equation) {
      paintMathBoxCanvas(
        ctx,
        block.equation,
        block.x,
        block.y + block.equation.baseline,
        MATH_FONT_FAMILY,
        DEFAULT_CHAR_STYLE.color,
      );
      return;
    }
    if (block.table) {
      const rows = block.table.rows;
      // Cell fills are painted earlier (paintCellFills) so the highlight layers
      // land on top of them; here we paint 1) cell contents, 2) borders on top —
      // so a neighbour's fill never clips an already-drawn shared edge.
      for (const row of rows) {
        for (const cell of row.cells) {
          const clip = cell.contentClip;
          if (clip) {
            ctx.save();
            ctx.beginPath();
            ctx.rect(clip.x, clip.y, clip.width, clip.height);
            ctx.clip();
          }
          for (const cb of cell.blocks) paintBlock(ctx, cb, pageIndex);
          if (clip) ctx.restore();
        }
      }
      for (const row of rows) {
        for (const cell of row.cells) paintCellBorders(ctx, cell, theme.grid);
      }
      ctx.setLineDash([]);
      return;
    }
    // List marker — paint-only, on the first line's baseline in the hanging indent.
    const firstLine = block.lines[0];
    if (block.marker && firstLine) {
      ctx.font = charStyleToFont(block.marker.style);
      ctx.fillStyle = block.marker.style.color;
      (ctx as CanvasRenderingContext2D & { wordSpacing: string }).wordSpacing = "0px";
      ctx.fillText(block.marker.text, block.marker.x, block.y + firstLine.y + firstLine.ascent);
    }
    // TOC decoration — paint-only page number + dot leader on its line.
    if (block.toc) {
      const line = block.lines[block.toc.lineIndex];
      if (line) {
        const baseline = block.y + line.y + line.ascent;
        ctx.font = charStyleToFont(block.toc.style);
        // The number inherits the entry's first-run style; normalize the
        // imported Hyperlink blue so it reads as plain text like the leader.
        ctx.fillStyle = normalizeLinkBlue(block.toc.style.color);
        (ctx as CanvasRenderingContext2D & { wordSpacing: string }).wordSpacing = "0px";
        ctx.fillText(block.toc.numText, block.toc.numX, baseline);
        const lastFrag = line.fragments[line.fragments.length - 1];
        const fromX = block.x + (lastFrag ? lastFrag.x + lastFrag.width : 0) + TOC_LEADER_GAP_PX;
        const toX = block.toc.numX - TOC_LEADER_GAP_PX;
        if (toX > fromX) {
          ctx.save();
          ctx.strokeStyle = theme.tocLeader;
          ctx.lineWidth = 1;
          ctx.setLineDash(TOC_LEADER_DASH);
          ctx.beginPath();
          ctx.moveTo(fromX, baseline);
          ctx.lineTo(toX, baseline);
          ctx.stroke();
          ctx.restore();
        }
      }
    }
    for (const line of block.lines) {
      const baselineY = block.y + line.y + line.ascent;
      if (line.leaders) {
        for (const ld of line.leaders) {
          ctx.save();
          // Leaders inherit the entry's run color; a TOC entry is a hyperlink, so
          // normalize its blue to plain text (matching the de-linked entry text).
          ctx.strokeStyle = normalizeLinkBlue(ld.color);
          ctx.lineWidth = leaderWidth(ld.fontSizePx);
          ctx.setLineDash(leaderDash(ld.kind));
          ctx.beginPath();
          ctx.moveTo(block.x + ld.x1, baselineY);
          ctx.lineTo(block.x + ld.x2, baselineY);
          ctx.stroke();
          ctx.restore();
        }
      }
      for (const frag of line.fragments) {
        const s = frag.style;
        const x = block.x + frag.x;
        // sub/superscript: scaled font (already measured that way) + baseline shift
        const vShift = verticalShift(s.verticalAlign, s.fontSizePx);
        if (s.highlightColor) {
          ctx.fillStyle = s.highlightColor;
          ctx.fillRect(x, block.y + line.y, frag.width, line.height);
        }
        if (frag.equation) {
          ctx.direction = "ltr";
          ctx.textAlign = "left";
          // Honor the run's color/decorations: a hyperlinked or underlined/struck
          // inline equation paints in the link color with the rule(s) drawn over
          // its box (highlight was already filled above, with the other frags).
          const rp = runPaint(s, theme.externalLink);
          paintMathBoxCanvasRaw(ctx, frag.equation, x, baselineY, MATH_FONT_FAMILY, rp.color);
          const th = decorationThickness(s.fontSizePx);
          if (rp.underline) paintUnderline(ctx, x, baselineY + UNDERLINE_OFFSET_PX, frag.width, s.fontSizePx, rp);
          ctx.fillStyle = rp.color;
          if (rp.strike) ctx.fillRect(x, baselineY + strikeOffset(s.fontSizePx), frag.width, th);
          continue;
        }
        ctx.font = charStyleToFont(s);
        // EXTERNAL hyperlinks paint blue+underlined as an affordance. In-document
        // anchors ("#bookmark" — TOC entries, cross-references) read as normal
        // paragraph text like Word (handled by runPaint's link normalization).
        const rp = runPaint(s, theme.externalLink);
        ctx.fillStyle = rp.color;
        (ctx as CanvasRenderingContext2D & { wordSpacing: string }).wordSpacing =
          `${frag.wordSpacingPx ?? 0}px`;
        // RTL fragment: anchor at its RIGHT edge with an RTL base so the canvas
        // shapes/orders the glyphs correctly (matters for Arabic joining and for
        // neutrals at the run edges). LTR uses the default left anchor.
        const rtl = ((frag.level ?? 0) & 1) === 1;
        ctx.direction = rtl ? "rtl" : "ltr";
        ctx.textAlign = rtl ? "right" : "left";
        ctx.fillText(frag.text, rtl ? x + frag.width : x, baselineY + vShift);

        const th = decorationThickness(s.fontSizePx);
        if (rp.underline) {
          paintUnderline(ctx, x, baselineY + vShift + UNDERLINE_OFFSET_PX, frag.width, s.fontSizePx, rp);
        }
        if (rp.strike) {
          ctx.fillStyle = rp.color;
          ctx.fillRect(x, baselineY + vShift + strikeOffset(s.fontSizePx), frag.width, th);
        }
      }
      // Reset the text-direction state RTL fragments may have left, so markers,
      // the next block, and overlays draw with the default left anchor.
      ctx.direction = "ltr";
      ctx.textAlign = "left";
      if (showFormattingMarks) paintFormattingMarks(ctx, block, line, baselineY);
    }
    // Paragraph border box (w:pBdr) — drawn over the (already-filled) shading and
    // text, like a cell's edges. `between` is not drawn for a standalone paragraph.
    if (block.paraDecor?.borders) {
      const d = block.paraDecor;
      const b = d.borders!;
      const x = block.x;
      const yT = block.y;
      const xR = block.x + d.width;
      const yB = block.y + d.height;
      strokeCellEdge(ctx, b.top, x, yT, xR, yT, 0, 1);
      strokeCellEdge(ctx, b.bottom, x, yB, xR, yB, 0, -1);
      strokeCellEdge(ctx, b.left, x, yT, x, yB, 1, 0);
      strokeCellEdge(ctx, b.right, xR, yT, xR, yB, -1, 0);
      ctx.setLineDash([]);
    }
  }

  /** Non-printing marks for one line, drawn on top of its text: a center dot per
   *  space, a "→" in each tab gap, and a "¶"/"↵" at a paragraph end / soft break.
   *  Overlay only — positions come from layout (engine flags + geometry), so paint
   *  never measures anything that could affect line breaking. */
  function paintFormattingMarks(
    ctx: CanvasRenderingContext2D,
    block: PlacedBlock,
    line: LineBox,
    baselineY: number,
  ): void {
    ctx.fillStyle = theme.formattingMark;
    ctx.strokeStyle = theme.formattingMark;

    // Space dots — a small disc at the mid-height of each rendered U+0020.
    const midY = block.y + line.y + line.height / 2;
    for (const frag of line.fragments) {
      if (frag.text.indexOf(" ") < 0) continue;
      const r = Math.max(0.6, Math.min(1.4, frag.style.fontSizePx / 18));
      for (const cx of spaceMarkXs(frag)) {
        ctx.beginPath();
        ctx.arc(block.x + cx, midY, r, 0, Math.PI * 2);
        ctx.fill();
      }
    }

    // Tab arrows — a "→" spanning the gap the tab opened, on the baseline.
    if (line.tabArrows) {
      const arrowSize = Math.round(line.ascent * 0.85);
      ctx.font = `${arrowSize}px sans-serif`;
      ctx.textAlign = "center";
      for (const a of line.tabArrows) {
        ctx.fillText("→", block.x + (a.x1 + a.x2) / 2, baselineY);
      }
      ctx.textAlign = "left";
    }

    // Paragraph end (¶) or manual line break (↵), just past the last glyph.
    if (line.paragraphEnd || line.lineBreak) {
      const lastFrag = line.fragments[line.fragments.length - 1];
      const endX = lastFrag ? block.x + lastFrag.x + lastFrag.width : block.x;
      const glyphSize = Math.round(line.ascent * 0.95);
      ctx.font = `${glyphSize}px sans-serif`;
      ctx.fillText(line.paragraphEnd ? "¶" : "↵", endX + 1, baselineY);
    }
  }

  const pagesOf = (rects: Rect[]): number[] => [...new Set(rects.map((r) => r.pageIndex))];

  return {
    setTree(next: LayoutTree): void {
      // Placeholders must be rebuilt when the page COUNT or any page's
      // DIMENSIONS change (sections can resize pages without adding any).
      const dimsChanged =
        !tree ||
        tree.pages.length !== next.pages.length ||
        next.pages.some((p, i) => {
          const prev = tree!.pages[i]!;
          return prev.widthPx !== p.widthPx || prev.heightPx !== p.heightPx;
        });
      tree = next;
      if (dimsChanged) {
        pageSigs.clear(); // page set / geometry changed wholesale
        rebuildPlaceholders();
      } else {
        // Same page set & dimensions: repaint only the live pages whose paint
        // inputs actually changed (typing usually touches one page; the rest of
        // the mounted canvases are identical and can be left alone).
        for (const i of liveCanvases.keys()) {
          const sig = pageSignature(next.pages[i]!);
          if (pageSigs.get(i) !== sig) {
            dirty.add(i);
            pageSigs.set(i, sig);
          }
        }
      }
      schedule();
    },

    setBandEditMode(band: "header" | "footer" | null): void {
      if (band === bandEditMode) return;
      bandEditMode = band;
      for (const i of liveCanvases.keys()) dirty.add(i);
      schedule();
    },

    setSelectionRects(rects: Rect[]): void {
      const affected = new Set([...pagesOf(selectionRects), ...pagesOf(rects)]);
      selectionRects = rects;
      for (const i of affected) if (liveCanvases.has(i)) dirty.add(i);
      schedule();
    },

    setSearchRects(rects: Rect[]): void {
      const affected = new Set([...pagesOf(searchRects), ...pagesOf(rects)]);
      searchRects = rects;
      for (const i of affected) if (liveCanvases.has(i)) dirty.add(i);
      schedule();
    },

    setReviewDecorations(decos: ReviewDecorations): void {
      const all = (d: ReviewDecorations): number[] => [
        ...d.comments.map((r) => r.pageIndex),
        ...d.inserts.map((r) => r.pageIndex),
        ...d.deletes.map((r) => r.pageIndex),
        ...d.changeBars.map((r) => r.pageIndex),
        ...d.pins.map((p) => p.pageIndex),
      ];
      const affected = new Set([...all(reviewDecos), ...all(decos)]);
      reviewDecos = decos;
      for (const i of affected) if (liveCanvases.has(i)) dirty.add(i);
      schedule();
    },

    reviewPinAt(clientX: number, clientY: number): string | null {
      const pt = this.clientToPage(clientX, clientY);
      if (!pt) return null;
      let best: { id: string; d2: number } | null = null;
      for (const pin of reviewDecos.pins) {
        if (pin.pageIndex !== pt.pageIndex) continue;
        const dx = pt.x - pin.x;
        const dy = pt.y - (pin.y + PIN_OFFSET_PX);
        const d2 = dx * dx + dy * dy;
        if (d2 <= 100 && (!best || d2 < best.d2)) best = { id: pin.threadId, d2 }; // 10px hit radius (larger than the painted disc)
      }
      return best ? best.id : null;
    },

    setSdtAdornment(layers: { rects: Rect[]; label: string }[] | null): void {
      const flat = (ls: { rects: Rect[] }[] | null): Rect[] => (ls ?? []).flatMap((l) => l.rects);
      const affected = new Set([...pagesOf(flat(sdtAdorn)), ...pagesOf(flat(layers))]);
      sdtAdorn = layers && layers.length > 0 ? layers : null;
      for (const i of affected) if (liveCanvases.has(i)) dirty.add(i);
      schedule();
    },

    setFieldAdornment(adorn: { rects: Rect[]; label: string } | null): void {
      const affected = new Set([...pagesOf(fieldAdorn?.rects ?? []), ...pagesOf(adorn?.rects ?? [])]);
      fieldAdorn = adorn;
      for (const i of affected) if (liveCanvases.has(i)) dirty.add(i);
      schedule();
    },

    setInspectorRects(adorn: { rects: Rect[]; label: string } | null): void {
      const affected = new Set([...pagesOf(inspectorAdorn?.rects ?? []), ...pagesOf(adorn?.rects ?? [])]);
      inspectorAdorn = adorn;
      for (const i of affected) if (liveCanvases.has(i)) dirty.add(i);
      schedule();
    },

    setColumnGuide(guide: ColumnGuide | null): void {
      // Value-dedupe: hover fires every mousemove; only repaint when it moves.
      const same =
        colGuide === guide ||
        (!!colGuide && !!guide &&
          colGuide.pageIndex === guide.pageIndex && colGuide.x === guide.x &&
          colGuide.y === guide.y && colGuide.height === guide.height);
      if (same) return;
      const affected = new Set<number>();
      if (colGuide) affected.add(colGuide.pageIndex);
      if (guide) affected.add(guide.pageIndex);
      colGuide = guide;
      for (const i of affected) if (liveCanvases.has(i)) dirty.add(i);
      schedule();
    },

    setRowGuide(guide: RowGuide | null): void {
      // Value-dedupe: hover fires every mousemove; only repaint when it moves.
      const same =
        rowGuide === guide ||
        (!!rowGuide && !!guide &&
          rowGuide.pageIndex === guide.pageIndex && rowGuide.y === guide.y &&
          rowGuide.x === guide.x && rowGuide.width === guide.width);
      if (same) return;
      const affected = new Set<number>();
      if (rowGuide) affected.add(rowGuide.pageIndex);
      if (guide) affected.add(guide.pageIndex);
      rowGuide = guide;
      for (const i of affected) if (liveCanvases.has(i)) dirty.add(i);
      schedule();
    },

    setDebugOverlay(kind: string, on: boolean): void {
      if (on === debugOverlays.has(kind)) return;
      if (on) debugOverlays.add(kind);
      else debugOverlays.delete(kind);
      for (const i of liveCanvases.keys()) dirty.add(i);
      schedule();
    },

    setCaret(caret: CaretRect | null): void {
      lastCaret = caret;
      if (!caret) {
        caretEl.style.display = "none";
        return;
      }
      const ph = placeholders[caret.pageIndex];
      if (!ph) return;
      if (caretEl.parentElement !== ph) ph.appendChild(caretEl);
      caretEl.style.display = "block";
      caretEl.style.left = `${caret.x * zoom - 1}px`;
      caretEl.style.top = `${caret.y * zoom}px`;
      caretEl.style.height = `${caret.height * zoom}px`;
      // restart the blink so the caret is solid right after every move (Word behavior)
      caretEl.style.animation = "none";
      void caretEl.offsetWidth;
      caretEl.style.animation = "";
    },

    setRemoteCarets(carets: RemoteCaret[]): void {
      lastRemoteCarets = carets;
      applyRemoteCarets(carets);
    },

    invalidatePages(pageIndexes: number[]): void {
      for (const i of pageIndexes) if (liveCanvases.has(i)) dirty.add(i);
      schedule();
    },

    clientToPage(clientX: number, clientY: number): PagePoint | null {
      if (!tree || placeholders.length === 0) return null;
      const wrapRect = pagesWrap.getBoundingClientRect();
      const yIn = clientY - wrapRect.top - gap;
      // Walk cumulative offsets in DISPLAY px (heights × zoom); pages can have
      // per-section heights and the gap is unscaled (it's flex layout, not zoomed).
      let pageIndex = tree.pages.length - 1;
      let top = 0;
      for (let i = 0; i < tree.pages.length; i++) {
        const h = tree.pages[i]!.heightPx * zoom;
        if (yIn < top + h + gap / 2) {
          pageIndex = i;
          break;
        }
        top += h + gap;
      }
      const pg = tree.pages[pageIndex]!;
      const pageLeft = wrapRect.left + (wrapRect.width - pg.widthPx * zoom) / 2;
      // Return DOCUMENT coords (÷ zoom) so all hit-testing stays zoom-agnostic.
      const rawX = (clientX - pageLeft) / zoom;
      const rawY = (yIn - top) / zoom;
      return {
        pageIndex,
        x: Math.min(pg.widthPx, Math.max(0, rawX)),
        y: Math.min(pg.heightPx, Math.max(0, rawY)),
        inside: rawX >= 0 && rawX <= pg.widthPx && rawY >= 0 && rawY <= pg.heightPx,
      };
    },

    getPageElement(pageIndex: number): HTMLElement | null {
      return placeholders[pageIndex] ?? null;
    },

    caretToContainer(caret: CaretRect): { left: number; top: number } | null {
      const ph = placeholders[caret.pageIndex];
      if (!ph) return null;
      const phRect = ph.getBoundingClientRect();
      const cRect = container.getBoundingClientRect();
      return {
        left: phRect.left - cRect.left + caret.x * zoom,
        top: phRect.top - cRect.top + container.scrollTop + caret.y * zoom,
      };
    },

    ensureVisible(caret: CaretRect, align: "nearest" | "center" = "nearest"): void {
      const ph = placeholders[caret.pageIndex];
      if (!ph) return;
      const phRect = ph.getBoundingClientRect();
      const cRect = container.getBoundingClientRect();
      const caretTop = phRect.top - cRect.top + caret.y * zoom; // viewport-relative
      const caretHeight = caret.height * zoom;
      const caretBottom = caretTop + caretHeight;
      if (align === "center") {
        // Place the target's middle at the viewport's middle (clamped by the
        // container's own scroll range, so the ends don't overscroll).
        container.scrollTop += caretTop + caretHeight / 2 - cRect.height / 2;
        return;
      }
      const margin = 32;
      if (caretTop < margin) container.scrollTop += caretTop - margin;
      else if (caretBottom > cRect.height - margin) {
        container.scrollTop += caretBottom - (cRect.height - margin);
      }
    },

    setZoom(next: number): void {
      const z = Math.min(zoomMax, Math.max(zoomMin, next));
      if (z === zoom) return;
      zoom = z;
      // Resize placeholders, repaint live canvases at the new scale, reposition
      // the caret. The layout tree is untouched — zoom is pure presentation.
      for (const ph of placeholders) {
        const page = tree?.pages[Number(ph.dataset["page"])];
        if (page) {
          ph.style.width = `${page.widthPx * zoom}px`;
          ph.style.height = `${page.heightPx * zoom}px`;
        }
      }
      for (const i of liveCanvases.keys()) dirty.add(i);
      schedule();
      if (lastCaret) {
        const ph = placeholders[lastCaret.pageIndex];
        if (ph) {
          caretEl.style.left = `${lastCaret.x * zoom - 1}px`;
          caretEl.style.top = `${lastCaret.y * zoom}px`;
          caretEl.style.height = `${lastCaret.height * zoom}px`;
        }
      }
      applyRemoteCarets(lastRemoteCarets); // peers' carets follow the new zoom
    },
    getZoom(): number {
      return zoom;
    },

    setShowGrid(show: boolean): void {
      if (show === showGrid) return;
      showGrid = show;
      repaintAllLive();
    },
    getShowGrid(): boolean {
      return showGrid;
    },
    setSnapToGrid(snap: boolean): void {
      snapToGrid = snap; // drag-time only — no repaint needed
    },
    getSnapToGrid(): boolean {
      return snapToGrid;
    },
    setGridSpacing(px: number): void {
      const next = Math.max(1, Math.round(px));
      if (next === gridSpacingPx) return;
      gridSpacingPx = next;
      if (showGrid) repaintAllLive();
    },
    getGridSpacing(): number {
      return gridSpacingPx;
    },

    setShowFormattingMarks(show: boolean): void {
      if (show === showFormattingMarks) return;
      showFormattingMarks = show;
      repaintAllLive();
    },
    getShowFormattingMarks(): boolean {
      return showFormattingMarks;
    },

    destroy(): void {
      observer.disconnect();
      document.removeEventListener("visibilitychange", onVisible);
      if (rafId !== null) cancelAnimationFrame(rafId);
      pagesWrap.remove();
    },
  };
}

/** Paint a placed cell's borders. No `borders` field → the legacy uniform light
 *  grid (keeps native/unstyled tables visibly gridded). Otherwise draw exactly
 *  the edges present; an omitted edge means no line on that side. */
function paintCellBorders(ctx: CanvasRenderingContext2D, cell: PlacedTableCell, gridColor: string): void {
  const { x, y, width: w, height: h } = cell;
  if (cell.borders === undefined) {
    ctx.setLineDash([]);
    ctx.strokeStyle = gridColor;
    ctx.lineWidth = 1;
    ctx.strokeRect(x + 0.5, y + 0.5, w, h);
    return;
  }
  const b = cell.borders;
  strokeCellEdge(ctx, b.top, x, y, x + w, y, 0, 1);
  strokeCellEdge(ctx, b.bottom, x, y + h, x + w, y + h, 0, -1);
  strokeCellEdge(ctx, b.left, x, y, x, y + h, 1, 0);
  strokeCellEdge(ctx, b.right, x + w, y, x + w, y + h, -1, 0);
}

/** One cell edge from (x1,y1) to (x2,y2). (ix,iy) points into the cell interior
 *  and positions the inner stroke of a "double" border. */
function strokeCellEdge(
  ctx: CanvasRenderingContext2D,
  spec: CellBorder | undefined,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  ix: number,
  iy: number,
): void {
  if (!spec) return;
  const w = cellBorderWidth(spec.widthPx);
  ctx.strokeStyle = spec.color;
  ctx.lineWidth = w;
  ctx.setLineDash(cellBorderDash(spec.style, w));
  // Half-pixel align odd-width lines on the perpendicular axis for crisp edges.
  const horizontal = y1 === y2;
  const off = Math.round(w) % 2 ? 0.5 : 0;
  const ox = horizontal ? 0 : off;
  const oy = horizontal ? off : 0;
  const line = (dx: number, dy: number): void => {
    ctx.beginPath();
    ctx.moveTo(x1 + ox + dx, y1 + oy + dy);
    ctx.lineTo(x2 + ox + dx, y2 + oy + dy);
    ctx.stroke();
  };
  line(0, 0);
  if (spec.style === "double") {
    const g = doubleBorderGap(w);
    line(ix * g, iy * g);
  }
  ctx.setLineDash([]);
}
