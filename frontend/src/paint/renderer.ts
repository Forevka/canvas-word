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

import type { LayoutTree, Page, PlacedBlock, PlacedTableCell } from "../layout/layoutTree";
import type { CaretRect, Rect } from "../layout/geometry";
import type { CellBorder } from "@cw/shared";
import { charStyleToFont } from "../layout/metrics";
import {
  cellBorderDash,
  cellBorderWidth,
  decorationThickness,
  DEFAULT_GRID_COLOR,
  doubleBorderGap,
  FOOTNOTE_RULE_COLOR,
  FOOTNOTE_RULE_WIDTH_FRACTION,
  IMAGE_PLACEHOLDER_COLOR,
  leaderDash,
  leaderWidth,
  normalizeLinkBlue,
  runPaint,
  strikeOffset,
  TOC_LEADER_COLOR,
  TOC_LEADER_DASH,
  TOC_LEADER_GAP_PX,
  UNDERLINE_OFFSET_PX,
  verticalShift,
} from "./paintStyle";

export interface PagePoint {
  pageIndex: number;
  x: number;
  y: number;
}

export interface PaintScheduler {
  setTree(tree: LayoutTree): void;
  /** Story-edit affordance: dim the body and mark the band boundary (Word's
   *  header/footer mode). Null restores normal painting. */
  setBandEditMode(band: "header" | "footer" | null): void;
  /** Precomputed highlight rects (geometry.selectionRects). Repaints affected pages. */
  setSelectionRects(rects: Rect[]): void;
  /** Find-bar match highlights — painted under the selection. */
  setSearchRects(rects: Rect[]): void;
  /** Content-control focus adornment: bounding boxes + a title tab (Word's
   *  gray frame around the active control). Null clears. */
  setSdtAdornment(adorn: { rects: Rect[]; label: string } | null): void;
  setCaret(caret: CaretRect | null): void;
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
  destroy(): void;
}

const PAGE_GAP_PX = 24;
const SELECTION_COLOR = "rgba(38, 111, 219, 0.28)";

let caretCssInjected = false;
function injectCaretCss(): void {
  if (caretCssInjected) return;
  caretCssInjected = true;
  const style = document.createElement("style");
  style.textContent =
    "@keyframes cw-caret-blink{0%,55%{opacity:1}56%,100%{opacity:0}}" +
    ".cw-caret{position:absolute;width:2px;background:#1a1a2e;pointer-events:none;animation:cw-caret-blink 1.06s step-end infinite;}";
  document.head.appendChild(style);
}

export function createPaintLayer(container: HTMLElement): PaintScheduler {
  injectCaretCss();
  let tree: LayoutTree | null = null;
  let selectionRects: Rect[] = [];
  let searchRects: Rect[] = [];
  let sdtAdorn: { rects: Rect[]; label: string } | null = null;
  let bandEditMode: "header" | "footer" | null = null;

  const pagesWrap = document.createElement("div");
  pagesWrap.style.cssText = `display:flex;flex-direction:column;align-items:center;gap:${PAGE_GAP_PX}px;padding:${PAGE_GAP_PX}px 0;`;
  container.appendChild(pagesWrap);

  const caretEl = document.createElement("div");
  caretEl.className = "cw-caret";
  caretEl.style.display = "none";

  const placeholders: HTMLDivElement[] = [];
  const liveCanvases = new Map<number, HTMLCanvasElement>();
  const dirty = new Set<number>();
  let rafId: number | null = null;
  let zoom = 1;
  let lastCaret: CaretRect | null = null;

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

  function rebuildPlaceholders(): void {
    observer.disconnect();
    liveCanvases.clear();
    caretEl.remove();
    pagesWrap.textContent = "";
    placeholders.length = 0;
    if (!tree) return;
    for (const page of tree.pages) {
      const ph = document.createElement("div");
      ph.dataset["page"] = String(page.index);
      ph.style.cssText = `position:relative;width:${page.widthPx * zoom}px;height:${page.heightPx * zoom}px;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.25);flex-shrink:0;`;
      pagesWrap.appendChild(ph);
      placeholders.push(ph);
      observer.observe(ph);
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
    for (const index of dirty) {
      const canvas = liveCanvases.get(index);
      const page = tree.pages[index];
      if (canvas && page) paintPage(canvas, page);
    }
    dirty.clear();
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

    // 1. page background
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, page.widthPx, page.heightPx);

    // 2a. search-match highlights (under everything)
    ctx.fillStyle = "rgba(251, 188, 4, 0.45)";
    for (const r of searchRects) {
      if (r.pageIndex === page.index) ctx.fillRect(r.x, r.y, r.width, r.height);
    }
    // 2b. selection highlights (under text)
    ctx.fillStyle = SELECTION_COLOR;
    for (const r of selectionRects) {
      if (r.pageIndex === page.index) ctx.fillRect(r.x, r.y, r.width, r.height);
    }

    // 2c. content-control adornment: gray frame + title tab (Word's active
    // control chrome). The tab text is UI chrome, not document text — measuring
    // it here doesn't violate the paint-never-measures-layout invariant.
    if (sdtAdorn) {
      const onPage = sdtAdorn.rects.filter((r) => r.pageIndex === page.index);
      ctx.strokeStyle = "#a8a8a8";
      ctx.lineWidth = 1;
      for (const r of onPage) {
        ctx.strokeRect(r.x - 2.5, r.y - 1.5, r.width + 5, r.height + 3);
      }
      const first = onPage[0];
      if (first && sdtAdorn.label) {
        ctx.font = "10px Arial";
        const w = ctx.measureText(sdtAdorn.label).width + 10;
        const tabH = 15;
        const tx = first.x - 2.5;
        const ty = first.y - 1.5 - tabH;
        ctx.fillStyle = "#d8d8d8";
        ctx.beginPath();
        ctx.roundRect(tx, ty, w, tabH, [3, 3, 0, 0]);
        ctx.fill();
        ctx.fillStyle = "#3c4043";
        ctx.fillText(sdtAdorn.label, tx + 5, ty + 11);
      }
    }

    // 3. blocks: text fragments (one fillText each), images, table grids
    ctx.textBaseline = "alphabetic";
    for (const block of page.blocks) paintBlock(ctx, block, page.index);

    // 3b. footnote separator rule (1/3 content width, Word style)
    if (page.footnoteRuleY !== undefined) {
      const cw = page.widthPx - page.marginPx.left - page.marginPx.right;
      ctx.strokeStyle = FOOTNOTE_RULE_COLOR;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(page.marginPx.left, page.footnoteRuleY + 0.5);
      ctx.lineTo(page.marginPx.left + cw * FOOTNOTE_RULE_WIDTH_FRACTION, page.footnoteRuleY + 0.5);
      ctx.stroke();
    }

    // 4. margin-band stories (headers/footers) — pre-laid-out rich blocks,
    //    painted through the same block painter, outside the selectable tree.
    if (page.header) for (const b of page.header) paintBlock(ctx, b, page.index);
    if (page.footer) for (const b of page.footer) paintBlock(ctx, b, page.index);

    // 5. story-edit affordance: dim the body, dash the band boundary.
    if (bandEditMode) {
      // The body's REAL content box — tall bands push it past the margins, and
      // the boundary dash must sit at the body/band edge, never across band text.
      const contentTop = page.contentTopPx;
      const contentBottom = page.contentBottomPx;
      ctx.fillStyle = "rgba(255,255,255,0.55)";
      ctx.fillRect(0, contentTop, page.widthPx, contentBottom - contentTop);
      const boundaryY = bandEditMode === "header" ? contentTop : contentBottom;
      ctx.strokeStyle = "#1a73e8";
      ctx.setLineDash([4, 3]);
      ctx.beginPath();
      ctx.moveTo(0, boundaryY + 0.5);
      ctx.lineTo(page.widthPx, boundaryY + 0.5);
      ctx.stroke();
      ctx.setLineDash([]);
    }
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
        ctx.drawImage(img, block.x, block.y, block.image.width, block.image.height);
      } else {
        ctx.fillStyle = IMAGE_PLACEHOLDER_COLOR;
        ctx.fillRect(block.x, block.y, block.image.width, block.image.height);
      }
      if (clip) ctx.restore();
      return;
    }
    if (block.table) {
      const rows = block.table.rows;
      // 1) cell fills (under everything), 2) cell contents, 3) borders on top —
      // so a neighbour's fill never clips an already-drawn shared edge.
      for (const row of rows) {
        for (const cell of row.cells) {
          if (cell.shading) {
            ctx.fillStyle = cell.shading;
            ctx.fillRect(cell.x, cell.y, cell.width, cell.height);
          }
        }
      }
      for (const row of rows) {
        for (const cell of row.cells) {
          for (const cb of cell.blocks) paintBlock(ctx, cb, pageIndex);
        }
      }
      for (const row of rows) {
        for (const cell of row.cells) paintCellBorders(ctx, cell);
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
          ctx.strokeStyle = TOC_LEADER_COLOR;
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
        ctx.font = charStyleToFont(s);
        // EXTERNAL hyperlinks paint blue+underlined as an affordance. In-document
        // anchors ("#bookmark" — TOC entries, cross-references) read as normal
        // paragraph text like Word (handled by runPaint's link normalization).
        const rp = runPaint(s);
        ctx.fillStyle = rp.color;
        (ctx as CanvasRenderingContext2D & { wordSpacing: string }).wordSpacing =
          `${frag.wordSpacingPx ?? 0}px`;
        ctx.fillText(frag.text, x, baselineY + vShift);

        const th = decorationThickness(s.fontSizePx);
        if (rp.underline) {
          ctx.fillRect(x, baselineY + vShift + UNDERLINE_OFFSET_PX, frag.width, th);
        }
        if (rp.strike) {
          ctx.fillRect(x, baselineY + vShift + strikeOffset(s.fontSizePx), frag.width, th);
        }
      }
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
      if (dimsChanged) rebuildPlaceholders();
      else for (const i of liveCanvases.keys()) dirty.add(i);
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

    setSdtAdornment(adorn: { rects: Rect[]; label: string } | null): void {
      const affected = new Set([...pagesOf(sdtAdorn?.rects ?? []), ...pagesOf(adorn?.rects ?? [])]);
      sdtAdorn = adorn;
      for (const i of affected) if (liveCanvases.has(i)) dirty.add(i);
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

    invalidatePages(pageIndexes: number[]): void {
      for (const i of pageIndexes) if (liveCanvases.has(i)) dirty.add(i);
      schedule();
    },

    clientToPage(clientX: number, clientY: number): PagePoint | null {
      if (!tree || placeholders.length === 0) return null;
      const wrapRect = pagesWrap.getBoundingClientRect();
      const yIn = clientY - wrapRect.top - PAGE_GAP_PX;
      // Walk cumulative offsets in DISPLAY px (heights × zoom); pages can have
      // per-section heights and the gap is unscaled (it's flex layout, not zoomed).
      let pageIndex = tree.pages.length - 1;
      let top = 0;
      for (let i = 0; i < tree.pages.length; i++) {
        const h = tree.pages[i]!.heightPx * zoom;
        if (yIn < top + h + PAGE_GAP_PX / 2) {
          pageIndex = i;
          break;
        }
        top += h + PAGE_GAP_PX;
      }
      const pg = tree.pages[pageIndex]!;
      const pageLeft = wrapRect.left + (wrapRect.width - pg.widthPx * zoom) / 2;
      // Return DOCUMENT coords (÷ zoom) so all hit-testing stays zoom-agnostic.
      return {
        pageIndex,
        x: Math.min(pg.widthPx, Math.max(0, (clientX - pageLeft) / zoom)),
        y: Math.min(pg.heightPx, Math.max(0, (yIn - top) / zoom)),
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
      const z = Math.min(5, Math.max(0.25, next));
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
    },
    getZoom(): number {
      return zoom;
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
function paintCellBorders(ctx: CanvasRenderingContext2D, cell: PlacedTableCell): void {
  const { x, y, width: w, height: h } = cell;
  if (cell.borders === undefined) {
    ctx.setLineDash([]);
    ctx.strokeStyle = DEFAULT_GRID_COLOR;
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
