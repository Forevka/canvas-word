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

import type { LayoutTree, Page, PlacedBlock } from "../layout/layoutTree";
import type { CaretRect, Rect } from "../layout/geometry";
import { charStyleToFont } from "../layout/metrics";

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
  /** Scroll the container the minimum amount to reveal the caret. */
  ensureVisible(caret: CaretRect): void;
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
      ph.style.cssText = `position:relative;width:${page.widthPx}px;height:${page.heightPx}px;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.25);flex-shrink:0;`;
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
    const dpr = window.devicePixelRatio || 1;
    const w = Math.round(page.widthPx * dpr);
    const h = Math.round(page.heightPx * dpr);
    if (canvas.width !== w || canvas.height !== h) {
      canvas.width = w;
      canvas.height = h;
    }
    const ctx = canvas.getContext("2d")!;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

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

    // 3. blocks: text fragments (one fillText each), images, table grids
    ctx.textBaseline = "alphabetic";
    for (const block of page.blocks) paintBlock(ctx, block, page.index);

    // 3b. footnote separator rule (1/3 content width, Word style)
    if (page.footnoteRuleY !== undefined) {
      const cw = page.widthPx - page.marginPx.left - page.marginPx.right;
      ctx.strokeStyle = "#80868b";
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(page.marginPx.left, page.footnoteRuleY + 0.5);
      ctx.lineTo(page.marginPx.left + cw / 3, page.footnoteRuleY + 0.5);
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

  function getImage(src: string, pageIndex: number): HTMLImageElement {
    let img = imageCache.get(src);
    if (!img) {
      img = new Image();
      img.src = src;
      img.addEventListener("load", () => {
        // repaint the page once the bitmap arrives
        if (liveCanvases.has(pageIndex)) dirty.add(pageIndex);
        schedule();
      });
      imageCache.set(src, img);
    }
    return img;
  }

  function paintBlock(ctx: CanvasRenderingContext2D, block: PlacedBlock, pageIndex: number): void {
    if (block.image) {
      const img = getImage(block.image.src, pageIndex);
      if (img.complete && img.naturalWidth > 0) {
        ctx.drawImage(img, block.x, block.y, block.image.width, block.image.height);
      } else {
        ctx.fillStyle = "#f1f3f4";
        ctx.fillRect(block.x, block.y, block.image.width, block.image.height);
      }
      return;
    }
    if (block.table) {
      ctx.strokeStyle = "#c0c4c9";
      ctx.lineWidth = 1;
      for (const row of block.table.rows) {
        for (const cell of row.cells) {
          ctx.strokeRect(cell.x + 0.5, cell.y + 0.5, cell.width, cell.height);
          for (const cb of cell.blocks) paintBlock(ctx, cb, pageIndex);
        }
      }
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
        ctx.fillStyle = block.toc.style.color;
        (ctx as CanvasRenderingContext2D & { wordSpacing: string }).wordSpacing = "0px";
        ctx.fillText(block.toc.numText, block.toc.numX, baseline);
        const lastFrag = line.fragments[line.fragments.length - 1];
        const fromX = block.x + (lastFrag ? lastFrag.x + lastFrag.width : 0) + 8;
        const toX = block.toc.numX - 8;
        if (toX > fromX) {
          ctx.save();
          ctx.strokeStyle = "#9aa0a6";
          ctx.lineWidth = 1;
          ctx.setLineDash([1, 4]);
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
      for (const frag of line.fragments) {
        const s = frag.style;
        const x = block.x + frag.x;
        // sub/superscript: scaled font (already measured that way) + baseline shift
        const vShift =
          s.verticalAlign === "super" ? -0.38 * s.fontSizePx
          : s.verticalAlign === "sub" ? 0.16 * s.fontSizePx
          : 0;
        if (s.highlightColor) {
          ctx.fillStyle = s.highlightColor;
          ctx.fillRect(x, block.y + line.y, frag.width, line.height);
        }
        ctx.font = charStyleToFont(s);
        // hyperlinks paint blue+underlined at render time (model keeps the
        // user's own color so removing the link restores it)
        ctx.fillStyle = s.link ? "#0b57d0" : s.color;
        (ctx as CanvasRenderingContext2D & { wordSpacing: string }).wordSpacing =
          `${frag.wordSpacingPx ?? 0}px`;
        ctx.fillText(frag.text, x, baselineY + vShift);

        if (s.underline || s.link) {
          ctx.fillRect(x, baselineY + vShift + 1.5, frag.width, Math.max(1, s.fontSizePx / 14));
        }
        if (s.strikethrough) {
          const mid = baselineY + vShift - s.fontSizePx * 0.28;
          ctx.fillRect(x, mid, frag.width, Math.max(1, s.fontSizePx / 14));
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

    setCaret(caret: CaretRect | null): void {
      if (!caret) {
        caretEl.style.display = "none";
        return;
      }
      const ph = placeholders[caret.pageIndex];
      if (!ph) return;
      if (caretEl.parentElement !== ph) ph.appendChild(caretEl);
      caretEl.style.display = "block";
      caretEl.style.left = `${caret.x - 1}px`;
      caretEl.style.top = `${caret.y}px`;
      caretEl.style.height = `${caret.height}px`;
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
      // Pages can have per-section heights — walk the cumulative offsets.
      let pageIndex = tree.pages.length - 1;
      let top = 0;
      for (let i = 0; i < tree.pages.length; i++) {
        const h = tree.pages[i]!.heightPx;
        if (yIn < top + h + PAGE_GAP_PX / 2) {
          pageIndex = i;
          break;
        }
        top += h + PAGE_GAP_PX;
      }
      const pg = tree.pages[pageIndex]!;
      const pageLeft = wrapRect.left + (wrapRect.width - pg.widthPx) / 2;
      return {
        pageIndex,
        x: Math.min(pg.widthPx, Math.max(0, clientX - pageLeft)),
        y: Math.min(pg.heightPx, Math.max(0, yIn - top)),
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
        left: phRect.left - cRect.left + caret.x,
        top: phRect.top - cRect.top + container.scrollTop + caret.y,
      };
    },

    ensureVisible(caret: CaretRect): void {
      const ph = placeholders[caret.pageIndex];
      if (!ph) return;
      const phRect = ph.getBoundingClientRect();
      const cRect = container.getBoundingClientRect();
      const caretTop = phRect.top - cRect.top + caret.y; // viewport-relative
      const caretBottom = caretTop + caret.height;
      const margin = 32;
      if (caretTop < margin) container.scrollTop += caretTop - margin;
      else if (caretBottom > cRect.height - margin) {
        container.scrollTop += caretBottom - (cRect.height - margin);
      }
    },

    destroy(): void {
      observer.disconnect();
      if (rafId !== null) cancelAnimationFrame(rafId);
      pagesWrap.remove();
    },
  };
}
