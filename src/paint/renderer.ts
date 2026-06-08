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
import type { CellBorder } from "../model/document";
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
  /** Scroll the container the minimum amount to reveal the caret. */
  ensureVisible(caret: CaretRect): void;
  destroy(): void;
}

const PAGE_GAP_PX = 24;
const SELECTION_COLOR = "rgba(38, 111, 219, 0.28)";
// Office "Hyperlink" character-style blues — normalized to text color when they
// arrive on an in-document anchor (TOC/cross-ref), so those read as plain text.
const HYPERLINK_BLUES = new Set(["#0563c1", "#0000ff", "#0000ee", "#0b57d0", "#0066cc", "#1155cc"]);

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
        ctx.fillStyle = HYPERLINK_BLUES.has(block.toc.style.color.toLowerCase())
          ? "#202124"
          : block.toc.style.color;
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
        // EXTERNAL hyperlinks paint blue+underlined as an affordance. In-document
        // anchors ("#bookmark" — TOC entries, cross-references) read as normal
        // paragraph text like Word: drop the underline, and normalize the
        // imported Hyperlink-style blue to text color (other colors kept).
        const anchorLink = s.link !== undefined && s.link.startsWith("#");
        const externalLink = s.link !== undefined && !anchorLink;
        let color = s.color;
        if (externalLink) color = "#0b57d0";
        else if (anchorLink && HYPERLINK_BLUES.has(s.color.toLowerCase())) color = "#202124";
        ctx.fillStyle = color;
        (ctx as CanvasRenderingContext2D & { wordSpacing: string }).wordSpacing =
          `${frag.wordSpacingPx ?? 0}px`;
        ctx.fillText(frag.text, x, baselineY + vShift);

        if (externalLink || (s.underline && !anchorLink)) {
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

    setSdtAdornment(adorn: { rects: Rect[]; label: string } | null): void {
      const affected = new Set([...pagesOf(sdtAdorn?.rects ?? []), ...pagesOf(adorn?.rects ?? [])]);
      sdtAdorn = adorn;
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

const DEFAULT_GRID_COLOR = "#c0c4c9";

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
  const w = Math.max(0.5, spec.widthPx);
  ctx.strokeStyle = spec.color;
  ctx.lineWidth = w;
  ctx.setLineDash(
    spec.style === "dashed" ? [w * 3, w * 2] : spec.style === "dotted" ? [w, w * 1.5] : [],
  );
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
    const g = w + 1;
    line(ix * g, iy * g);
  }
  ctx.setLineDash([]);
}
