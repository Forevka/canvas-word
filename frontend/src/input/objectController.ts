// Object manipulation: image selection frames with 8 resize handles, and table
// column-boundary drags. Lives in the DOM overlay layer (like the caret) — the
// canvas never repaints for hover/handle state.
//
// Resize protocol (both images and columns): live preview during the drag via
// TRANSIENT ops (outside the undo stack), then on mouseup the wiring layer
// reverts the preview and commits ONE undoable op — same pattern as IME
// composition, so a whole drag is a single undo step.

import type { Rect } from "../layout/geometry";

export interface ObjectFrameDeps {
  getPageElement(pageIndex: number): HTMLElement | null;
  /** Current presentational zoom — the frame lives in zoomed page pixels, but
   *  sizes are document px, so positions scale up and drag deltas scale down. */
  getZoom(): number;
  /** Final size on mouseup (one undoable op). The whole drag is previewed purely
   *  in the DOM overlay (no model ops, no relayout) — this is the ONLY mutation. */
  onResizeCommit(width: number, height: number): void;
}

export interface ObjectFrame {
  /** Show the frame over a rect (page coords) for an image of natural ratio.
   *  `src` is the image URL used to paint the live resize ghost during a drag.
   *  `anchor` is the horizontal edge the layout keeps fixed as the width changes
   *  (in-flow images re-align on resize): "left" for left-aligned / anchored
   *  images, "center" / "right" for those alignments — so the ghost tracks the
   *  committed position instead of jumping on mouseup. */
  show(rect: Rect, maxWidth: number, src?: string, anchor?: ResizeAnchor, resizable?: boolean): void;
  hide(): void;
  destroy(): void;
}

/** Horizontal edge held fixed while an in-flow image resizes (mirrors the layout
 *  engine: x = colX + slack·alignFactor, slack = colWidth − width). */
export type ResizeAnchor = "left" | "center" | "right";

interface HandleSpec {
  name: string;
  dx: 0 | 0.5 | 1;
  dy: 0 | 0.5 | 1;
  cursor: string;
}

const HANDLES: HandleSpec[] = [
  { name: "nw", dx: 0, dy: 0, cursor: "nwse-resize" },
  { name: "n", dx: 0.5, dy: 0, cursor: "ns-resize" },
  { name: "ne", dx: 1, dy: 0, cursor: "nesw-resize" },
  { name: "e", dx: 1, dy: 0.5, cursor: "ew-resize" },
  { name: "se", dx: 1, dy: 1, cursor: "nwse-resize" },
  { name: "s", dx: 0.5, dy: 1, cursor: "ns-resize" },
  { name: "sw", dx: 0, dy: 1, cursor: "nesw-resize" },
  { name: "w", dx: 0, dy: 0.5, cursor: "ew-resize" },
];

const MIN_SIZE_PX = 16;

export function createObjectFrame(deps: ObjectFrameDeps): ObjectFrame {
  const frame = document.createElement("div");
  frame.style.cssText =
    "position:absolute;display:none;border:2px solid #1a73e8;box-sizing:border-box;" +
    "pointer-events:none;z-index:2;";

  // Live resize ghost: a bitmap of the image painted at the in-progress size over
  // an opaque page-colored backdrop covering the ORIGINAL footprint. This lets a
  // handle drag preview the new size purely in the DOM — no model op, no relayout
  // (a full relayout on this document costs ~33ms; doing it per drag frame was the
  // lag). The model is mutated once, on mouseup. The backdrop hides the
  // still-old-size bitmap the canvas painted underneath when shrinking.
  const ghost = document.createElement("div");
  ghost.style.cssText =
    "position:absolute;display:none;pointer-events:none;z-index:1;" +
    "background-repeat:no-repeat;background-position:top left;";
  let ghostSrc: string | undefined;

  // Original-size silhouette: a dashed outline pinned to the image's footprint at
  // drag start, so the user sees the "before" size next to the live "after" ghost
  // and frame. Sits above the ghost backdrop, below the blue frame.
  const sizeGhost = document.createElement("div");
  sizeGhost.style.cssText =
    "position:absolute;display:none;pointer-events:none;z-index:1;box-sizing:border-box;" +
    "border:1px dashed #5f6368;background:rgba(95,99,104,0.06);";

  const handleEls: HTMLDivElement[] = HANDLES.map((h) => {
    const el = document.createElement("div");
    el.dataset["handle"] = h.name;
    el.className = "cw-obj-handle"; // the mobile CSS enlarges its touch hit area
    el.style.cssText =
      "position:absolute;width:8px;height:8px;background:#fff;border:1.5px solid #1a73e8;" +
      `border-radius:50%;pointer-events:auto;cursor:${h.cursor};` +
      `left:calc(${h.dx * 100}% - 4px);top:calc(${h.dy * 100}% - 4px);`;
    frame.appendChild(el);
    return el;
  });

  let current: { rect: Rect; maxWidth: number; anchor: ResizeAnchor } | null = null;
  let drag: {
    spec: HandleSpec;
    startX: number;
    startY: number;
    startW: number;
    startH: number;
    /** Horizontal edge the committed layout keeps fixed as the width changes. */
    anchor: ResizeAnchor;
    /** Image footprint at drag start (page coords), in document px. The ghost is
     *  anchored here — in-flow images keep their layout position while resizing,
     *  so only width/height change during the preview. */
    startRect: Rect;
  } | null = null;

  // Pointer Events (not mouse) so the handles drag by finger/pen too. Pointer
  // capture routes move/up to the captured handle even when the finger leaves it.
  let dragEl: HTMLElement | null = null;
  let dragPointerId = 0;

  // The drag preview is painted into the ghost overlay (DOM only), coalesced to
  // one paint per animation frame. pointermove fires faster than the browser
  // composites; collapsing to the latest size per frame keeps the ghost glued to
  // the handle with zero model work. Same per-frame-throttle pattern as pinch-zoom.
  let movePending: { w: number; h: number } | null = null;
  let moveRaf = 0;
  const flushMove = (): void => {
    moveRaf = 0;
    if (movePending && drag) paintPreview(movePending.w, movePending.h);
    movePending = null;
  };
  const cancelPendingMove = (): void => {
    if (moveRaf) cancelAnimationFrame(moveRaf);
    moveRaf = 0;
    movePending = null;
  };

  /** Paint the in-progress resize into the ghost + frame, purely in the DOM. The
   *  ghost shows the bitmap at the new size over an opaque page-colored backdrop
   *  that covers the original footprint (so shrinking doesn't reveal the
   *  still-old-size bitmap the canvas painted underneath). */
  const paintPreview = (w: number, h: number): void => {
    if (!drag) return;
    const r = drag.startRect;
    const z = deps.getZoom();
    // Hold the same horizontal edge the committed layout will: left-aligned and
    // anchored images keep their left edge; centered images keep their center;
    // right-aligned keep their right edge. Without this the ghost grows from the
    // left and the image jumps sideways on mouseup once layout re-aligns it.
    const x = drag.anchor === "center" ? r.x + (r.width - w) / 2 : drag.anchor === "right" ? r.x + (r.width - w) : r.x;
    // Backdrop covers the union of old and new footprints; the bitmap is drawn at
    // the new size, top-left, within it.
    const coverLeft = Math.min(r.x, x);
    const coverW = (Math.max(r.x + r.width, x + w) - coverLeft) * z;
    const coverH = Math.max(r.height, h) * z;
    ghost.style.left = `${coverLeft * z}px`;
    ghost.style.top = `${r.y * z}px`;
    ghost.style.width = `${coverW}px`;
    ghost.style.height = `${coverH}px`;
    ghost.style.backgroundPosition = `${(x - coverLeft) * z}px 0`;
    ghost.style.backgroundSize = `${w * z}px ${h * z}px`;
    ghost.style.display = "block";
    // Frame border + handles track the new size at the anchored position.
    frame.style.left = `${x * z - 2}px`;
    frame.style.top = `${r.y * z - 2}px`;
    frame.style.width = `${w * z + 4}px`;
    frame.style.height = `${h * z + 4}px`;
  };

  const hideGhost = (): void => {
    ghost.style.display = "none";
    ghost.style.backgroundImage = "";
    sizeGhost.style.display = "none";
  };

  const endDrag = (): void => {
    if (!dragEl) return;
    dragEl.removeEventListener("pointermove", onDragMove);
    dragEl.removeEventListener("pointerup", onDragUp);
    dragEl.removeEventListener("pointercancel", onDragUp);
    try {
      dragEl.releasePointerCapture(dragPointerId);
    } catch {
      /* pointer already released */
    }
    dragEl = null;
  };

  const onHandleDown = (ev: PointerEvent): void => {
    if (!current) return;
    const el = ev.currentTarget as HTMLElement;
    const spec = HANDLES.find((h) => h.name === el.dataset["handle"]);
    if (!spec) return;
    ev.preventDefault();
    ev.stopPropagation();
    drag = {
      spec,
      startX: ev.clientX,
      startY: ev.clientY,
      startW: current.rect.width,
      startH: current.rect.height,
      anchor: current.anchor,
      startRect: current.rect,
    };
    // Arm the ghost with the image bitmap and a backdrop matching the page color,
    // so a shrink hides the old-size bitmap painted on the canvas underneath.
    const host = frame.parentElement;
    if (ghostSrc && host) {
      if (ghost.parentElement !== host) host.appendChild(ghost);
      const pageBg = getComputedStyle(host).backgroundColor;
      ghost.style.backgroundColor = pageBg && pageBg !== "rgba(0, 0, 0, 0)" ? pageBg : "#fff";
      ghost.style.backgroundImage = `url("${ghostSrc}")`;
    }
    // Pin the original-size silhouette to the starting footprint (it stays put for
    // the whole drag) so the user can compare the old size against the live ghost.
    if (host) {
      if (sizeGhost.parentElement !== host) host.appendChild(sizeGhost);
      const z = deps.getZoom();
      const r = drag.startRect;
      sizeGhost.style.left = `${r.x * z}px`;
      sizeGhost.style.top = `${r.y * z}px`;
      sizeGhost.style.width = `${r.width * z}px`;
      sizeGhost.style.height = `${r.height * z}px`;
      sizeGhost.style.display = "block";
    }
    dragEl = el;
    dragPointerId = ev.pointerId;
    el.setPointerCapture(ev.pointerId);
    el.addEventListener("pointermove", onDragMove);
    el.addEventListener("pointerup", onDragUp);
    el.addEventListener("pointercancel", onDragUp);
  };

  const sizeFromDrag = (ev: MouseEvent): { w: number; h: number } => {
    const d = drag!;
    const c = current!;
    // Horizontal handles pull away from the opposite edge; middle handles don't move that axis.
    const sx = d.spec.dx === 0 ? -1 : d.spec.dx === 1 ? 1 : 0;
    const sy = d.spec.dy === 0 ? -1 : d.spec.dy === 1 ? 1 : 0;
    // Mouse deltas are screen px; convert to document px (÷ zoom) before applying.
    const z = deps.getZoom();
    let w = d.startW + (sx * (ev.clientX - d.startX)) / z;
    let h = d.startH + (sy * (ev.clientY - d.startY)) / z;
    const corner = sx !== 0 && sy !== 0;
    if (corner) {
      // Word behavior: corner handles preserve aspect ratio.
      const ratio = d.startW / Math.max(1, d.startH);
      const byW = Math.abs(w - d.startW) >= Math.abs(h - d.startH) * ratio;
      if (byW) h = w / ratio;
      else w = h * ratio;
    }
    if (sx === 0) w = d.startW;
    if (sy === 0) h = d.startH;
    w = Math.min(c.maxWidth, Math.max(MIN_SIZE_PX, w));
    if (corner) h = (w / Math.max(1, d.startW)) * d.startH; // re-clamp keeps ratio
    h = Math.max(MIN_SIZE_PX, h);
    return { w: Math.round(w), h: Math.round(h) };
  };

  const onDragMove = (ev: PointerEvent): void => {
    if (!drag) return;
    movePending = sizeFromDrag(ev);
    if (!moveRaf) moveRaf = requestAnimationFrame(flushMove);
  };

  const onDragUp = (ev: PointerEvent): void => {
    if (!drag) return;
    const { w, h } = sizeFromDrag(ev);
    cancelPendingMove(); // drop any queued preview; the commit below is the final size
    drag = null;
    endDrag();
    hideGhost(); // the committed relayout repaints the image at the new size
    deps.onResizeCommit(w, h);
  };

  for (const el of handleEls) el.addEventListener("pointerdown", onHandleDown);

  return {
    show(rect: Rect, maxWidth: number, src?: string, anchor: ResizeAnchor = "left", resizable = true): void {
      const host = deps.getPageElement(rect.pageIndex);
      if (!host) return;
      current = { rect, maxWidth, anchor };
      ghostSrc = src;
      // Non-resizable objects (equations) get a plain selection box — no handles.
      for (const el of handleEls) el.style.display = resizable ? "block" : "none";
      if (frame.parentElement !== host) {
        host.appendChild(frame);
        // Re-parenting the frame detaches it from the document for an instant.
        // If a resize is in flight (e.g. a reflow moved the image onto another
        // page mid-drag), that detach implicitly releases the handle's pointer
        // capture, so re-acquire it — otherwise the drag silently dies and the
        // user has to grab the handle again to keep resizing.
        if (drag && dragEl) {
          try {
            dragEl.setPointerCapture(dragPointerId);
          } catch {
            /* pointer no longer active */
          }
        }
      }
      const z = deps.getZoom();
      frame.style.display = "block";
      frame.style.left = `${rect.x * z - 2}px`;
      frame.style.top = `${rect.y * z - 2}px`;
      frame.style.width = `${rect.width * z + 4}px`;
      frame.style.height = `${rect.height * z + 4}px`;
    },
    hide(): void {
      current = null;
      ghostSrc = undefined;
      frame.style.display = "none";
      hideGhost();
    },
    destroy(): void {
      cancelPendingMove(); // drop any in-flight resize frame so it can't fire post-teardown
      hideGhost();
      ghost.remove();
      sizeGhost.remove();
      frame.remove();
    },
  };
}
