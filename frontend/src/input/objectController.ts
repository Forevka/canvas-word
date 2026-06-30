// Object manipulation: image selection frames with 8 resize handles, and table
// column-boundary drags. Lives in the DOM overlay layer (like the caret) — the
// canvas never repaints for hover/handle state.
//
// Resize protocol (both images and columns): live preview during the drag via
// TRANSIENT ops (outside the undo stack), then on mouseup the wiring layer
// reverts the preview and commits ONE undoable op — same pattern as IME
// composition, so a whole drag is a single undo step.

import type { Rect } from "../layout/geometry";

/** Image crop insets — 0..1 fractions trimmed off each edge of the source (OOXML
 *  a:srcRect). The visible source window is [left,1-right] × [top,1-bottom]. */
export interface CropInsets {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

const NO_CROP: CropInsets = { left: 0, top: 0, right: 0, bottom: 0 };

/** The full source-image footprint (document px) such that the CURRENT crop window
 *  maps exactly onto the displayed box `rect`. The displayed box shows the
 *  [left,1-right]×[top,1-bottom] window of this rect, so the rect is generally
 *  larger than the box and bleeds past it on the cropped edges. */
export function fullSourceRect(rect: Rect, crop: CropInsets): Rect {
  const wx = Math.max(0.01, 1 - crop.left - crop.right);
  const wy = Math.max(0.01, 1 - crop.top - crop.bottom);
  const width = rect.width / wx;
  const height = rect.height / wy;
  return { x: rect.x - crop.left * width, y: rect.y - crop.top * height, width, height, pageIndex: rect.pageIndex };
}

/** The bright crop window (document px) for a set of insets within the full
 *  source rect — inverse of {@link cropInsetsFromWindow}. */
export function windowRectFromInsets(full: Rect, crop: CropInsets): Rect {
  return {
    x: full.x + crop.left * full.width,
    y: full.y + crop.top * full.height,
    width: Math.max(1, (1 - crop.left - crop.right) * full.width),
    height: Math.max(1, (1 - crop.top - crop.bottom) * full.height),
    pageIndex: full.pageIndex,
  };
}

/** Crop insets for a window rect positioned within the full source rect — the
 *  fraction of each edge trimmed away. Clamped to [0,1). */
export function cropInsetsFromWindow(full: Rect, win: Rect): CropInsets {
  const clamp = (v: number): number => Math.min(0.999, Math.max(0, v));
  return {
    left: clamp((win.x - full.x) / full.width),
    top: clamp((win.y - full.y) / full.height),
    right: clamp(1 - (win.x + win.width - full.x) / full.width),
    bottom: clamp(1 - (win.y + win.height - full.y) / full.height),
  };
}

/** A crop is "empty" (no trim) when every inset rounds to zero — used to clear the
 *  field rather than write a degenerate full-frame crop. */
export function isEmptyCrop(crop: CropInsets): boolean {
  return crop.left < 1e-3 && crop.top < 1e-3 && crop.right < 1e-3 && crop.bottom < 1e-3;
}

export interface ObjectFrameDeps {
  getPageElement(pageIndex: number): HTMLElement | null;
  /** Current presentational zoom — the frame lives in zoomed page pixels, but
   *  sizes are document px, so positions scale up and drag deltas scale down. */
  getZoom(): number;
  /** Final size on mouseup (one undoable op). The whole drag is previewed purely
   *  in the DOM overlay (no model ops, no relayout) — this is the ONLY mutation. */
  onResizeCommit(width: number, height: number): void;
  /** Live crop preview: fired when a crop handle is released, with the new insets.
   *  The wiring writes them as a TRANSIENT op (so the model tracks the crop live)
   *  and reverts + commits once on exit — the drag-to-resize-row-height protocol. */
  onCropPreview(crop: CropInsets): void;
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
  /** Enter crop mode for the image at `rect` (its displayed box) with the given
   *  current crop. Shows the faded full source behind a bright draggable window
   *  with 8 crop handles; previews purely in the DOM (no model op). */
  beginCrop(rect: Rect, src: string, crop?: CropInsets | null): void;
  /** Re-pin the crop overlay after a zoom/relayout (the live window is held in
   *  crop fractions, so it survives the rescale). No-op when not cropping. */
  refreshCrop(rect: Rect): void;
  /** Whether crop mode is active. */
  isCropping(): boolean;
  /** Leave crop mode and return the final insets the caller should commit — `null`
   *  when the window covers the whole source (i.e. clear the crop). */
  endCrop(): CropInsets | null;
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

  // ---- crop mode ---------------------------------------------------------
  // A separate overlay (independent of the resize frame above) that mirrors the
  // resize handle protocol: the whole crop session previews purely in the DOM —
  // the faded full source behind a bright, draggable window with 8 handles — and
  // the model is mutated ONCE, on exit (Esc / click-away), as a single undo step.
  // The live window is held in crop FRACTIONS so it survives a zoom/relayout.
  const cropFull = document.createElement("div"); // whole source, dimmed underneath
  cropFull.style.cssText =
    "position:absolute;display:none;pointer-events:none;z-index:3;" +
    "background-repeat:no-repeat;background-position:top left;";
  const cropDim = document.createElement("div"); // darkening veil over the trimmed area
  cropDim.style.cssText =
    "position:absolute;display:none;pointer-events:none;z-index:4;background:rgba(0,0,0,0.45);";
  const cropWin = document.createElement("div"); // the kept window, bright (undimmed)
  cropWin.style.cssText =
    "position:absolute;display:none;pointer-events:none;z-index:5;overflow:hidden;" +
    "background-repeat:no-repeat;";
  const cropFrame = document.createElement("div"); // window border + handles
  cropFrame.style.cssText =
    "position:absolute;display:none;border:2px solid #1a73e8;box-sizing:border-box;" +
    "pointer-events:none;z-index:6;";
  const cropHandleEls: HTMLDivElement[] = HANDLES.map((h) => {
    const el = document.createElement("div");
    el.dataset["handle"] = h.name;
    el.className = "cw-obj-handle";
    el.style.cssText =
      "position:absolute;width:8px;height:8px;background:#fff;border:1.5px solid #1a73e8;" +
      `pointer-events:auto;cursor:${h.cursor};` +
      `left:calc(${h.dx * 100}% - 4px);top:calc(${h.dy * 100}% - 4px);`;
    cropFrame.appendChild(el);
    return el;
  });

  let cropSession: { rect: Rect; src: string; originalCrop: CropInsets; liveCrop: CropInsets } | null = null;

  const setBox = (el: HTMLElement, r: Rect, z: number): void => {
    el.style.left = `${r.x * z}px`;
    el.style.top = `${r.y * z}px`;
    el.style.width = `${r.width * z}px`;
    el.style.height = `${r.height * z}px`;
  };

  const paintCrop = (): void => {
    const s = cropSession;
    if (!s) return;
    const z = deps.getZoom();
    const full = fullSourceRect(s.rect, s.originalCrop);
    const win = windowRectFromInsets(full, s.liveCrop);
    setBox(cropFull, full, z);
    cropFull.style.backgroundSize = `${full.width * z}px ${full.height * z}px`;
    setBox(cropDim, full, z);
    setBox(cropWin, win, z);
    cropWin.style.backgroundSize = `${full.width * z}px ${full.height * z}px`;
    cropWin.style.backgroundPosition = `${-(win.x - full.x) * z}px ${-(win.y - full.y) * z}px`;
    setBox(cropFrame, win, z);
  };

  const hideCropOverlay = (): void => {
    for (const el of [cropFull, cropDim, cropWin, cropFrame]) el.style.display = "none";
    cropFull.style.backgroundImage = "";
    cropWin.style.backgroundImage = "";
  };

  const showCropOverlay = (rect: Rect, src: string, crop: CropInsets): void => {
    const host = deps.getPageElement(rect.pageIndex);
    if (!host) return;
    frame.style.display = "none"; // crop mode supersedes the resize frame
    hideGhost();
    cropSession = { rect, src, originalCrop: crop, liveCrop: crop };
    cropFull.style.backgroundImage = `url("${src}")`;
    cropWin.style.backgroundImage = `url("${src}")`;
    for (const el of [cropFull, cropDim, cropWin, cropFrame]) {
      if (el.parentElement !== host) host.appendChild(el);
      el.style.display = "block";
    }
    paintCrop();
  };

  // Crop handle drag — its own pointer-capture loop (like the resize handles), but
  // it only moves the window in the DOM; no model op fires until endCrop().
  let cropDragEl: HTMLElement | null = null;
  let cropPointerId = 0;
  let cropDrag: { spec: HandleSpec; startX: number; startY: number; full: Rect; startWin: Rect } | null = null;
  let cropMoveRaf = 0;
  let cropMovePending = false;

  const winFromCropDrag = (ev: MouseEvent): Rect => {
    const d = cropDrag!;
    const z = deps.getZoom();
    const dx = (ev.clientX - d.startX) / z;
    const dy = (ev.clientY - d.startY) / z;
    const F = d.full;
    const MIN = 8; // window never shrinks below 8 document px on either axis
    let { x, y, width, height } = d.startWin;
    if (d.spec.dx === 0) {
      const nx = Math.min(d.startWin.x + d.startWin.width - MIN, Math.max(F.x, d.startWin.x + dx));
      width = d.startWin.x + d.startWin.width - nx;
      x = nx;
    } else if (d.spec.dx === 1) {
      const right = Math.min(F.x + F.width, Math.max(d.startWin.x + MIN, d.startWin.x + d.startWin.width + dx));
      width = right - d.startWin.x;
    }
    if (d.spec.dy === 0) {
      const ny = Math.min(d.startWin.y + d.startWin.height - MIN, Math.max(F.y, d.startWin.y + dy));
      height = d.startWin.y + d.startWin.height - ny;
      y = ny;
    } else if (d.spec.dy === 1) {
      const bottom = Math.min(F.y + F.height, Math.max(d.startWin.y + MIN, d.startWin.y + d.startWin.height + dy));
      height = bottom - d.startWin.y;
    }
    return { x, y, width, height, pageIndex: F.pageIndex };
  };

  const flushCropMove = (): void => {
    cropMoveRaf = 0;
    if (cropMovePending) paintCrop();
    cropMovePending = false;
  };

  const onCropDragMove = (ev: PointerEvent): void => {
    if (!cropDrag || !cropSession) return;
    cropSession.liveCrop = cropInsetsFromWindow(cropDrag.full, winFromCropDrag(ev));
    cropMovePending = true;
    if (!cropMoveRaf) cropMoveRaf = requestAnimationFrame(flushCropMove);
  };

  const endCropDrag = (): void => {
    if (!cropDragEl) return;
    cropDragEl.removeEventListener("pointermove", onCropDragMove);
    cropDragEl.removeEventListener("pointerup", onCropDragUp);
    cropDragEl.removeEventListener("pointercancel", onCropDragUp);
    try {
      cropDragEl.releasePointerCapture(cropPointerId);
    } catch {
      /* pointer already released */
    }
    cropDragEl = null;
  };

  function onCropDragUp(ev: PointerEvent): void {
    if (!cropDrag) return;
    onCropDragMove(ev); // settle on the final window
    if (cropMoveRaf) cancelAnimationFrame(cropMoveRaf);
    cropMoveRaf = 0;
    cropMovePending = false;
    cropDrag = null;
    endCropDrag();
    paintCrop();
    // Write the crop live (transient) now that the handle is released — the pointer
    // capture is already gone, so the relayout this triggers can't kill the drag.
    if (cropSession) deps.onCropPreview(cropSession.liveCrop);
  }

  const onCropHandleDown = (ev: PointerEvent): void => {
    if (!cropSession) return;
    const el = ev.currentTarget as HTMLElement;
    const spec = HANDLES.find((h) => h.name === el.dataset["handle"]);
    if (!spec) return;
    ev.preventDefault();
    ev.stopPropagation();
    const full = fullSourceRect(cropSession.rect, cropSession.originalCrop);
    cropDrag = { spec, startX: ev.clientX, startY: ev.clientY, full, startWin: windowRectFromInsets(full, cropSession.liveCrop) };
    cropDragEl = el;
    cropPointerId = ev.pointerId;
    el.setPointerCapture(ev.pointerId);
    el.addEventListener("pointermove", onCropDragMove);
    el.addEventListener("pointerup", onCropDragUp);
    el.addEventListener("pointercancel", onCropDragUp);
  };

  for (const el of cropHandleEls) el.addEventListener("pointerdown", onCropHandleDown);

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
    beginCrop(rect: Rect, src: string, crop?: CropInsets | null): void {
      showCropOverlay(rect, src, crop ?? NO_CROP);
    },
    refreshCrop(rect: Rect): void {
      if (!cropSession) return;
      // A relayout/remote op can move the image to another page; follow it so the
      // crop chrome doesn't linger on the old page (mirrors the resize `show`).
      const host = deps.getPageElement(rect.pageIndex);
      if (host) {
        for (const el of [cropFull, cropDim, cropWin, cropFrame]) {
          if (el.parentElement !== host) host.appendChild(el);
        }
      }
      cropSession.rect = rect; // track reflow/zoom; the live window is held in fractions
      paintCrop();
    },
    isCropping(): boolean {
      return cropSession !== null;
    },
    endCrop(): CropInsets | null {
      const s = cropSession;
      if (cropDrag) {
        cropDrag = null;
        endCropDrag();
      }
      if (cropMoveRaf) cancelAnimationFrame(cropMoveRaf);
      cropMoveRaf = 0;
      cropMovePending = false;
      hideCropOverlay();
      cropSession = null;
      if (!s) return null;
      return isEmptyCrop(s.liveCrop) ? null : s.liveCrop;
    },
    destroy(): void {
      cancelPendingMove(); // drop any in-flight resize frame so it can't fire post-teardown
      if (cropMoveRaf) cancelAnimationFrame(cropMoveRaf);
      hideGhost();
      ghost.remove();
      sizeGhost.remove();
      frame.remove();
      cropFull.remove();
      cropDim.remove();
      cropWin.remove();
      cropFrame.remove();
    },
  };
}
