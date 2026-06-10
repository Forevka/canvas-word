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
  /** Live preview during handle drag (transient). */
  onResizePreview(width: number, height: number): void;
  /** Final size on mouseup (one undoable op). */
  onResizeCommit(width: number, height: number): void;
}

export interface ObjectFrame {
  /** Show the frame over a rect (page coords) for an image of natural ratio. */
  show(rect: Rect, maxWidth: number): void;
  hide(): void;
  destroy(): void;
}

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

  let current: { rect: Rect; maxWidth: number } | null = null;
  let drag: {
    spec: HandleSpec;
    startX: number;
    startY: number;
    startW: number;
    startH: number;
  } | null = null;

  // Pointer Events (not mouse) so the handles drag by finger/pen too. Pointer
  // capture routes move/up to the captured handle even when the finger leaves it.
  let dragEl: HTMLElement | null = null;
  let dragPointerId = 0;

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
    };
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
    const { w, h } = sizeFromDrag(ev);
    deps.onResizePreview(w, h);
  };

  const onDragUp = (ev: PointerEvent): void => {
    if (!drag) return;
    const { w, h } = sizeFromDrag(ev);
    drag = null;
    endDrag();
    deps.onResizeCommit(w, h);
  };

  for (const el of handleEls) el.addEventListener("pointerdown", onHandleDown);

  return {
    show(rect: Rect, maxWidth: number): void {
      const host = deps.getPageElement(rect.pageIndex);
      if (!host) return;
      current = { rect, maxWidth };
      if (frame.parentElement !== host) host.appendChild(frame);
      const z = deps.getZoom();
      frame.style.display = "block";
      frame.style.left = `${rect.x * z - 2}px`;
      frame.style.top = `${rect.y * z - 2}px`;
      frame.style.width = `${rect.width * z + 4}px`;
      frame.style.height = `${rect.height * z + 4}px`;
    },
    hide(): void {
      current = null;
      frame.style.display = "none";
    },
    destroy(): void {
      frame.remove();
    },
  };
}
