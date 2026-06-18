// Shared behavior for the editor's editing dialogs (Borders & Shading, content-
// control inspector, field constructor, TOC options): turn a centered, page-dimming
// backdrop+modal into a DRAGGABLE, NON-BLOCKING floating panel.
//
// Why: these dialogs edit the visible document, so dimming/trapping the page hides
// the very change being made. As a floating panel the document stays visible and
// interactive (edits are seen live), clicks outside don't dismiss it, and it can be
// dragged by its header out of the way.

import { injectCssOnce } from "./styles";

const DRAG_CSS = `
.cw-drag-handle{cursor:move;user-select:none;}
.cw-drag-handle::before{content:"⠿";color:#b0b4ba;font-size:14px;line-height:1;margin-right:4px;flex:0 0 auto;}`;

export interface FloatingDialogOptions {
  /** The full-screen backdrop element (made click-through). */
  backdrop: HTMLElement;
  /** The dialog box (positioned + draggable). */
  modal: HTMLElement;
  /** The element to drag the dialog by (usually its header). */
  handle: HTMLElement;
  /** Aborted on close — detaches the drag listeners. */
  signal: AbortSignal;
  /** Selector for descendants of `handle` that must NOT start a drag (e.g. the
   *  close "×" button). */
  noDrag?: string;
}

/** Make a backdrop+modal dialog a draggable, non-blocking floating panel. Call once
 *  after the dialog is in the DOM (it reads the modal's measured width to place it). */
export function makeFloatingDialog(o: FloatingDialogOptions): void {
  injectCssOnce("cw-drag-handle-styles", DRAG_CSS);
  const { backdrop, modal, handle, signal } = o;

  // Non-blocking layer: the page shows through and stays interactive; clicking
  // outside no longer closes the dialog (a floating tool, like Word's).
  backdrop.style.pointerEvents = "none";
  backdrop.style.background = "transparent";
  backdrop.style.display = "block";

  // Free-floating, positioned panel.
  modal.style.position = "fixed";
  modal.style.margin = "0";
  modal.style.pointerEvents = "auto";
  handle.classList.add("cw-drag-handle");

  const place = (left: number, top: number): void => {
    const maxLeft = Math.max(4, window.innerWidth - modal.offsetWidth - 6);
    const maxTop = Math.max(4, window.innerHeight - 44);
    modal.style.left = `${Math.min(Math.max(4, left), maxLeft)}px`;
    modal.style.top = `${Math.min(Math.max(4, top), maxTop)}px`;
  };
  // Default to the right so it doesn't cover the (often left-aligned) content.
  place(window.innerWidth - modal.offsetWidth - 24, 72);

  let drag: { dx: number; dy: number } | null = null;
  handle.addEventListener(
    "mousedown",
    (ev) => {
      if (o.noDrag && (ev.target as HTMLElement).closest(o.noDrag)) return;
      const r = modal.getBoundingClientRect();
      drag = { dx: ev.clientX - r.left, dy: ev.clientY - r.top };
      ev.preventDefault();
    },
    { signal },
  );
  window.addEventListener("mousemove", (ev) => { if (drag) place(ev.clientX - drag.dx, ev.clientY - drag.dy); }, { signal });
  window.addEventListener("mouseup", () => { drag = null; }, { signal });
}
