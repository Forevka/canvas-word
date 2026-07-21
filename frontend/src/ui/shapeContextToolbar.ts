// Shape context toolbar (Word's drawing-shape hover bar): fill, outline, wrap, align,
// z-order and delete, shown above a selected drawing shape. Mirrors imageContextToolbar
// as a ContextToolbar so the shared manager coordinates it with the image/text/link
// bars. Priority 31 — just above the image bar (30), so a selected shape always wins.
// Fill and Outline open the host's shared colour popover (anchored to their button),
// the same picker font-colour / highlight use — no new colour UI.

import { ICONS } from "./icons";
import { createFloatingBar, type ContextToolbar } from "./contextToolbar";
import type { AnchorRect } from "./floatingBarPosition";

export interface ShapeContextToolbarActions {
  /** Open the fill picker anchored to `anchor` (the Fill button). */
  fill(anchor: HTMLElement): void;
  /** Open the outline picker (colour + width + dash) anchored to `anchor`. */
  outline(anchor: HTMLElement): void;
  wrapInline(): void;
  wrapSquare(): void;
  alignLeft(): void;
  alignCenter(): void;
  alignRight(): void;
  bringToFront(): void;
  sendToBack(): void;
  remove(): void;
}

export interface ShapeContextToolbarDeps {
  /** Viewport rect of the selected shape, or null when none is selected. */
  anchorRect(): AnchorRect | null;
  actions: ShapeContextToolbarActions;
}

/** The shape bar as a ContextToolbar (priority 31 — above the image bar). */
export function createShapeContextToolbar(deps: ShapeContextToolbarDeps): ContextToolbar {
  const fb = createFloatingBar({ className: "cw-img-toolbar", ariaLabel: "Shape" });

  const ibtn = (icon: string, title: string, onClick: (btn: HTMLButtonElement) => void, cls = ""): void => {
    const b = document.createElement("button");
    if (cls) b.className = cls;
    b.innerHTML = icon;
    b.title = title;
    b.addEventListener("click", () => onClick(b));
    fb.el.appendChild(b);
  };
  const sep = (): void => {
    const s = document.createElement("div");
    s.className = "sep";
    fb.el.appendChild(s);
  };

  ibtn(ICONS.shapeFill, "Fill colour", (b) => deps.actions.fill(b));
  ibtn(ICONS.outline, "Outline (colour, width, dash)", (b) => deps.actions.outline(b));
  sep();
  ibtn(ICONS.wrapInline, "In line with text", () => deps.actions.wrapInline());
  ibtn(ICONS.wrapSquare, "Wrap text (square)", () => deps.actions.wrapSquare());
  sep();
  ibtn(ICONS.alignLeft, "Align left", () => deps.actions.alignLeft());
  ibtn(ICONS.alignCenter, "Align center", () => deps.actions.alignCenter());
  ibtn(ICONS.alignRight, "Align right", () => deps.actions.alignRight());
  sep();
  ibtn(ICONS.bringFront, "Bring to front", () => deps.actions.bringToFront());
  ibtn(ICONS.sendBack, "Send to back", () => deps.actions.sendToBack());
  sep();
  ibtn(ICONS.trash, "Delete shape (Del)", () => deps.actions.remove(), "danger");

  return {
    id: "shape",
    priority: 31,
    resolve: (): AnchorRect | null => deps.anchorRect(),
    show: (anchor, viewport) => fb.place(anchor, viewport),
    hide: () => fb.hide(),
    destroy: () => fb.destroy(),
  };
}
