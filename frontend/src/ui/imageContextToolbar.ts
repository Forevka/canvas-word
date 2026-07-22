// Image context toolbar (Word's image hover bar): wrap, align, delete — shown above
// a selected image. Extracted from the inline block in editorApp into a ContextToolbar
// so the shared manager coordinates it with the text/link bars. Shares the common
// `cw-ctxbar` base class (`cw-iconbar` modifier for the square icon buttons) with the
// shape/table bars so the three converge on one look (issue #244 C4).

import { ICONS } from "./icons";
import { createFloatingBar, injectCtxBarCss, type ContextToolbar } from "./contextToolbar";
import type { AnchorRect } from "./floatingBarPosition";

export interface ImageContextToolbarActions {
  wrapInline(): void;
  wrapSquare(): void;
  alignLeft(): void;
  alignCenter(): void;
  alignRight(): void;
  remove(): void;
}

export interface ImageContextToolbarDeps {
  /** Viewport rect of the selected image, or null when none is selected. */
  anchorRect(): AnchorRect | null;
  actions: ImageContextToolbarActions;
}

/** The image bar as a ContextToolbar (priority 30 — the highest, so a selected image
 *  always wins over a text selection). */
export function createImageContextToolbar(deps: ImageContextToolbarDeps): ContextToolbar {
  injectCtxBarCss();
  const fb = createFloatingBar({ className: "cw-ctxbar cw-iconbar", ariaLabel: "Image" });

  const ibtn = (icon: string, title: string, onClick: () => void, cls = ""): void => {
    const b = document.createElement("button");
    if (cls) b.className = cls;
    b.innerHTML = icon;
    b.title = title;
    b.addEventListener("click", onClick);
    fb.el.appendChild(b);
  };
  const sep = (): void => {
    const s = document.createElement("div");
    s.className = "sep";
    fb.el.appendChild(s);
  };

  ibtn(ICONS.wrapInline, "In line with text", () => deps.actions.wrapInline());
  ibtn(ICONS.wrapSquare, "Wrap text (square)", () => deps.actions.wrapSquare());
  sep();
  ibtn(ICONS.alignLeft, "Align left", () => deps.actions.alignLeft());
  ibtn(ICONS.alignCenter, "Align center", () => deps.actions.alignCenter());
  ibtn(ICONS.alignRight, "Align right", () => deps.actions.alignRight());
  sep();
  ibtn(ICONS.trash, "Delete image (Del)", () => deps.actions.remove(), "danger");

  return {
    id: "image",
    priority: 30,
    resolve: (): AnchorRect | null => deps.anchorRect(),
    show: (anchor, viewport) => fb.place(anchor, viewport),
    hide: () => fb.hide(),
    destroy: () => fb.destroy(),
  };
}
