// Equation context toolbar: appears above a selected equation block — edit (opens the
// equation editor), align, delete. One of the built-in ContextToolbars.

import { ICONS } from "./icons";
import { createFloatingBar, injectCtxBarCss, type ContextToolbar } from "./contextToolbar";
import type { AnchorRect } from "./floatingBarPosition";

export interface EquationContextToolbarActions {
  edit(): void;
  alignLeft(): void;
  alignCenter(): void;
  alignRight(): void;
  remove(): void;
}

export interface EquationContextToolbarDeps {
  /** Viewport rect of the selected equation, or null when none is selected. */
  anchorRect(): AnchorRect | null;
  actions: EquationContextToolbarActions;
}

/** The equation bar as a ContextToolbar (priority 29 — just under the image bar). */
export function createEquationContextToolbar(deps: EquationContextToolbarDeps): ContextToolbar {
  injectCtxBarCss();
  const fb = createFloatingBar({ className: "cw-ctxbar", ariaLabel: "Equation" });

  const btn = (icon: string, title: string, onClick: () => void, cls = ""): void => {
    const b = document.createElement("button");
    b.type = "button";
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

  btn(ICONS.stylePencil, "Edit equation", () => deps.actions.edit());
  sep();
  btn(ICONS.alignLeft, "Align left", () => deps.actions.alignLeft());
  btn(ICONS.alignCenter, "Align center", () => deps.actions.alignCenter());
  btn(ICONS.alignRight, "Align right", () => deps.actions.alignRight());
  sep();
  btn(ICONS.trash, "Delete equation", () => deps.actions.remove(), "danger");

  return {
    id: "equation",
    priority: 29,
    resolve: (): AnchorRect | null => deps.anchorRect(),
    show: (anchor: AnchorRect): void => fb.place(anchor),
    hide: () => fb.hide(),
    destroy: () => fb.destroy(),
  };
}
