// List-item context toolbar: appears at the caret inside a bulleted/numbered list —
// promote/demote the level, or remove the list. Collapsed-caret only.

import { ICONS } from "./icons";
import { createFloatingBar, injectCtxBarCss, type ContextToolbar } from "./contextToolbar";
import type { AnchorRect } from "./floatingBarPosition";

export interface ListContextToolbarActions {
  promote(): void; // decrease level (outdent)
  demote(): void; // increase level (indent)
  removeList(): void;
}

export interface ListContextToolbarDeps {
  /** Whether the caret paragraph is in a list ("bullet"/"number"), else null. */
  listKind(): "bullet" | "number" | null;
  /** True when a non-empty range is selected (defers to the format bar then). */
  hasRangeSelection(): boolean;
  anchorRect(): AnchorRect | null;
  actions: ListContextToolbarActions;
}

/** The list-item bar as a ContextToolbar (priority 18). */
export function createListContextToolbar(deps: ListContextToolbarDeps): ContextToolbar {
  injectCtxBarCss();
  const fb = createFloatingBar({ className: "cw-ctxbar", ariaLabel: "List" });

  const iconBtn = (icon: string, title: string, onClick: () => void): void => {
    const b = document.createElement("button");
    b.type = "button";
    b.innerHTML = icon;
    b.title = title;
    b.addEventListener("click", onClick);
    fb.el.appendChild(b);
  };
  const textBtn = (label: string, title: string, onClick: () => void): void => {
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = label;
    b.title = title;
    b.addEventListener("click", onClick);
    fb.el.appendChild(b);
  };
  const sep = (): void => {
    const s = document.createElement("div");
    s.className = "sep";
    fb.el.appendChild(s);
  };

  iconBtn(ICONS.indentDecrease, "Promote (decrease level)", () => deps.actions.promote());
  iconBtn(ICONS.indentIncrease, "Demote (increase level)", () => deps.actions.demote());
  sep();
  textBtn("Remove", "Remove list formatting", () => deps.actions.removeList());

  return {
    id: "list",
    priority: 18,
    resolve: (): AnchorRect | null => {
      if (deps.hasRangeSelection() || deps.listKind() === null) return null;
      return deps.anchorRect();
    },
    show: (anchor, viewport) => fb.place(anchor, viewport),
    hide: () => fb.hide(),
    destroy: () => fb.destroy(),
  };
}
