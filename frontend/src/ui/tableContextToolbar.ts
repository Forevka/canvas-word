// Table context toolbar: appears above a MULTI-cell selection (2+ rows/columns) with
// the common table-editing actions — merge, insert/delete rows & columns — so they're
// reachable without switching to the Table ribbon tab. One of the built-in
// ContextToolbars coordinated by the shared manager.

import { ICONS } from "./icons";
import { createFloatingBar, injectCtxBarCss, type ContextToolbar } from "./contextToolbar";
import type { AnchorRect } from "./floatingBarPosition";

export interface TableContextToolbarActions {
  mergeCells(): void;
  insertRowAbove(): void;
  insertRowBelow(): void;
  insertColumnLeft(): void;
  insertColumnRight(): void;
  deleteRow(): void;
  deleteColumn(): void;
}

export interface TableContextToolbarDeps {
  /** Viewport bounding rect of the multi-cell selection, or null. */
  anchorRect(): AnchorRect | null;
  actions: TableContextToolbarActions;
}

/** The table bar as a ContextToolbar (priority 28 — below the image bar, above the
 *  hyperlink/text bars; a cell selection never coexists with those anyway). Shares the
 *  common `cw-ctxbar` base (`cw-iconbar` modifier) with the image/shape bars so the
 *  three floating object bars converge on one look (issue #244 C4). */
export function createTableContextToolbar(deps: TableContextToolbarDeps): ContextToolbar {
  injectCtxBarCss();
  const fb = createFloatingBar({ className: "cw-ctxbar cw-iconbar", ariaLabel: "Table" });

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

  btn(ICONS.mergeCells, "Merge cells", () => deps.actions.mergeCells());
  sep();
  btn(ICONS.rowAbove, "Insert row above", () => deps.actions.insertRowAbove());
  btn(ICONS.rowBelow, "Insert row below", () => deps.actions.insertRowBelow());
  btn(ICONS.colLeft, "Insert column left", () => deps.actions.insertColumnLeft());
  btn(ICONS.colRight, "Insert column right", () => deps.actions.insertColumnRight());
  sep();
  btn(ICONS.deleteRow, "Delete rows", () => deps.actions.deleteRow(), "danger");
  btn(ICONS.deleteCol, "Delete columns", () => deps.actions.deleteColumn(), "danger");

  return {
    id: "table",
    priority: 28,
    resolve: (): AnchorRect | null => deps.anchorRect(),
    show: (anchor, viewport) => fb.place(anchor, viewport),
    hide: () => fb.hide(),
    destroy: () => fb.destroy(),
  };
}
