// Table context toolbar: appears above a MULTI-cell selection (2+ rows/columns) with
// the common table-editing actions — merge, insert/delete rows & columns — so they're
// reachable without switching to the Table ribbon tab. One of the built-in
// ContextToolbars coordinated by the shared manager.

import { ICONS } from "./icons";
import { injectCssOnce } from "./styles";
import { createFloatingBar, type ContextToolbar } from "./contextToolbar";
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

const CSS = `
.cw-tablebar{position:fixed;display:none;align-items:center;gap:2px;z-index:41;
  background:#fff;border:1px solid #c8c6c4;border-radius:6px;padding:3px;
  box-shadow:0 3px 12px rgba(0,0,0,.18);}
.cw-tablebar button{width:28px;height:28px;display:inline-flex;align-items:center;
  justify-content:center;border:1px solid transparent;border-radius:4px;background:transparent;
  cursor:pointer;color:#323130;}
.cw-tablebar button:hover{background:#e1dfdd;}
.cw-tablebar button.danger:hover{background:#fde7e9;color:#a4262c;}
.cw-tablebar button svg{width:16px;height:16px;}
.cw-tablebar .sep{width:1px;height:18px;background:#e1dfdd;margin:0 2px;}
@media (prefers-color-scheme: dark){
  :root[data-theme="dark"] .cw-tablebar{background:#2b2b2b;border-color:#4a4a4a;}
  :root[data-theme="dark"] .cw-tablebar button{color:#e6e6e6;}
  :root[data-theme="dark"] .cw-tablebar button:hover{background:#3c3c3c;}
  :root[data-theme="dark"] .cw-tablebar .sep{background:#4a4a4a;}
}`;

/** The table bar as a ContextToolbar (priority 28 — below the image bar, above the
 *  hyperlink/text bars; a cell selection never coexists with those anyway). */
export function createTableContextToolbar(deps: TableContextToolbarDeps): ContextToolbar {
  injectCssOnce("cw-tablebar-styles", CSS);
  const fb = createFloatingBar({ className: "cw-tablebar", ariaLabel: "Table" });

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
