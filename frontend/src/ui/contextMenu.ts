// Contextual right-click menu — a self-contained popup renderer. The editor
// (index.ts) detects what sits under the cursor and hands this module a list of
// sections; the matrix of which sections appear lives there, not here.
//
// Supports items, separators, group headers, one level of submenu, disabled and
// destructive styling, optional leading icons. Dismisses on outside-press,
// Escape, blur, or after an item runs. Clamps inside the viewport.

import { injectCssOnce } from "./styles";

export interface MenuItem {
  kind: "item";
  label: string;
  icon?: string; // innerHTML for a 16px glyph (optional)
  shortcut?: string;
  disabled?: boolean;
  danger?: boolean;
  onClick: () => void;
}

export interface MenuSubmenu {
  kind: "submenu";
  label: string;
  icon?: string;
  items: MenuEntry[];
}

export type MenuEntry =
  | MenuItem
  | MenuSubmenu
  | { kind: "sep" }
  | { kind: "header"; label: string };

export interface ContextMenuHandle {
  close(): void;
}

const MENU_CSS = `
.cw-menu{position:fixed;z-index:1300;min-width:200px;max-width:320px;background:#fff;
  border:1px solid #c8ccd1;border-radius:8px;padding:5px;
  box-shadow:0 6px 22px rgba(0,0,0,.22);font:13px Arial,sans-serif;color:#2b2f33;
  user-select:none;}
.cw-menu-item{display:flex;align-items:center;gap:9px;padding:6px 10px;border-radius:5px;
  cursor:default;white-space:nowrap;}
.cw-menu-item:hover:not(.cw-disabled){background:#e8eef9;}
.cw-menu-item.cw-disabled{color:#9aa0a6;cursor:default;}
.cw-menu-item.cw-danger:hover:not(.cw-disabled){background:#fce8e6;color:#c5221f;}
.cw-menu-ico{width:16px;height:16px;display:inline-flex;align-items:center;justify-content:center;
  flex:0 0 16px;color:#5f6368;}
.cw-menu-ico svg{width:16px;height:16px;}
.cw-menu-lbl{flex:1 1 auto;}
.cw-menu-acc{color:#9aa0a6;font-size:12px;margin-left:14px;}
.cw-menu-arrow{color:#9aa0a6;margin-left:10px;}
.cw-menu-sep{height:1px;background:#e6e8eb;margin:5px 8px;}
.cw-menu-header{padding:5px 10px 2px;font-size:11px;color:#80868b;text-transform:uppercase;letter-spacing:.04em;}`;

/** Render `entries` as a popup at (clientX, clientY). */
export function showContextMenu(
  clientX: number,
  clientY: number,
  entries: MenuEntry[],
): ContextMenuHandle {
  injectCssOnce("cw-menu-styles", MENU_CSS);
  let submenuEl: HTMLElement | null = null;

  const root = buildPanel(entries, () => close(), true);
  document.body.appendChild(root);
  placeAt(root, clientX, clientY);

  function close(): void {
    submenuEl?.remove();
    submenuEl = null;
    root.remove();
    window.removeEventListener("mousedown", onOutside, true);
    window.removeEventListener("keydown", onKey, true);
    window.removeEventListener("blur", onBlurClose);
    window.removeEventListener("resize", onBlurClose);
  }

  const onOutside = (ev: MouseEvent): void => {
    const t = ev.target as Node;
    if (root.contains(t) || submenuEl?.contains(t)) return;
    close();
  };
  const onKey = (ev: KeyboardEvent): void => {
    if (ev.key === "Escape") {
      ev.preventDefault();
      ev.stopPropagation();
      close();
    }
  };
  const onBlurClose = (): void => close();

  // Defer listener attach a tick so the opening right-click doesn't self-close.
  setTimeout(() => {
    window.addEventListener("mousedown", onOutside, true);
    window.addEventListener("keydown", onKey, true);
    window.addEventListener("blur", onBlurClose);
    window.addEventListener("resize", onBlurClose);
  }, 0);

  // Submenu plumbing: a submenu row opens its panel to the side on hover. The
  // submenu overlaps the parent's right edge (no gap), so the pointer reaches
  // it without crossing other rows.
  function attachSubmenu(row: HTMLElement, sub: MenuSubmenu): void {
    row.addEventListener("mouseenter", () => {
      submenuEl?.remove();
      submenuEl = buildPanel(sub.items, () => close(), false);
      document.body.appendChild(submenuEl);
      const r = row.getBoundingClientRect();
      placeAt(submenuEl, r.right - 4, r.top - 5, r);
    });
  }

  /** `isRoot` rows close any open submenu on hover; submenu-panel rows must NOT
   *  (else hovering a sub-item would remove the very submenu it lives in). */
  function buildPanel(items: MenuEntry[], dismiss: () => void, isRoot: boolean): HTMLElement {
    const panel = document.createElement("div");
    panel.className = "cw-menu";
    panel.addEventListener("contextmenu", (e) => e.preventDefault());
    // Keep the editor's focus (its hidden IME proxy) while the menu is used:
    // items fire on mouseup, so suppressing the default mousedown stops the menu
    // from stealing focus. Otherwise focus falls to <body> and the next
    // keystroke — typically Ctrl+Z to undo the menu action — is dropped until
    // the user clicks back into the document.
    panel.addEventListener("mousedown", (e) => e.preventDefault());
    for (const entry of items) {
      if (entry.kind === "sep") {
        const s = document.createElement("div");
        s.className = "cw-menu-sep";
        panel.appendChild(s);
        continue;
      }
      if (entry.kind === "header") {
        const h = document.createElement("div");
        h.className = "cw-menu-header";
        h.textContent = entry.label;
        panel.appendChild(h);
        continue;
      }
      const row = document.createElement("div");
      row.className = "cw-menu-item";
      const ico = document.createElement("span");
      ico.className = "cw-menu-ico";
      if (entry.icon) ico.innerHTML = entry.icon;
      const lbl = document.createElement("span");
      lbl.className = "cw-menu-lbl";
      lbl.textContent = entry.label;
      row.append(ico, lbl);
      if (entry.kind === "submenu") {
        const arrow = document.createElement("span");
        arrow.className = "cw-menu-arrow";
        arrow.textContent = "›";
        row.appendChild(arrow);
        attachSubmenu(row, entry);
      } else {
        if (entry.shortcut) {
          const acc = document.createElement("span");
          acc.className = "cw-menu-acc";
          acc.textContent = entry.shortcut;
          row.appendChild(acc);
        }
        if (entry.disabled) row.classList.add("cw-disabled");
        if (entry.danger) row.classList.add("cw-danger");
        if (!entry.disabled) {
          row.addEventListener("mouseup", () => {
            dismiss();
            entry.onClick();
          });
        }
      }
      // A ROOT row that is NOT a submenu closes any open submenu on hover —
      // never on submenu-panel rows, or navigating into one would dismiss it.
      if (isRoot && entry.kind !== "submenu") {
        row.addEventListener("mouseenter", () => {
          submenuEl?.remove();
          submenuEl = null;
        });
      }
      panel.appendChild(row);
    }
    return panel;
  }

  return { close };
}

/** Position a panel at (x,y), flipping/clamping to stay on-screen. `avoid` is a
 *  rect (the parent row) a submenu should flip to the left of when needed. */
function placeAt(el: HTMLElement, x: number, y: number, avoid?: DOMRect): void {
  const vw = window.innerWidth;
  const vh = window.innerHeight;
  const r = el.getBoundingClientRect();
  let left = x;
  let top = y;
  if (left + r.width > vw - 6) {
    left = avoid ? avoid.left - r.width + 4 : vw - r.width - 6;
  }
  if (top + r.height > vh - 6) top = Math.max(6, vh - r.height - 6);
  el.style.left = `${Math.max(6, left)}px`;
  el.style.top = `${Math.max(6, top)}px`;
}
