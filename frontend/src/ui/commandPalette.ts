// Ctrl/Cmd+K command palette (critique "Move 1"). Makes every ribbon command —
// ~145 of them — reachable by name without hunting through tabs, and doubles as
// keyboard-first discovery. Fed the live list of ribbon controls (label + a run
// thunk) plus any embedder commands. Registers with the surface manager (row 5)
// so it shares the one Escape / z-order.

import { injectCssOnce } from "./styles";
import { openSurface, type SurfaceHandle } from "./surfaceManager";

export interface PaletteCommand {
  id: string;
  label: string;
  /** Where it lives (tab/group) or a category — shown dimmed on the right. */
  hint?: string;
  run: () => void;
}

const CSS = `
.cw-palette-backdrop{position:fixed;inset:0;z-index:1500;background:rgba(20,22,26,.28);display:flex;justify-content:center;align-items:flex-start;}
.cw-palette{margin-top:12vh;width:min(560px,92vw);max-height:64vh;background:#fff;border:1px solid #d0d5dd;border-radius:12px;
  box-shadow:0 18px 60px rgba(0,0,0,.28);display:flex;flex-direction:column;overflow:hidden;font:13px/1.4 system-ui,Segoe UI,Arial,sans-serif;color:#1a1a2e;}
.cw-palette-input{border:none;border-bottom:1px solid #ececec;padding:13px 16px;font:inherit;font-size:15px;color:#201f1e;outline:none;}
.cw-palette-list{overflow-y:auto;padding:6px;}
.cw-palette-item{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:7px;cursor:pointer;}
.cw-palette-item .lbl{flex:1 1 auto;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.cw-palette-item .hint{flex:0 0 auto;color:#9aa0a6;font-size:11.5px;}
.cw-palette-item.sel{background:#eef3fb;}
.cw-palette-item.sel .hint{color:#6b7280;}
.cw-palette-empty{padding:18px 16px;color:#80868b;font-size:12.5px;}
:root[data-theme="dark"] .cw-palette{background:#26282c;border-color:#4a4a4a;color:#e6e6e6;}
:root[data-theme="dark"] .cw-palette-input{background:#26282c;color:#f0f0f0;border-bottom-color:#3a3a3a;}
:root[data-theme="dark"] .cw-palette-item.sel{background:#2a3f5f;}
:root[data-theme="dark"] .cw-palette-item .hint{color:#7a808a;}
:root[data-theme="dark"] .cw-palette-item.sel .hint{color:#aab3c0;}
:root[data-theme="dark"] .cw-palette-empty{color:#80868b;}
`;

/** Open the palette over `commands`. Non-blocking; runs the chosen command and
 *  closes. Only one palette at a time. */
export function showCommandPalette(commands: PaletteCommand[]): void {
  injectCssOnce("cw-palette-styles", CSS);

  const backdrop = document.createElement("div");
  backdrop.className = "cw-palette-backdrop";
  const box = document.createElement("div");
  box.className = "cw-palette";
  box.setAttribute("role", "dialog");
  box.setAttribute("aria-label", "Command palette");
  const input = document.createElement("input");
  input.className = "cw-palette-input";
  input.type = "text";
  input.placeholder = "Type a command…";
  input.setAttribute("aria-label", "Type a command");
  const list = document.createElement("div");
  list.className = "cw-palette-list";
  box.append(input, list);
  backdrop.append(box);
  // Clicking the dim backdrop (outside the box) closes it.
  backdrop.addEventListener("mousedown", (e) => { if (e.target === backdrop) close(); });
  document.body.appendChild(backdrop);

  let surface: SurfaceHandle | null = null;
  let closed = false;
  const close = (): void => {
    if (closed) return;
    closed = true;
    surface?.release();
    backdrop.remove();
  };
  surface = openSurface({ el: backdrop, close, kind: "overlay" });

  let filtered: PaletteCommand[] = [];
  let sel = 0;

  const score = (c: PaletteCommand, tokens: string[]): boolean => {
    const hay = `${c.label} ${c.hint ?? ""}`.toLowerCase();
    return tokens.every((t) => hay.includes(t));
  };

  const render = (): void => {
    const tokens = input.value.toLowerCase().split(/\s+/).filter(Boolean);
    filtered = tokens.length === 0 ? commands.slice(0, 400) : commands.filter((c) => score(c, tokens)).slice(0, 400);
    sel = 0;
    list.textContent = "";
    if (filtered.length === 0) {
      const empty = document.createElement("div");
      empty.className = "cw-palette-empty";
      empty.textContent = "No matching commands.";
      list.appendChild(empty);
      return;
    }
    filtered.forEach((c, i) => {
      const row = document.createElement("div");
      row.className = "cw-palette-item" + (i === sel ? " sel" : "");
      const lbl = document.createElement("span");
      lbl.className = "lbl";
      lbl.textContent = c.label;
      row.appendChild(lbl);
      if (c.hint) {
        const hint = document.createElement("span");
        hint.className = "hint";
        hint.textContent = c.hint;
        row.appendChild(hint);
      }
      row.addEventListener("mousedown", (e) => { e.preventDefault(); runAt(i); });
      row.addEventListener("mousemove", () => setSel(i));
      list.appendChild(row);
    });
  };

  const setSel = (i: number): void => {
    if (i === sel || i < 0 || i >= filtered.length) return;
    list.children[sel]?.classList.remove("sel");
    sel = i;
    const el = list.children[sel] as HTMLElement | undefined;
    el?.classList.add("sel");
    el?.scrollIntoView({ block: "nearest" });
  };

  const runAt = (i: number): void => {
    const cmd = filtered[i];
    close();
    cmd?.run();
  };

  input.addEventListener("input", render);
  input.addEventListener("keydown", (e) => {
    if (e.key === "ArrowDown") { e.preventDefault(); setSel(Math.min(sel + 1, filtered.length - 1)); }
    else if (e.key === "ArrowUp") { e.preventDefault(); setSel(Math.max(sel - 1, 0)); }
    else if (e.key === "Enter") { e.preventDefault(); runAt(sel); }
    // Escape is handled centrally by the surface manager.
    e.stopPropagation();
  });

  render();
  setTimeout(() => input.focus(), 0);
}
