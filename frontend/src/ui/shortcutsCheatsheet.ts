// Keyboard-shortcuts cheat sheet (Ctrl+/, critique A3). Built on the shared
// dialog shell (surface-managed, Escape/×-closable). The editing chords come
// from KEYMAP_ENTRIES — the same declarative list the keymap handler is defined
// beside — so the sheet can't silently drift from what the keys actually do; the
// app-chrome / clipboard groups are listed here, and any embedder command with a
// keybinding is appended live.

import { createDialogShell } from "./dialogShell";
import { KEYMAP_ENTRIES } from "../input/keymap";

export interface ShortcutGroup {
  title: string;
  items: { chord: string; label: string }[];
}

const CSS = `
.cw-keys-backdrop{position:fixed;inset:0;z-index:1100;background:rgba(20,22,26,.32);}
.cw-keys-modal{position:fixed;min-width:520px;max-width:min(94vw,720px);background:#fff;border:1px solid #d0d5dd;border-radius:12px;
  box-shadow:0 16px 50px rgba(0,0,0,.26);font:13px/1.45 system-ui,Segoe UI,Arial,sans-serif;color:#1a1a2e;display:flex;flex-direction:column;overflow:hidden;}
.cw-keys-head{display:flex;align-items:center;justify-content:space-between;padding:12px 16px 10px;cursor:move;}
.cw-keys-head h2{margin:0;font-size:15px;font-weight:600;color:#201f1e;}
.cw-keys-x{border:none;background:transparent;font-size:20px;line-height:1;cursor:pointer;color:#605e5c;}
.cw-keys-x:hover{color:#201f1e;}
.cw-keys-body{padding:4px 16px 16px;overflow:auto;max-height:70vh;columns:2;column-gap:28px;}
.cw-keys-group{break-inside:avoid;margin-bottom:14px;}
.cw-keys-group h3{margin:0 0 6px;font-size:11px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;color:#8a8f98;}
.cw-keys-row{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:3px 0;}
.cw-keys-row .lbl{color:#323130;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.cw-keys-row .keys{flex:0 0 auto;display:flex;gap:3px;}
.cw-keys-row kbd{font:inherit;font-size:11px;background:#f3f2f1;border:1px solid #d8d6d4;border-bottom-width:2px;border-radius:4px;padding:1px 6px;color:#3c4043;min-width:16px;text-align:center;}
.cw-keys-foot{padding:0;}
:root[data-theme="dark"] .cw-keys-modal{background:#26282c;color:#e6e6e6;border-color:#4a4a4a;}
:root[data-theme="dark"] .cw-keys-head h2{color:#f0f0f0;}
:root[data-theme="dark"] .cw-keys-row .lbl{color:#d0d0d0;}
:root[data-theme="dark"] .cw-keys-row kbd{background:#1e1f22;border-color:#4a4a4a;color:#d0d0d0;}
`;

const isMac = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent);

/** Render a chord like "Ctrl+Shift+Z" as platform-appropriate <kbd> keys. */
function keysFor(chord: string): HTMLElement {
  const wrap = document.createElement("span");
  wrap.className = "keys";
  for (const part of chord.split("+")) {
    const k = document.createElement("kbd");
    k.textContent = isMac
      ? part === "Ctrl" || part === "Mod" ? "⌘" : part === "Alt" ? "⌥" : part === "Shift" ? "⇧" : part
      : part === "Mod" ? "Ctrl" : part;
    wrap.appendChild(k);
  }
  return wrap;
}

/** Show the cheat sheet. `extraGroups` (e.g. embedder commands with keybindings)
 *  are appended after the built-ins. */
export function showShortcutsCheatsheet(extraGroups: ShortcutGroup[] = []): void {
  const groups: ShortcutGroup[] = [
    { title: "Formatting", items: KEYMAP_ENTRIES.filter((e) => ["Bold", "Italic", "Underline"].includes(e.label)) },
    { title: "History", items: KEYMAP_ENTRIES.filter((e) => ["Undo", "Redo"].includes(e.label)) },
    { title: "Breaks", items: KEYMAP_ENTRIES.filter((e) => e.label.includes("break")) },
    {
      title: "Clipboard",
      items: [
        { chord: "Ctrl+X", label: "Cut" },
        { chord: "Ctrl+C", label: "Copy" },
        { chord: "Ctrl+V", label: "Paste" },
      ],
    },
    {
      title: "Lists",
      items: [
        { chord: "Tab", label: "Demote list item" },
        { chord: "Shift+Tab", label: "Promote list item" },
      ],
    },
    {
      title: "Application",
      items: [
        { chord: "Ctrl+S", label: "Save" },
        { chord: "Ctrl+F", label: "Find & replace" },
        { chord: "Ctrl+K", label: "Command palette" },
        { chord: "Ctrl+/", label: "Keyboard shortcuts" },
        { chord: "Ctrl+F1", label: "Collapse / expand ribbon" },
      ],
    },
    ...extraGroups,
  ].filter((g) => g.items.length > 0);

  const shell = createDialogShell({ prefix: "cw-keys", cssId: "cw-keys-styles", css: CSS, title: "Keyboard shortcuts" });
  for (const g of groups) {
    const gEl = document.createElement("div");
    gEl.className = "cw-keys-group";
    const h = document.createElement("h3");
    h.textContent = g.title;
    gEl.appendChild(h);
    for (const it of g.items) {
      const row = document.createElement("div");
      row.className = "cw-keys-row";
      const lbl = document.createElement("span");
      lbl.className = "lbl";
      lbl.textContent = it.label;
      row.append(lbl, keysFor(it.chord));
      gEl.appendChild(row);
    }
    shell.body.appendChild(gEl);
  }
}
