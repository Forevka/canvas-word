// Settings (File ▸ Settings, and the "Settings" command-palette entry) — the home
// for application PREFERENCES that are not document-authoring commands: today the
// appearance theme (Light / Dark / Match system) and, once wired, the chrome preset
// (Ribbon / Minimal). Word's "File ▸ Options" convention, minus a top-level ribbon
// tab (a rarely-opened surface shouldn't spend prime ribbon real estate).
//
// Built on the shared createDialogShell (draggable, surface-managed z-order + single
// Escape, dark-mode themed) exactly like Page Layout / the input dialog — no new
// dialog mechanism. The structure is a list of GROUPS, each a list of ROWS, so future
// preferences (measurement units, autosave) slot in without a redesign. Every control
// applies LIVE (no Apply button) — picking a value calls its `set` immediately and the
// dialog stays open so the effect is visible.

import { createDialogShell } from "./dialogShell";

const CSS = `
.cw-set-backdrop{position:fixed;inset:0;z-index:1100;background:rgba(20,22,26,.30);}
.cw-set-modal{position:fixed;width:min(460px,94vw);max-height:88vh;display:flex;flex-direction:column;background:#fff;
  border:1px solid #d0d5dd;border-radius:10px;box-shadow:0 18px 56px rgba(0,0,0,.30);
  font:13px/1.5 system-ui,"Segoe UI",Arial,sans-serif;color:#1a1a2e;overflow:hidden;}
.cw-set-head{display:flex;align-items:center;justify-content:space-between;padding:12px 16px;border-bottom:1px solid #e6e8eb;cursor:move;}
.cw-set-head h2{margin:0;font-size:15px;font-weight:600;color:#201f1e;}
.cw-set-x{border:none;background:transparent;font-size:20px;line-height:1;cursor:pointer;color:#605e5c;width:28px;height:28px;border-radius:6px;}
.cw-set-x:hover{background:#e8eaed;color:#201f1e;}
.cw-set-body{padding:8px 16px 14px;overflow:auto;display:flex;flex-direction:column;gap:18px;}
.cw-set-group{display:flex;flex-direction:column;gap:10px;}
.cw-set-group-ttl{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#80868b;margin-top:6px;}
.cw-set-row{display:flex;flex-direction:column;gap:6px;}
.cw-set-row-lab{font-size:13px;font-weight:600;color:#3c4043;}
.cw-set-row-hint{font-size:11.5px;color:#80868b;margin-top:-2px;}
.cw-set-seg{display:inline-flex;align-self:flex-start;border:1px solid #d0d4d9;border-radius:8px;overflow:hidden;background:#f6f7f8;}
.cw-set-seg-btn{border:none;background:transparent;font:inherit;font-size:12.5px;color:#3c4043;cursor:pointer;padding:6px 14px;
  border-right:1px solid #e1e3e6;min-width:88px;}
.cw-set-seg-btn:last-child{border-right:none;}
.cw-set-seg-btn:hover{background:#eceef1;}
.cw-set-seg-btn.active{background:#2b579a;color:#fff;font-weight:600;}
.cw-set-foot{display:flex;justify-content:flex-end;gap:8px;padding:10px 16px;border-top:1px solid #e6e8eb;}
.cw-set-foot button{height:30px;padding:0 16px;border-radius:6px;font:inherit;font-size:13px;font-weight:600;cursor:pointer;
  border:1px solid #2b579a;background:#2b579a;color:#fff;}
:root[data-theme="dark"] .cw-set-modal{background:#26282c;color:#e6e6e6;border-color:#4a4a4a;}
:root[data-theme="dark"] .cw-set-head{border-color:#34373c;}
:root[data-theme="dark"] .cw-set-head h2{color:#f0f0f0;}
:root[data-theme="dark"] .cw-set-x:hover{background:#34373c;color:#f0f0f0;}
:root[data-theme="dark"] .cw-set-group-ttl,:root[data-theme="dark"] .cw-set-row-hint{color:#9aa0a6;}
:root[data-theme="dark"] .cw-set-row-lab{color:#dfe1e5;}
:root[data-theme="dark"] .cw-set-seg{background:#1e1f22;border-color:#4a4a4a;}
:root[data-theme="dark"] .cw-set-seg-btn{color:#dfe1e5;border-right-color:#34373c;}
:root[data-theme="dark"] .cw-set-seg-btn:hover{background:#34373c;}
:root[data-theme="dark"] .cw-set-seg-btn.active{background:#3d6ba5;color:#fff;}
:root[data-theme="dark"] .cw-set-foot{border-color:#34373c;}
:root[data-theme="dark"] .cw-set-foot button{background:#3d6ba5;border-color:#3d6ba5;}
`;

/** A segmented (radio-style) choice among mutually-exclusive values. */
export interface SettingsSegmentedRow {
  /** Stable id (namespaced, e.g. "appearance.theme") — for the a11y group + tests. */
  id: string;
  label: string;
  /** Optional one-line description under the label. */
  hint?: string;
  options: { value: string; label: string }[];
  /** Current value (read to seed + refresh the pressed state). */
  get: () => string;
  /** Apply the chosen value LIVE (no Apply button). */
  set: (value: string) => void;
}

export type SettingsRow = SettingsSegmentedRow; // one kind today; room for more

export interface SettingsGroup {
  title: string;
  rows: SettingsRow[];
}

export interface SettingsDialogOptions {
  groups: SettingsGroup[];
  onClose?: () => void;
}

const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string): HTMLElementTagNameMap[K] => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
};

/** Build one segmented control; returns its element and a refresh() that re-reads
 *  `row.get()` and repaints the pressed state (so a live change elsewhere reflects). */
function segmented(row: SettingsSegmentedRow): { wrap: HTMLElement; refresh: () => void } {
  const wrap = el("div", "cw-set-row");
  wrap.append(el("div", "cw-set-row-lab", row.label));
  if (row.hint) wrap.append(el("div", "cw-set-row-hint", row.hint));
  const seg = el("div", "cw-set-seg");
  seg.setAttribute("role", "radiogroup");
  seg.setAttribute("aria-label", row.label);
  const btns: HTMLButtonElement[] = [];
  const refresh = (): void => {
    const cur = row.get();
    for (const b of btns) {
      const on = b.dataset["value"] === cur;
      b.classList.toggle("active", on);
      b.setAttribute("aria-checked", on ? "true" : "false");
      b.tabIndex = on ? 0 : -1; // roving tabindex within the radiogroup
    }
  };
  row.options.forEach((o, i) => {
    const b = el("button", "cw-set-seg-btn", o.label);
    b.type = "button";
    b.setAttribute("role", "radio");
    b.dataset["value"] = o.value;
    b.addEventListener("click", () => { row.set(o.value); refresh(); });
    // Arrow-key navigation within the group (radiogroup convention).
    b.addEventListener("keydown", (e) => {
      let d = 0;
      if (e.key === "ArrowRight" || e.key === "ArrowDown") d = 1;
      else if (e.key === "ArrowLeft" || e.key === "ArrowUp") d = -1;
      else return;
      e.preventDefault();
      const next = btns[(i + d + btns.length) % btns.length]!;
      row.set(next.dataset["value"]!);
      refresh();
      next.focus();
    });
    btns.push(b);
    seg.append(b);
  });
  wrap.append(seg);
  refresh();
  return { wrap, refresh };
}

/** Show the Settings dialog. Non-blocking; every control applies live. */
export function showSettingsDialog(o: SettingsDialogOptions): void {
  const shell = createDialogShell({
    prefix: "cw-set",
    cssId: "cw-set-styles",
    css: CSS,
    title: "Settings",
    ...(o.onClose ? { onClose: o.onClose } : {}),
  });

  for (const group of o.groups) {
    if (group.rows.length === 0) continue;
    const g = el("div", "cw-set-group");
    g.append(el("div", "cw-set-group-ttl", group.title));
    for (const row of group.rows) g.append(segmented(row).wrap);
    shell.body.append(g);
  }

  const done = el("button");
  done.type = "button";
  done.textContent = "Done";
  done.addEventListener("click", () => shell.close());
  shell.foot.append(done);
  setTimeout(() => done.focus(), 0);
}
