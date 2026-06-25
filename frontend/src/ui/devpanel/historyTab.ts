// History tab — the recorded edit log (getChangeLog), newest at the bottom. Each
// row summarizes a change (origin + op kinds + time); click to inspect its ops.
// Undo / Redo buttons step the document (recorded as literal forward changes, so
// the log grows — there is no hidden pointer to scrub, hence step buttons).

import type { Change, Op } from "@cw/shared";
import { el, type PanelCtx, type PanelTab } from "./types";

const opSummary = (ops: Op[]): string => {
  const counts: Record<string, number> = {};
  for (const o of ops) counts[o.type] = (counts[o.type] ?? 0) + 1;
  return Object.entries(counts).map(([t, n]) => (n > 1 ? `${t}×${n}` : t)).join(", ");
};

const timeOf = (ts: number): string => {
  try {
    return new Date(ts).toLocaleTimeString();
  } catch {
    return "";
  }
};

export function createHistoryTab(ctx: PanelCtx): PanelTab {
  const list = el("div", "cw-dev-history");

  const toolbar = el("div");
  toolbar.style.cssText = "display:flex;align-items:center;gap:6px;";
  const undoBtn = el("button", "cw-dev-btn", "↶ Undo");
  const redoBtn = el("button", "cw-dev-btn", "↷ Redo");
  const headLbl = el("span", "cw-dev-overlabel");
  undoBtn.addEventListener("click", () => { ctx.editor.undo(); render(true); });
  redoBtn.addEventListener("click", () => { ctx.editor.redo(); render(true); });
  toolbar.append(undoBtn, redoBtn, headLbl);

  let lastLen = -1;
  let selectedIdx = -1;

  const render = (force = false): void => {
    const log = ctx.editor.getChangeLog();
    headLbl.textContent = `${log.length} changes · head ${ctx.editor.getChangeHead()}`;
    if (!force && log.length === lastLen) return;
    const atBottom = list.scrollTop + list.clientHeight >= list.scrollHeight - 8;
    lastLen = log.length;
    list.replaceChildren();
    if (log.length === 0) {
      list.append(el("div", "cw-dev-empty", "No edits recorded yet. Type or run a command to see changes here."));
      return;
    }
    log.forEach((c: Change, i) => {
      const row = el("div", "cw-dev-hrow");
      if (i === selectedIdx) row.classList.add("sel");
      row.append(
        el("span", "cw-dev-hi", `#${i}`),
        el("span", `cw-dev-horigin ${c.origin}`, c.origin),
        el("span", "cw-dev-hops", opSummary(c.ops)),
        el("span", "cw-dev-hts", timeOf(c.ts)),
      );
      row.addEventListener("click", () => {
        selectedIdx = i;
        for (const r of list.querySelectorAll(".cw-dev-hrow.sel")) r.classList.remove("sel");
        row.classList.add("sel");
        ctx.showDetail(`change #${i} · ${c.origin}`, c);
      });
      list.append(row);
    });
    if (atBottom) list.scrollTop = list.scrollHeight; // follow the tail
  };

  return {
    id: "history",
    label: "History",
    element: list,
    toolbar,
    usesFilter: false,
    activate: () => render(true),
    refresh: () => render(),
    count: () => ctx.editor.getChangeLog().length || null,
    destroy: () => list.replaceChildren(),
  };
}
