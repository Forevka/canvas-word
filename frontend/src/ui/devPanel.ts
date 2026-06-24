// Develop-mode inspector — a browser-devtools-style floating panel over the editor.
// This file is the SHELL: a draggable/resizable window with a tab strip, a shared
// toolbar (filter + per-tab extras), a content area that swaps per tab, and a shared
// detail/JSON pane. Each tab lives in ./devpanel/*Tab.ts and is a `PanelTab`.
//
// Gated twice over: it only exists when the embedder passes `develop: true`, and even
// then runs nothing until the developer opens it from the Developer ribbon tab. The
// tree ↔ canvas bridge goes through the editor's setInspectorHighlight /
// revealInspectorTarget / onInspectorHover surface (see ../index.ts).
//
// Pattern mirrors styleManager.ts: injectCssOnce + makeFloatingDialog, an
// AbortController for teardown, Escape-to-close, a handle with close()/refresh().

import { makeFloatingDialog } from "./floatingDialog";
import { injectCssOnce } from "./styles";
import { createModelTab } from "./devpanel/modelTab";
import { el, type DevPanelEditor, type PanelCtx, type PanelTab } from "./devpanel/types";

export type { DevPanelEditor } from "./devpanel/types";

export interface DevPanelOptions {
  editor: DevPanelEditor;
  onClose?: () => void;
}

export interface DevPanelHandle {
  close(): void;
  /** Re-read state and refresh the active tab (preserving its UI state). */
  refresh(): void;
  /** Reverse sync: highlight the tree row for a canvas block id (tree tabs). */
  highlightNode(blockId: string | null): void;
}

const CSS = `
.cw-dev-backdrop{position:fixed;inset:0;z-index:1100;}
.cw-dev-modal{position:fixed;width:min(460px,96vw);height:min(680px,92vh);min-width:300px;min-height:240px;
  max-width:96vw;max-height:96vh;display:flex;flex-direction:column;background:#1e1f22;resize:both;
  border-radius:10px;box-shadow:0 18px 56px rgba(0,0,0,.5);font:12px/1.5 'Segoe UI',Arial,sans-serif;color:#d4d6da;overflow:hidden;}
.cw-dev-head{display:flex;align-items:center;gap:8px;padding:9px 12px;border-bottom:1px solid #34363b;background:#26282c;cursor:move;}
.cw-dev-head h2{margin:0;font-size:13px;font-weight:600;flex:1 1 auto;color:#e8eaed;}
.cw-dev-x{border:none;background:transparent;font-size:18px;line-height:1;color:#9aa0a6;cursor:pointer;width:26px;height:26px;border-radius:6px;}
.cw-dev-x:hover{background:#34363b;color:#fff;}
.cw-dev-tabs{display:flex;gap:2px;padding:4px 8px 0;border-bottom:1px solid #34363b;background:#26282c;}
.cw-dev-tab{border:none;background:transparent;color:#9aa0a6;cursor:pointer;font:inherit;padding:5px 10px;border-radius:6px 6px 0 0;border-bottom:2px solid transparent;}
.cw-dev-tab:hover{background:#2f3236;color:#d4d6da;}
.cw-dev-tab.active{color:#e8eaed;border-bottom-color:#5b9bd5;background:#1e1f22;}
.cw-dev-tab .cw-dev-tabcount{color:#6b6f76;font-size:10px;margin-left:5px;}
.cw-dev-toolbar{display:flex;align-items:center;gap:6px;padding:7px 10px;border-bottom:1px solid #34363b;background:#222428;flex-wrap:wrap;}
.cw-dev-filter{flex:1 1 140px;height:26px;padding:0 8px;border:1px solid #3a3d42;border-radius:6px;background:#1a1b1e;color:#e8eaed;font-size:12px;}
.cw-dev-filter::placeholder{color:#6b6f76;}
.cw-dev-content{flex:1 1 auto;display:flex;min-height:0;}
.cw-dev-tabpane{flex:1 1 auto;min-height:0;display:flex;flex-direction:column;}
.cw-dev-tree{flex:1 1 auto;overflow:auto;padding:4px 0;font-family:'Cascadia Code',Consolas,monospace;}
.cw-dev-row{display:flex;align-items:baseline;gap:5px;padding:1px 10px 1px 0;cursor:pointer;white-space:nowrap;}
.cw-dev-row:hover{background:#2b2d31;}
.cw-dev-row.sel{background:#33415e;}
.cw-dev-row.cur{outline:1px solid #5b9bd5;outline-offset:-1px;}
.cw-dev-row.match{background:#3a3320;}
.cw-dev-tw{display:inline-block;width:12px;flex:0 0 auto;text-align:center;color:#9aa0a6;user-select:none;}
.cw-dev-kind{color:#9cdcfe;flex:0 0 auto;}
.cw-dev-kind.tag{color:#c586c0;}
.cw-dev-kind.run{color:#7ec699;}
.cw-dev-kind.group{color:#d7ba7d;}
.cw-dev-kind.layout{color:#80cbc4;}
.cw-dev-kind.problem{color:#e0e0e0;}
.cw-dev-prev{color:#cea36a;overflow:hidden;text-overflow:ellipsis;}
.cw-dev-badge{color:#6b9bd1;font-size:10px;border:1px solid #3f5470;border-radius:7px;padding:0 5px;flex:0 0 auto;}
.cw-dev-id{color:#6b6f76;font-size:10px;flex:0 0 auto;}
.cw-dev-empty{color:#6b6f76;padding:6px 14px;font-style:italic;}
.cw-dev-chip{display:inline-flex;align-items:center;gap:4px;height:24px;padding:0 9px;border:1px solid #3a3d42;border-radius:12px;background:#1a1b1e;color:#9aa0a6;cursor:pointer;font-size:11px;user-select:none;}
.cw-dev-chip:hover{border-color:#4a4d52;color:#d4d6da;}
.cw-dev-chip.on{background:#2d4156;border-color:#5b9bd5;color:#cfe3f5;}
.cw-dev-btn{height:24px;padding:0 9px;border:1px solid #3a3d42;border-radius:6px;background:#1a1b1e;color:#9aa0a6;cursor:pointer;font-size:11px;}
.cw-dev-btn:hover{border-color:#4a4d52;color:#d4d6da;}
.cw-dev-detail{flex:0 0 200px;border-top:1px solid #34363b;background:#191a1d;display:flex;flex-direction:column;min-height:0;}
.cw-dev-detail-head{padding:6px 12px;border-bottom:1px solid #2b2d31;color:#9aa0a6;font-size:11px;text-transform:uppercase;letter-spacing:.05em;}
.cw-dev-json{flex:1 1 auto;overflow:auto;margin:0;padding:8px 12px;font-family:'Cascadia Code',Consolas,monospace;font-size:11px;color:#cdd1d6;white-space:pre;}`;

export function showDevPanel(opts: DevPanelOptions): DevPanelHandle {
  injectCssOnce("cw-dev-styles", CSS);
  const { editor } = opts;

  // ---- DOM scaffold ---------------------------------------------------------
  const backdrop = el("div", "cw-dev-backdrop");
  const modal = el("div", "cw-dev-modal");
  modal.addEventListener("mousedown", (e) => e.stopPropagation());

  const head = el("div", "cw-dev-head");
  const h2 = el("h2", undefined, "Document Inspector");
  const xBtn = el("button", "cw-dev-x", "×");
  head.append(h2, xBtn);

  const tabStrip = el("div", "cw-dev-tabs");

  const toolbar = el("div", "cw-dev-toolbar");
  const filterInput = el("input", "cw-dev-filter");
  filterInput.type = "search";
  filterInput.placeholder = "Filter…";
  const extrasHost = el("div", "cw-dev-extras");
  extrasHost.style.cssText = "display:flex;align-items:center;gap:6px;flex:1 1 auto;flex-wrap:wrap;";
  toolbar.append(filterInput, extrasHost);

  const content = el("div", "cw-dev-content");

  const detail = el("div", "cw-dev-detail");
  const detailHead = el("div", "cw-dev-detail-head", "Select a node");
  const json = el("pre", "cw-dev-json");
  detail.append(detailHead, json);

  modal.append(head, tabStrip, toolbar, content, detail);
  backdrop.append(modal);
  document.body.append(backdrop);

  // ---- Shared detail pane ---------------------------------------------------
  const showDetail = (title: string, data: unknown): void => {
    detailHead.textContent = title;
    try {
      json.textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
    } catch {
      json.textContent = String(data);
    }
  };

  // ---- Tabs -----------------------------------------------------------------
  const ctx: PanelCtx = { editor, showDetail };
  const tabs: PanelTab[] = [createModelTab(ctx)];
  let active: PanelTab = tabs[0]!;
  const tabButtons = new Map<string, HTMLButtonElement>();

  for (const tab of tabs) {
    tab.element.classList.add("cw-dev-tabpane");
    tab.element.style.display = "none";
    content.append(tab.element);
    if (tab.toolbar) { tab.toolbar.style.display = "none"; extrasHost.append(tab.toolbar); }
    const btn = el("button", "cw-dev-tab", tab.label);
    btn.addEventListener("click", () => switchTo(tab.id));
    tabButtons.set(tab.id, btn);
    tabStrip.append(btn);
  }

  const switchTo = (id: string): void => {
    const next = tabs.find((t) => t.id === id);
    if (!next) return;
    active.element.style.display = "none";
    if (active.toolbar) active.toolbar.style.display = "none";
    tabButtons.get(active.id)?.classList.remove("active");
    active = next;
    active.element.style.display = "flex";
    if (active.toolbar) active.toolbar.style.display = "flex";
    tabButtons.get(active.id)?.classList.add("active");
    filterInput.style.display = active.usesFilter ? "" : "none";
    if (active.usesFilter) active.setFilter?.(filterInput.value);
    active.activate();
  };

  filterInput.addEventListener("input", () => active.usesFilter && active.setFilter?.(filterInput.value));

  // ---- Lifecycle ------------------------------------------------------------
  const ac = new AbortController();
  editor.setInspectorActive(true);
  const handle: DevPanelHandle = {
    close(): void {
      editor.setInspectorActive(false);
      editor.setInspectorHighlight(null);
      for (const t of tabs) t.destroy?.();
      backdrop.remove();
      ac.abort();
      opts.onClose?.();
    },
    refresh(): void { active.refresh(); },
    highlightNode(blockId: string | null): void { active.highlightByBlockId?.(blockId); },
  };
  window.addEventListener("keydown", (ev: KeyboardEvent) => {
    if (ev.key === "Escape") { ev.preventDefault(); ev.stopPropagation(); handle.close(); }
  }, { capture: true, signal: ac.signal });
  makeFloatingDialog({ backdrop, modal, handle: head, signal: ac.signal, noDrag: ".cw-dev-x" });
  xBtn.addEventListener("click", () => handle.close());

  switchTo(active.id);
  return handle;
}
