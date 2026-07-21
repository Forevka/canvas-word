// Hyperlink context toolbar (Word's link popup): shows the URL plus Open / Edit /
// Copy / Remove when the caret sits inside a hyperlink. Only shows at a *collapsed*
// caret — a text range defers to the format bar (higher-value at that moment). One
// of the built-in ContextToolbars coordinated by the shared manager.

import { injectCssOnce } from "./styles";
import { createFloatingBar, type ContextToolbar } from "./contextToolbar";
import type { AnchorRect } from "./floatingBarPosition";

export interface LinkContextToolbarActions {
  /** Open the URL (typically a new tab). */
  open(url: string): void;
  /** Edit the link — `anchor` is the Edit button, to position the link dialog. */
  edit(anchor: HTMLElement, url: string): void;
  /** Copy the URL to the clipboard. */
  copy(url: string): void;
  /** Remove the hyperlink from the run(s). */
  remove(): void;
}

export interface LinkContextToolbarDeps {
  /** The hyperlink URL at the caret, or null when the caret isn't on a link. */
  linkUrl(): string | null;
  /** True when a non-empty range is selected — the link bar defers to the format
   *  bar then, so it only appears at a collapsed caret inside a link. */
  hasRangeSelection(): boolean;
  /** Anchor rect (caret/selection start line). */
  anchorRect(): AnchorRect | null;
  actions: LinkContextToolbarActions;
}

const CSS = `
.cw-linkbar{position:fixed;display:none;align-items:center;gap:1px;z-index:41;
  background:#fff;border:1px solid #c8c6c4;border-radius:6px;padding:3px;
  box-shadow:0 3px 12px rgba(0,0,0,.18);font:13px Arial,sans-serif;color:#323130;
  user-select:none;white-space:nowrap;}
.cw-linkbar button{height:26px;min-width:26px;display:inline-flex;align-items:center;
  justify-content:center;padding:0 8px;border:1px solid transparent;border-radius:4px;
  background:transparent;cursor:pointer;color:#323130;font:13px Arial,sans-serif;}
.cw-linkbar button:hover{background:#e1dfdd;}
.cw-linkbar button.danger:hover{background:#fde7e9;color:#a4262c;}
.cw-linkbar .url{max-width:200px;overflow:hidden;text-overflow:ellipsis;color:#0563c1;padding:0 6px;}
.cw-linkbar .sep{width:1px;height:18px;background:#e1dfdd;margin:0 3px;}
@media (prefers-color-scheme: dark){
  :root[data-theme="dark"] .cw-linkbar{background:#2b2b2b;border-color:#4a4a4a;color:#e6e6e6;}
  :root[data-theme="dark"] .cw-linkbar button{color:#e6e6e6;}
  :root[data-theme="dark"] .cw-linkbar button:hover{background:#3c3c3c;}
  :root[data-theme="dark"] .cw-linkbar .url{color:#6ca6ff;}
  :root[data-theme="dark"] .cw-linkbar .sep{background:#4a4a4a;}
}`;

/** The hyperlink bar as a ContextToolbar (priority 25 — above the text format bar,
 *  below the image bar). */
export function createLinkContextToolbar(deps: LinkContextToolbarDeps): ContextToolbar {
  injectCssOnce("cw-linkbar-styles", CSS);
  const fb = createFloatingBar({ className: "cw-linkbar", ariaLabel: "Hyperlink" });

  const urlLabel = document.createElement("span");
  urlLabel.className = "url";
  fb.el.appendChild(urlLabel);

  const sep = (): void => {
    const s = document.createElement("div");
    s.className = "sep";
    fb.el.appendChild(s);
  };
  const button = (label: string, title: string, cls: string, onClick: (b: HTMLButtonElement) => void): void => {
    const b = document.createElement("button");
    b.type = "button";
    if (cls) b.className = cls;
    b.textContent = label;
    b.title = title;
    // Read the URL live at click time (the caret may have moved since the last show).
    b.addEventListener("click", () => onClick(b));
    fb.el.appendChild(b);
  };
  // Actions read the live URL so a moved caret can't fire a stale link.
  const withUrl = (fn: (url: string) => void) => (): void => {
    const url = deps.linkUrl();
    if (url) fn(url);
  };

  sep();
  button("Open", "Open link in a new tab", "", withUrl((u) => deps.actions.open(u)));
  button("Edit", "Edit link", "", (b) => {
    const url = deps.linkUrl();
    if (url) deps.actions.edit(b, url);
  });
  button("Copy", "Copy link address", "", withUrl((u) => deps.actions.copy(u)));
  sep();
  button("Remove", "Remove link", "danger", () => deps.actions.remove());

  return {
    id: "hyperlink",
    priority: 25,
    resolve: (): AnchorRect | null => {
      if (deps.hasRangeSelection()) return null; // a range shows the format bar
      return deps.linkUrl() ? deps.anchorRect() : null;
    },
    show: (anchor, viewport) => {
      const url = deps.linkUrl() ?? "";
      urlLabel.textContent = url;
      urlLabel.title = url;
      fb.place(anchor, viewport);
    },
    hide: () => fb.hide(),
    destroy: () => fb.destroy(),
  };
}
