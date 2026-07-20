// Contextual floating-toolbar framework. Generalizes the two hand-wired floating
// mini-toolbars (image + text selection) into one priority-based manager: it shows
// the single most-relevant bar for whatever the caret/selection is currently on,
// and hides the rest. Each context (image, hyperlink, text, or an embedder's own)
// is a `ContextToolbar` — a small object that reports an anchor rect when its
// context is active and renders itself there.

import { anchorInView, placeSelectionBar, type AnchorRect } from "./floatingBarPosition";
import { injectCssOnce } from "./styles";

// Shared look for the built-in context bars (equation / review / TOC / list / insert
// menu). Apply the `cw-ctxbar` class alongside a bar-specific one and call
// `injectCtxBarCss()` once; per-bar CSS then only needs overrides.
const CTX_BAR_CSS = `
.cw-ctxbar{position:fixed;display:none;align-items:center;gap:2px;z-index:41;
  background:#fff;border:1px solid #c8c6c4;border-radius:6px;padding:3px;
  box-shadow:0 3px 12px rgba(0,0,0,.18);font:13px Arial,sans-serif;color:#323130;
  user-select:none;white-space:nowrap;}
.cw-ctxbar button{height:26px;min-width:26px;display:inline-flex;align-items:center;
  justify-content:center;gap:4px;padding:0 8px;border:1px solid transparent;border-radius:4px;
  background:transparent;cursor:pointer;color:#323130;font:13px Arial,sans-serif;}
.cw-ctxbar button:hover{background:#e1dfdd;}
.cw-ctxbar button.active{background:#cfe2ff;border-color:#9ac2ff;}
.cw-ctxbar button.danger:hover{background:#fde7e9;color:#a4262c;}
.cw-ctxbar button svg{width:16px;height:16px;}
.cw-ctxbar .sep{width:1px;height:18px;background:#e1dfdd;margin:0 3px;flex:0 0 auto;}
@media (prefers-color-scheme: dark){
  :root[data-theme="dark"] .cw-ctxbar{background:#2b2b2b;border-color:#4a4a4a;color:#e6e6e6;}
  :root[data-theme="dark"] .cw-ctxbar button{color:#e6e6e6;}
  :root[data-theme="dark"] .cw-ctxbar button:hover{background:#3c3c3c;}
  :root[data-theme="dark"] .cw-ctxbar button.active{background:#264b73;border-color:#3d6ba5;}
  :root[data-theme="dark"] .cw-ctxbar .sep{background:#4a4a4a;}
}`;

/** Inject the shared context-bar stylesheet once (idempotent). */
export function injectCtxBarCss(): void {
  injectCssOnce("cw-ctxbar-shared", CTX_BAR_CSS);
}

/** A body-level `position:fixed` bar: the shared DOM + placement/visibility the
 *  image and text bars both need. Content is built by the caller into `.el`. */
export interface FloatingBar {
  /** The bar element (append your buttons here). */
  readonly el: HTMLDivElement;
  /** Position at `anchor` (flip/clamp via placeSelectionBar) and show. Call AFTER
   *  syncing content, since placement measures the bar. */
  place(anchor: AnchorRect): void;
  hide(): void;
  visible(): boolean;
  destroy(): void;
}

export interface FloatingBarOptions {
  /** Class applied to the bar element (the caller injects its matching CSS). */
  className: string;
  /** aria-label; also sets role="toolbar" when present. */
  ariaLabel?: string;
  /** Placement overrides passed to placeSelectionBar (gap/margin/topGuard). */
  place?: { gap?: number; margin?: number; topGuard?: number };
}

export function createFloatingBar(opts: FloatingBarOptions): FloatingBar {
  const el = document.createElement("div");
  el.className = opts.className;
  el.style.display = "none";
  if (opts.ariaLabel) {
    el.setAttribute("role", "toolbar");
    el.setAttribute("aria-label", opts.ariaLabel);
  }
  // Never steal the selection: act on click/mouseup, suppress the mousedown.
  el.addEventListener("mousedown", (e) => e.preventDefault());
  document.body.appendChild(el);

  let shown = false;
  return {
    el,
    place(anchor) {
      el.style.display = "flex";
      shown = true;
      const { left, top } = placeSelectionBar(
        anchor,
        { width: el.offsetWidth, height: el.offsetHeight },
        { width: window.innerWidth, height: window.innerHeight },
        opts.place ?? {},
      );
      el.style.left = `${left}px`;
      el.style.top = `${top}px`;
    },
    hide() {
      if (!shown) return;
      shown = false;
      el.style.display = "none";
    },
    visible: () => shown,
    destroy() {
      el.remove();
    },
  };
}

/** One context's floating bar. The manager coordinates many of these, showing only
 *  the highest-priority one whose context is active. */
export interface ContextToolbar {
  /** Stable id (for debugging / dedupe). */
  id: string;
  /** Higher wins when several contexts are active at once. */
  priority: number;
  /** The context detector + anchor: return the viewport anchor rect when this
   *  toolbar's context is active right now, else null. Should be side-effect-free. */
  resolve(): AnchorRect | null;
  /** Sync button state for the active context, then place + show at `anchor`. */
  show(anchor: AnchorRect): void;
  hide(): void;
  destroy(): void;
}

export interface ActivePick {
  toolbar: ContextToolbar;
  anchor: AnchorRect;
}

/** Pure winner-selection: the highest-priority toolbar whose `resolve()` returns a
 *  non-null anchor that is still within the viewport. Ties keep registration order
 *  (stable sort). Exported for unit testing — no DOM access. */
export function pickActive(
  toolbars: readonly ContextToolbar[],
  viewport: { width: number; height: number },
): ActivePick | null {
  const byPriority = [...toolbars].sort((a, b) => b.priority - a.priority);
  for (const toolbar of byPriority) {
    const anchor = toolbar.resolve();
    if (anchor && anchorInView(anchor, viewport)) return { toolbar, anchor };
  }
  return null;
}

export interface ContextToolbarManager {
  /** Add a toolbar to the coordination set. */
  register(toolbar: ContextToolbar): void;
  /** Re-evaluate which toolbar (if any) should be visible, and where. Cheap —
   *  call on every change / scroll / zoom / resize. */
  refresh(): void;
  /** Hide whatever is currently shown (e.g. on Escape). */
  hideActive(): void;
  /** Remove every toolbar's DOM. */
  destroy(): void;
}

export interface ContextToolbarManagerDeps {
  /** The scroll container to reposition against (the editor's page canvas). */
  scrollEl: HTMLElement;
  /** Aborted on editor teardown — detaches listeners + destroys the toolbars. */
  signal: AbortSignal;
  /** Optional global suppression (e.g. a modal owns the screen) — hides all bars. */
  suppressed?: () => boolean;
}

export function createContextToolbarManager(deps: ContextToolbarManagerDeps): ContextToolbarManager {
  const toolbars: ContextToolbar[] = [];
  let active: ContextToolbar | null = null;

  const refresh = (): void => {
    const viewport = { width: window.innerWidth, height: window.innerHeight };
    const pick = deps.suppressed?.() ? null : pickActive(toolbars, viewport);
    for (const t of toolbars) if (t !== pick?.toolbar) t.hide();
    active = pick?.toolbar ?? null;
    if (pick) pick.toolbar.show(pick.anchor);
  };
  const hideActive = (): void => {
    active?.hide();
    active = null;
  };
  const destroy = (): void => {
    for (const t of toolbars) t.destroy();
  };

  // Coalesce the high-frequency scroll/resize streams to one refresh per frame —
  // refresh reads layout (getBoundingClientRect + offsetWidth/Height), so running it
  // per scroll tick thrashes reflow (the editor throttles its hover/pinch paths the
  // same way). onChange/onZoomChange still call refresh() synchronously for immediacy.
  let rafId: number | null = null;
  const scheduleRefresh = (): void => {
    if (rafId !== null) return;
    rafId = requestAnimationFrame(() => {
      rafId = null;
      refresh();
    });
  };

  deps.scrollEl.addEventListener("scroll", scheduleRefresh, { signal: deps.signal });
  window.addEventListener("resize", scheduleRefresh, { signal: deps.signal });
  window.addEventListener(
    "keydown",
    (e) => {
      if (e.key === "Escape") hideActive();
    },
    { signal: deps.signal },
  );
  deps.signal.addEventListener(
    "abort",
    () => {
      if (rafId !== null) cancelAnimationFrame(rafId);
      destroy();
    },
    { once: true },
  );

  return {
    register(toolbar) {
      toolbars.push(toolbar);
      toolbars.sort((a, b) => b.priority - a.priority);
    },
    refresh,
    hideActive,
    destroy,
  };
}
