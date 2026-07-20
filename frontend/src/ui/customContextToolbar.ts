// Adapts an embedder-registered `ContextToolbarSpec` (public `contextToolbars`
// option) into a `ContextToolbar` the manager can coordinate. The buttons are the
// same custom-button shape used by the ribbon and the floating format bar.

import { injectCssOnce } from "./styles";
import { createFloatingBar, type ContextToolbar } from "./contextToolbar";
import type { AnchorRect } from "./floatingBarPosition";
import type { ContextToolbarSpec, FloatingToolbarButtonSpec, ToolbarContext } from "../config";

export interface CustomContextToolbarDeps {
  /** Build the live ToolbarContext (host wires the editor accessors). */
  context(): ToolbarContext;
  /** Invoke a button's onClick via the ribbon action context (guarded). */
  runButton(button: FloatingToolbarButtonSpec): void;
}

const CSS = `
.cw-ctxbar{position:fixed;display:none;align-items:center;gap:1px;z-index:41;
  background:#fff;border:1px solid #c8c6c4;border-radius:6px;padding:3px;
  box-shadow:0 3px 12px rgba(0,0,0,.18);font:13px Arial,sans-serif;color:#323130;
  user-select:none;white-space:nowrap;}
.cw-ctxbar button{height:26px;min-width:26px;display:inline-flex;align-items:center;
  justify-content:center;gap:4px;padding:0 8px;border:1px solid transparent;border-radius:4px;
  background:transparent;cursor:pointer;color:#323130;font:13px Arial,sans-serif;}
.cw-ctxbar button:hover{background:#e1dfdd;}
.cw-ctxbar button.active{background:#cfe2ff;border-color:#9ac2ff;}
@media (prefers-color-scheme: dark){
  :root[data-theme="dark"] .cw-ctxbar{background:#2b2b2b;border-color:#4a4a4a;color:#e6e6e6;}
  :root[data-theme="dark"] .cw-ctxbar button{color:#e6e6e6;}
  :root[data-theme="dark"] .cw-ctxbar button:hover{background:#3c3c3c;}
  :root[data-theme="dark"] .cw-ctxbar button.active{background:#264b73;border-color:#3d6ba5;}
}`;

/** Build a ContextToolbar from an embedder spec (priority defaults to 15 — below the
 *  built-in image/hyperlink/text bars). */
export function createCustomContextToolbar(
  spec: ContextToolbarSpec,
  deps: CustomContextToolbarDeps,
): ContextToolbar {
  injectCssOnce("cw-ctxbar-styles", CSS);
  const fb = createFloatingBar({ className: "cw-ctxbar", ariaLabel: spec.id });

  const refreshers: ((ctx: ToolbarContext) => void)[] = [];
  for (const button of spec.buttons) {
    const b = document.createElement("button");
    b.type = "button";
    b.title = button.tooltip ?? button.label ?? button.id;
    if (button.icon) b.innerHTML = button.icon;
    else b.textContent = button.label ?? button.id;
    b.addEventListener("click", () => deps.runButton(button));
    const { active } = button;
    if (active) {
      b.setAttribute("aria-pressed", "false");
      refreshers.push((ctx) => {
        const on = active(ctx.format);
        b.classList.toggle("active", on);
        b.setAttribute("aria-pressed", String(on));
      });
    }
    fb.el.appendChild(b);
  }

  return {
    id: spec.id,
    priority: spec.priority ?? 15,
    resolve: (): AnchorRect | null => {
      const ctx = deps.context();
      if (!spec.when(ctx)) return null;
      return spec.anchor ? spec.anchor(ctx) : ctx.selectionRect();
    },
    show: (anchor: AnchorRect): void => {
      const ctx = deps.context();
      for (const update of refreshers) update(ctx);
      fb.place(anchor);
    },
    hide: () => fb.hide(),
    destroy: () => fb.destroy(),
  };
}
