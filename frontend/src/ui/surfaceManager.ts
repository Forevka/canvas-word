// One arbiter for the editor's overlapping "surfaces" — floating dialogs (Font,
// Paragraph, Page Layout, Style Manager, …) and the find/replace bar. Before this
// each dialog installed its own capture-phase Escape and each surface picked its
// own z-index, so they stacked and occluded each other with no coordination:
// opening Manage Styles, then Ctrl+F (the find bar appeared BEHIND it at z-index
// 10), then Review left three surfaces mutually hiding, and Escape's behavior was
// ambiguous (critique L1). This module owns z-order, exclusivity, and a single
// Escape → close-the-topmost, mirroring how ui/contextToolbar arbitrates the
// floating bars.

export interface ManagedSurface {
  /** The positioned root whose z-index the manager controls (a dialog backdrop
   *  or the find bar). */
  el: HTMLElement;
  /** Close + tear the surface down. Must be idempotent — the manager may call it
   *  during exclusivity eviction or an Escape. */
  close: () => void;
  /** "dialog" surfaces are mutually exclusive: opening one evicts any other open
   *  dialog. "overlay" surfaces (the find bar) coexist and simply stack on top. */
  kind: "dialog" | "overlay";
}

export interface SurfaceHandle {
  /** Remove this surface from the stack (call from the surface's own close()). */
  release(): void;
}

// Above context menus (1300); below the notice/organize-pages max-z overlays.
const BASE_Z = 1400;
const STEP = 10;

const stack: ManagedSurface[] = [];
let escapeInstalled = false;

/** Re-assign z-index by stack position so the most-recently-opened surface is on
 *  top and Escape/close always targets what the user last summoned. */
function restack(): void {
  stack.forEach((s, i) => {
    s.el.style.zIndex = String(BASE_Z + i * STEP);
  });
}

function ensureEscape(): void {
  if (escapeInstalled) return;
  escapeInstalled = true;
  // Capture-phase so it wins over the editor's own key handling, exactly as the
  // per-dialog handlers used to. Only the TOPMOST surface closes per press.
  window.addEventListener(
    "keydown",
    (e: KeyboardEvent) => {
      const top = stack[stack.length - 1];
      if (e.key !== "Escape" || !top) return;
      e.preventDefault();
      e.stopPropagation();
      top.close();
    },
    { capture: true },
  );
}

/** Register a surface as open. Applies exclusivity + z-order and returns a handle
 *  the surface releases when it closes. */
export function openSurface(surface: ManagedSurface): SurfaceHandle {
  ensureEscape();
  if (surface.kind === "dialog") {
    // Only one dialog at a time — evict any others (iterate a copy; close()
    // releases them, mutating the stack).
    for (const other of [...stack]) if (other.kind === "dialog") other.close();
  }
  stack.push(surface);
  restack();
  let released = false;
  return {
    release(): void {
      if (released) return;
      released = true;
      const i = stack.indexOf(surface);
      if (i >= 0) {
        stack.splice(i, 1);
        restack();
      }
    },
  };
}

/** Close the topmost managed surface, if any. Returns whether one was closed —
 *  lets a caller (e.g. a global handler) fall through when nothing is open. */
export function closeTopSurface(): boolean {
  const top = stack[stack.length - 1];
  if (!top) return false;
  top.close();
  return true;
}

/** Whether any managed surface is currently open. */
export function hasOpenSurface(): boolean {
  return stack.length > 0;
}
