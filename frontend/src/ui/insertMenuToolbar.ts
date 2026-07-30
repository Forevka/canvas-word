// Empty-paragraph insert menu (Notion-style "＋"): appears at the caret on an empty
// line and opens a menu of block inserts (heading, list, table, page break, TOC,
// footnote). Collapsed-caret only. The menu itself is built by the host (it owns the
// insert commands) and shown via showContextMenu.

import { createFloatingBar, injectCtxBarCss, type ContextToolbar } from "./contextToolbar";
import type { AnchorRect } from "./floatingBarPosition";

export interface InsertMenuToolbarDeps {
  /** True when the caret is collapsed in an empty paragraph. */
  onEmptyParagraph(): boolean;
  anchorRect(): AnchorRect | null;
  /** Open the insert menu anchored to `anchor` (the ＋ button). */
  openMenu(anchor: HTMLElement): void;
}

/** The empty-paragraph insert menu as a ContextToolbar (priority 16 — the lowest of
 *  the built-ins, so any real content context wins). */
export function createInsertMenuToolbar(deps: InsertMenuToolbarDeps): ContextToolbar {
  injectCtxBarCss();
  // Gutter placement so the chip sits in the left margin beside the empty line,
  // not centered above it over the previous line's text (critique L4).
  const fb = createFloatingBar({ className: "cw-ctxbar", ariaLabel: "Insert", place: { mode: "gutter", gap: 6 } });

  const b = document.createElement("button");
  b.type = "button";
  b.textContent = "＋ Insert";
  b.title = "Insert a block";
  b.addEventListener("click", () => deps.openMenu(b));
  fb.el.appendChild(b);

  return {
    id: "insert-menu",
    priority: 16,
    resolve: (): AnchorRect | null => (deps.onEmptyParagraph() ? deps.anchorRect() : null),
    show: (anchor, viewport) => fb.place(anchor, viewport),
    hide: () => fb.hide(),
    destroy: () => fb.destroy(),
  };
}
