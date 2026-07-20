// TOC context toolbar: appears at the caret inside a Table-of-Contents entry, with a
// one-click "Update table of contents". Collapsed-caret only.

import { createFloatingBar, injectCtxBarCss, type ContextToolbar } from "./contextToolbar";
import type { AnchorRect } from "./floatingBarPosition";

export interface TocContextToolbarDeps {
  /** True when the caret is inside a TOC entry. */
  inToc(): boolean;
  /** True when a non-empty range is selected (defers to the format bar then). */
  hasRangeSelection(): boolean;
  anchorRect(): AnchorRect | null;
  /** Recompute the TOC (returns the number of entries updated). */
  update(): void;
}

/** The TOC bar as a ContextToolbar (priority 22). */
export function createTocContextToolbar(deps: TocContextToolbarDeps): ContextToolbar {
  injectCtxBarCss();
  const fb = createFloatingBar({ className: "cw-ctxbar", ariaLabel: "Table of contents" });

  const b = document.createElement("button");
  b.type = "button";
  b.textContent = "Update table of contents";
  b.title = "Refresh entries and page numbers";
  b.addEventListener("click", () => deps.update());
  fb.el.appendChild(b);

  return {
    id: "toc",
    priority: 22,
    resolve: (): AnchorRect | null => {
      if (deps.hasRangeSelection() || !deps.inToc()) return null;
      return deps.anchorRect();
    },
    show: (anchor: AnchorRect): void => fb.place(anchor),
    hide: () => fb.hide(),
    destroy: () => fb.destroy(),
  };
}
