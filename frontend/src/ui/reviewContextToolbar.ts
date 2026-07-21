// Review context toolbar: appears at the caret when it's inside a tracked-change
// suggestion (Accept / Reject) or a comment thread (Reply / Resolve). Collapsed-caret
// only — a text range shows the format bar.

import { createFloatingBar, injectCtxBarCss, type ContextToolbar } from "./contextToolbar";
import type { AnchorRect } from "./floatingBarPosition";

/** The review item under the caret, as reported by the host. */
export interface ReviewCaretInfo {
  kind: "comment" | "suggestion";
  id: string;
  resolved?: boolean;
}

export interface ReviewContextToolbarActions {
  accept(id: string): void;
  reject(id: string): void;
  /** Scroll to / select the thread so the user can read + reply in the panel. */
  open(id: string): void;
  resolve(id: string, resolved: boolean): void;
}

export interface ReviewContextToolbarDeps {
  /** The review item whose range contains the caret, or null. */
  review(): ReviewCaretInfo | null;
  /** True when a non-empty range is selected (the bar defers to the format bar then). */
  hasRangeSelection(): boolean;
  /** Caret anchor rect. */
  anchorRect(): AnchorRect | null;
  actions: ReviewContextToolbarActions;
}

/** The review bar as a ContextToolbar (priority 26 — above the hyperlink bar). */
export function createReviewContextToolbar(deps: ReviewContextToolbarDeps): ContextToolbar {
  injectCtxBarCss();
  const fb = createFloatingBar({ className: "cw-ctxbar cw-reviewbar", ariaLabel: "Review" });

  let current: ReviewCaretInfo | null = null;
  const button = (label: string, title: string, cls: string, onClick: () => void): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = label;
    b.title = title;
    b.addEventListener("click", () => {
      if (current) onClick();
    });
    fb.el.appendChild(b);
    return b;
  };

  // Suggestion actions.
  const acceptBtn = button("Accept", "Accept this change", "", () => deps.actions.accept(current!.id));
  const rejectBtn = button("Reject", "Reject this change", "danger", () => deps.actions.reject(current!.id));
  // Comment actions.
  const replyBtn = button("Open", "Go to this comment thread", "", () => deps.actions.open(current!.id));
  const resolveBtn = button("Resolve", "Resolve this comment", "", () =>
    deps.actions.resolve(current!.id, !current!.resolved),
  );

  const setVisible = (b: HTMLButtonElement, on: boolean): void => {
    b.style.display = on ? "" : "none";
  };

  return {
    id: "review",
    priority: 26,
    resolve: (): AnchorRect | null => {
      if (deps.hasRangeSelection()) return null;
      return deps.review() ? deps.anchorRect() : null;
    },
    show: (anchor, viewport) => {
      current = deps.review();
      const isSuggestion = current?.kind === "suggestion";
      setVisible(acceptBtn, isSuggestion);
      setVisible(rejectBtn, isSuggestion);
      setVisible(replyBtn, !isSuggestion);
      setVisible(resolveBtn, !isSuggestion);
      resolveBtn.textContent = current?.resolved ? "Reopen" : "Resolve";
      fb.place(anchor, viewport);
    },
    hide: () => fb.hide(),
    destroy: () => fb.destroy(),
  };
}
