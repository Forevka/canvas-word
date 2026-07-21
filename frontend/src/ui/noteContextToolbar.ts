// Footnote/endnote context toolbar: appears at the caret when it sits on a note's
// superscript reference marker — "Go to note" jumps into the note body, "Delete"
// removes the marker + body and renumbers the rest. Collapsed-caret only (a text
// range shows the format bar instead).

import { createFloatingBar, injectCtxBarCss, type ContextToolbar } from "./contextToolbar";
import type { AnchorRect } from "./floatingBarPosition";

/** The note under the caret, as reported by the host. */
export interface NoteCaretInfo {
  kind: "footnote" | "endnote";
  id: string;
}

export interface NoteContextToolbarActions {
  /** Reveal the note body and drop the caret into it. */
  goTo(info: NoteCaretInfo): void;
  /** Delete the note (marker + body) and renumber the remaining notes. */
  remove(info: NoteCaretInfo): void;
}

export interface NoteContextToolbarDeps {
  /** The note whose marker contains the caret, or null. */
  note(): NoteCaretInfo | null;
  /** True when a non-empty range is selected (the bar defers to the format bar then). */
  hasRangeSelection(): boolean;
  /** Caret anchor rect. */
  anchorRect(): AnchorRect | null;
  actions: NoteContextToolbarActions;
}

/** The note bar as a ContextToolbar (priority 24 — below the review bar, above TOC). */
export function createNoteContextToolbar(deps: NoteContextToolbarDeps): ContextToolbar {
  injectCtxBarCss();
  const fb = createFloatingBar({ className: "cw-ctxbar cw-notebar", ariaLabel: "Footnote" });

  let current: NoteCaretInfo | null = null;
  const button = (label: string, cls: string, onClick: () => void): HTMLButtonElement => {
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = label;
    if (cls) b.className = cls;
    b.addEventListener("click", () => {
      if (current) onClick();
    });
    fb.el.appendChild(b);
    return b;
  };

  const goToBtn = button("Go to note", "", () => deps.actions.goTo(current!));
  const deleteBtn = button("Delete", "danger", () => deps.actions.remove(current!));

  return {
    id: "note",
    priority: 24,
    resolve: (): AnchorRect | null => {
      if (deps.hasRangeSelection()) return null;
      return deps.note() ? deps.anchorRect() : null;
    },
    show: (anchor: AnchorRect): void => {
      current = deps.note();
      const noun = current?.kind === "endnote" ? "endnote" : "footnote";
      // Keep the a11y label + button titles in sync with the marker kind, so a
      // screen reader announces "Endnote" (not "Footnote") on an endnote marker.
      fb.el.setAttribute("aria-label", noun === "endnote" ? "Endnote" : "Footnote");
      goToBtn.title = `Go to this ${noun}`;
      deleteBtn.title = `Delete this ${noun}`;
      fb.place(anchor);
    },
    hide: () => fb.hide(),
    destroy: () => fb.destroy(),
  };
}
