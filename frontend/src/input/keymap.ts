// Editing shortcuts -> commands. Navigation keys live in selectionController;
// this map handles the modifier-chord editing surface. Both listen on the
// container, so keys bubbling out of the IME proxy land here too.

import type { Command } from "../editor/state";
import { insertColumnBreak, insertPageBreak } from "../editor/commands";

export type StyleKey = "bold" | "italic" | "underline" | "strikethrough" | "caps" | "smallCaps" | "doubleStrikethrough";

export interface KeymapDeps {
  dispatch(cmd: Command): void;
  undo(): void;
  redo(): void;
  /** Range -> restyle ops; collapsed caret -> pending typing style. */
  toggleStyle(key: StyleKey): void;
}

/** Declarative twin of the chords handled below — the SINGLE source the shortcuts
 *  cheat sheet reads (critique A3), co-located with the handler so the two can't
 *  drift out of sync. Edit both together. `Ctrl` renders as `⌘` on macOS in the UI. */
export const KEYMAP_ENTRIES: { chord: string; label: string }[] = [
  { chord: "Ctrl+B", label: "Bold" },
  { chord: "Ctrl+I", label: "Italic" },
  { chord: "Ctrl+U", label: "Underline" },
  { chord: "Ctrl+Z", label: "Undo" },
  { chord: "Ctrl+Shift+Z", label: "Redo" },
  { chord: "Ctrl+Y", label: "Redo" },
  { chord: "Ctrl+Enter", label: "Page break" },
  { chord: "Ctrl+Shift+Enter", label: "Column break" },
];

export function createKeymapHandler(deps: KeymapDeps): (ev: KeyboardEvent) => void {
  return (ev: KeyboardEvent): void => {
    const ctrl = ev.ctrlKey || ev.metaKey;
    if (!ctrl) return;
    if (ev.key === "Enter") {
      // Ctrl+Enter = page break; Ctrl+Shift+Enter = column break (Word).
      deps.dispatch(ev.shiftKey ? insertColumnBreak() : insertPageBreak());
      ev.preventDefault(); // also stops the proxy's beforeinput insertParagraph/LineBreak
      return;
    }
    switch (ev.key.toLowerCase()) {
      case "z":
        ev.shiftKey ? deps.redo() : deps.undo();
        ev.preventDefault();
        return;
      case "y":
        deps.redo();
        ev.preventDefault();
        return;
      case "b":
        deps.toggleStyle("bold");
        ev.preventDefault();
        return;
      case "i":
        deps.toggleStyle("italic");
        ev.preventDefault();
        return;
      case "u":
        deps.toggleStyle("underline");
        ev.preventDefault();
        return;
    }
  };
}
