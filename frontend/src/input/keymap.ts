// Editing shortcuts -> commands. Navigation keys live in selectionController;
// this map handles the modifier-chord editing surface. Both listen on the
// container, so keys bubbling out of the IME proxy land here too.

import type { Command } from "../editor/state";
import { insertColumnBreak, insertPageBreak } from "../editor/commands";
import { keyMatches } from "../commands";

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
    // Alt must be absent: Windows reports AltGr as Ctrl+Alt, and the physical-code
    // fallback below would otherwise read AltGr+Z on a Polish layout (produces 'ż',
    // code 'KeyZ') as Ctrl+Z — swallowing a character the user meant to type.
    if (!ctrl || ev.altKey) return;
    // Match via keyMatches (hybrid key/physical-code) so these chords fire under
    // non-Latin keyboard layouts, where ev.key is the produced character (e.g. 'я'
    // for the physical Z), not the Latin letter.
    if (keyMatches("enter", ev)) {
      // Ctrl+Enter = page break; Ctrl+Shift+Enter = column break (Word).
      deps.dispatch(ev.shiftKey ? insertColumnBreak() : insertPageBreak());
      ev.preventDefault(); // also stops the proxy's beforeinput insertParagraph/LineBreak
      return;
    }
    if (keyMatches("z", ev)) {
      ev.shiftKey ? deps.redo() : deps.undo();
      ev.preventDefault();
      return;
    }
    if (keyMatches("y", ev)) {
      deps.redo();
      ev.preventDefault();
      return;
    }
    if (keyMatches("b", ev)) {
      deps.toggleStyle("bold");
      ev.preventDefault();
      return;
    }
    if (keyMatches("i", ev)) {
      deps.toggleStyle("italic");
      ev.preventDefault();
      return;
    }
    if (keyMatches("u", ev)) {
      deps.toggleStyle("underline");
      ev.preventDefault();
      return;
    }
  };
}
