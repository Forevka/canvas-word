// Regression coverage for the keyboard-layout i18n bug: the built-in editing
// chords (Ctrl+Z/Y/B/I/U, Ctrl+Enter) matched on KeyboardEvent.key, which carries
// the LAYOUT-PRODUCED character — so on a Ukrainian layout the physical Z key emits
// 'я' and Undo was silently dead. This bug is invisible on a Latin layout, which is
// exactly why it survived, so these tests synthesise non-Latin events explicitly.

import { describe, it, expect, vi } from "vitest";
import { createKeymapHandler, type KeymapDeps } from "./keymap";

function setup() {
  const deps: KeymapDeps = { dispatch: vi.fn(), undo: vi.fn(), redo: vi.fn(), toggleStyle: vi.fn() };
  return { handle: createKeymapHandler(deps), deps };
}

/** A KeyboardEvent-shaped stub: `key` is the produced character, `code` the physical
 *  position. Only the fields the handler reads are provided. */
const kev = (
  key: string,
  code: string,
  mods: Partial<{ ctrlKey: boolean; metaKey: boolean; shiftKey: boolean; altKey: boolean }> = {},
) =>
  ({ key, code, ctrlKey: false, metaKey: false, shiftKey: false, altKey: false, preventDefault: vi.fn(), ...mods }) as unknown as KeyboardEvent;

describe("createKeymapHandler under non-Latin keyboard layouts", () => {
  it("Ctrl+Z undoes even though the physical Z emits 'я' (Ukrainian layout)", () => {
    const { handle, deps } = setup();
    const e = kev("я", "KeyZ", { ctrlKey: true });
    handle(e);
    expect(deps.undo).toHaveBeenCalledTimes(1);
    expect(deps.redo).not.toHaveBeenCalled();
    expect(e.preventDefault).toHaveBeenCalled();
  });

  it("Ctrl+Shift+Z redoes on a Cyrillic layout", () => {
    const { handle, deps } = setup();
    handle(kev("я", "KeyZ", { ctrlKey: true, shiftKey: true }));
    expect(deps.redo).toHaveBeenCalledTimes(1);
  });

  it("Ctrl+Y redoes when the physical Y emits 'н'", () => {
    const { handle, deps } = setup();
    handle(kev("н", "KeyY", { ctrlKey: true }));
    expect(deps.redo).toHaveBeenCalledTimes(1);
  });

  it("Ctrl+B/I/U toggle bold/italic/underline via physical position", () => {
    const { handle, deps } = setup();
    handle(kev("и", "KeyB", { ctrlKey: true }));
    handle(kev("ш", "KeyI", { ctrlKey: true }));
    handle(kev("г", "KeyU", { ctrlKey: true }));
    expect(deps.toggleStyle).toHaveBeenNthCalledWith(1, "bold");
    expect(deps.toggleStyle).toHaveBeenNthCalledWith(2, "italic");
    expect(deps.toggleStyle).toHaveBeenNthCalledWith(3, "underline");
  });

  it("a Dvorak Ctrl+Z (produced 'z' on the physical ';' key) still undoes via the key path", () => {
    const { handle, deps } = setup();
    handle(kev("z", "KeySemicolon", { ctrlKey: true }));
    expect(deps.undo).toHaveBeenCalledTimes(1);
  });

  it("a plain Latin Ctrl+Z still undoes (no regression)", () => {
    const { handle, deps } = setup();
    handle(kev("z", "KeyZ", { ctrlKey: true }));
    expect(deps.undo).toHaveBeenCalledTimes(1);
  });

  it("does nothing without the Ctrl/Meta modifier", () => {
    const { handle, deps } = setup();
    handle(kev("я", "KeyZ")); // no modifier — a literal keystroke
    expect(deps.undo).not.toHaveBeenCalled();
  });
});
