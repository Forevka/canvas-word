import { describe, it, expect } from "vitest";
import { parseChord, chordMatches, keyMatches, normalizeChord, resolveCommandBindings, type EditorCommand, type KeyEventLike } from "./commands";

const ev = (key: string, mods: Partial<KeyEventLike> = {}): KeyEventLike => ({
  key,
  ctrlKey: false,
  altKey: false,
  shiftKey: false,
  metaKey: false,
  ...mods,
});

describe("parseChord", () => {
  it("parses modifiers + key, case/space-insensitive", () => {
    expect(parseChord("Mod+Shift+K")).toEqual({ key: "k", mod: true, ctrl: false, alt: false, shift: true, meta: false });
    expect(parseChord("  ctrl + alt + Enter ")).toEqual({ key: "enter", mod: false, ctrl: true, alt: true, shift: false, meta: false });
  });

  it("treats cmd/command/win/super as meta", () => {
    expect(parseChord("Cmd+P")?.meta).toBe(true);
    expect(parseChord("Super+P")?.meta).toBe(true);
  });

  it("returns null when there is no non-modifier key", () => {
    expect(parseChord("Ctrl+Shift")).toBeNull();
    expect(parseChord("")).toBeNull();
  });

  it("uses the last non-modifier segment as the key", () => {
    expect(parseChord("Mod+A+B")?.key).toBe("b");
  });
});

describe("chordMatches", () => {
  it("matches key + exact modifiers", () => {
    const c = parseChord("Ctrl+Shift+K")!;
    expect(chordMatches(c, ev("K", { ctrlKey: true, shiftKey: true }))).toBe(true);
    expect(chordMatches(c, ev("k", { ctrlKey: true, shiftKey: true }))).toBe(true); // case-insensitive
    expect(chordMatches(c, ev("k", { ctrlKey: true }))).toBe(false); // missing shift
    expect(chordMatches(c, ev("j", { ctrlKey: true, shiftKey: true }))).toBe(false); // wrong key
  });

  it("Mod resolves to Ctrl on win/linux (mac=false) and rejects Meta", () => {
    const c = parseChord("Mod+S")!;
    expect(chordMatches(c, ev("s", { ctrlKey: true }), false)).toBe(true);
    expect(chordMatches(c, ev("s", { metaKey: true }), false)).toBe(false); // Meta is NOT Mod on windows
    expect(chordMatches(c, ev("s"), false)).toBe(false);
  });

  it("Mod resolves to Meta on macOS (mac=true) and rejects Ctrl", () => {
    const c = parseChord("Mod+S")!;
    expect(chordMatches(c, ev("s", { metaKey: true }), true)).toBe(true);
    expect(chordMatches(c, ev("s", { ctrlKey: true }), true)).toBe(false); // Ctrl is NOT Mod on mac
  });

  it("explicit Ctrl does not match a bare Meta press", () => {
    const c = parseChord("Ctrl+S")!;
    expect(chordMatches(c, ev("s", { metaKey: true }))).toBe(false);
    expect(chordMatches(c, ev("s", { ctrlKey: true }))).toBe(true);
  });

  it("rejects when an unwanted modifier is held", () => {
    const c = parseChord("Mod+K")!;
    expect(chordMatches(c, ev("k", { ctrlKey: true, altKey: true }), false)).toBe(false);
  });
});

describe("keyMatches (keyboard-layout-independent shortcut matching)", () => {
  it("matches an ASCII produced character (a Latin layout), case-insensitively", () => {
    expect(keyMatches("z", { key: "z", code: "KeyZ" })).toBe(true);
    expect(keyMatches("z", { key: "Z", code: "KeyZ" })).toBe(true);
    expect(keyMatches("z", { key: "y", code: "KeyY" })).toBe(false);
  });

  it("falls back to the physical code when the produced char is non-ASCII (Cyrillic)", () => {
    // Ukrainian layout: the physical Z key emits 'я', the physical A key emits 'ф'.
    expect(keyMatches("z", { key: "я", code: "KeyZ" })).toBe(true);
    expect(keyMatches("a", { key: "ф", code: "KeyA" })).toBe(true);
    expect(keyMatches("z", { key: "я", code: "KeyX" })).toBe(false); // wrong physical key
  });

  it("trusts the ASCII char for deliberate remaps (Dvorak) and ignores the code", () => {
    // Dvorak: the physical ';' key produces 'z' — the user expects 'z' shortcuts there.
    expect(keyMatches("z", { key: "z", code: "KeySemicolon" })).toBe(true);
    // ...and the physical Z position (now producing ';') is NOT a 'z' shortcut.
    expect(keyMatches("z", { key: ";", code: "KeyZ" })).toBe(false);
  });

  it("matches layout-independent named keys directly on key", () => {
    expect(keyMatches("enter", { key: "Enter", code: "Enter" })).toBe(true);
    expect(keyMatches("arrowup", { key: "ArrowUp", code: "ArrowUp" })).toBe(true);
    expect(keyMatches("f2", { key: "F2", code: "F2" })).toBe(true);
    expect(keyMatches("enter", { key: "Escape", code: "Escape" })).toBe(false);
  });

  it("recovers punctuation shortcuts by physical position when the char is non-ASCII", () => {
    expect(keyMatches("/", { key: "/", code: "Slash" })).toBe(true); // ASCII '/', trusted
    expect(keyMatches("/", { key: "ж", code: "Slash" })).toBe(true); // non-ASCII → physical Slash
  });

  it("degrades to key-only when code is missing (virtual keyboards / old browsers)", () => {
    expect(keyMatches("z", { key: "z" })).toBe(true);
    expect(keyMatches("z", { key: "я" })).toBe(false); // no code → cannot recover; never throws
    expect(keyMatches("z", { key: "я", code: "" })).toBe(false);
  });
});

describe("chordMatches under non-Latin keyboard layouts", () => {
  it("matches Ctrl+Z when a Cyrillic layout emits 'я' on the physical Z", () => {
    expect(chordMatches(parseChord("Ctrl+Z")!, ev("я", { ctrlKey: true, code: "KeyZ" }))).toBe(true);
  });
  it("matches Ctrl+A when the physical A emits 'ф'", () => {
    expect(chordMatches(parseChord("Ctrl+A")!, ev("ф", { ctrlKey: true, code: "KeyA" }))).toBe(true);
  });
  it("still matches a Dvorak user through the produced ASCII letter", () => {
    expect(chordMatches(parseChord("Ctrl+Z")!, ev("z", { ctrlKey: true, code: "KeySemicolon" }))).toBe(true);
  });
  it("does not match when the physical key is wrong on a non-Latin layout", () => {
    expect(chordMatches(parseChord("Ctrl+Z")!, ev("я", { ctrlKey: true, code: "KeyX" }))).toBe(false);
  });
});

describe("normalizeChord", () => {
  it("is order-independent", () => {
    expect(normalizeChord(parseChord("Shift+Ctrl+K")!)).toBe(normalizeChord(parseChord("Ctrl+Shift+K")!));
  });

  it("resolves Mod to Ctrl on win/linux and Meta on mac", () => {
    expect(normalizeChord(parseChord("Mod+K")!, false)).toBe(normalizeChord(parseChord("Ctrl+K")!, false));
    expect(normalizeChord(parseChord("Mod+K")!, true)).toBe(normalizeChord(parseChord("Meta+K")!, true));
    // ...and NOT the opposite modifier
    expect(normalizeChord(parseChord("Mod+K")!, false)).not.toBe(normalizeChord(parseChord("Meta+K")!, false));
  });
});

describe("resolveCommandBindings", () => {
  const cmd = (id: string, keybinding?: string | string[]): EditorCommand => ({ id, ...(keybinding !== undefined ? { keybinding } : {}), run: () => {} });

  it("flattens single + array bindings", () => {
    const { bindings } = resolveCommandBindings([cmd("a", "Mod+A"), cmd("b", ["Mod+B", "F2"])]);
    expect(bindings.map((b) => b.commandId)).toEqual(["a", "b", "b"]);
    expect(bindings.map((b) => b.raw)).toEqual(["Mod+A", "Mod+B", "F2"]);
  });

  it("first command to claim a chord wins; the later one is a conflict", () => {
    const { bindings, conflicts } = resolveCommandBindings([cmd("a", "Mod+K"), cmd("b", "Mod+K")]);
    expect(bindings).toHaveLength(1);
    expect(bindings[0]!.commandId).toBe("a");
    expect(conflicts).toEqual([{ raw: "Mod+K", commandId: "b", heldBy: "a" }]);
  });

  it("reports duplicate ids and keeps the first", () => {
    const { bindings, duplicateIds } = resolveCommandBindings([cmd("a", "Mod+A"), cmd("a", "Mod+B")]);
    expect(duplicateIds).toEqual(["a"]);
    expect(bindings.map((b) => b.raw)).toEqual(["Mod+A"]); // second "a" dropped entirely
  });

  it("ignores commands with no keybinding", () => {
    const { bindings } = resolveCommandBindings([cmd("a"), cmd("b", "Mod+B")]);
    expect(bindings.map((b) => b.commandId)).toEqual(["b"]);
  });

  it("detects Mod-vs-explicit conflicts per platform", () => {
    // On win/linux, Mod == Ctrl, so "Mod+K" and "Ctrl+K" collide (but "Meta+K" doesn't).
    const win = resolveCommandBindings([cmd("a", "Mod+K"), cmd("b", "Ctrl+K"), cmd("c", "Meta+K")], false);
    expect(win.conflicts.map((c) => c.commandId)).toEqual(["b"]);
    expect(win.bindings.map((b) => b.commandId)).toEqual(["a", "c"]);
    // On macOS, Mod == Meta, so "Mod+K" collides with "Meta+K" instead.
    const mac = resolveCommandBindings([cmd("a", "Mod+K"), cmd("b", "Ctrl+K"), cmd("c", "Meta+K")], true);
    expect(mac.conflicts.map((c) => c.commandId)).toEqual(["c"]);
    expect(mac.bindings.map((b) => b.commandId)).toEqual(["a", "b"]);
  });
});
