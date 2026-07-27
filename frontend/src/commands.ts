// Public command + keymap registry — types + pure keybinding helpers.
//
// An embedder passes `commands: EditorCommand[]` to the WordCanvas constructor to
// register named commands and (optionally) bind keyboard shortcuts to them. A
// command's `run` receives the SAME `RibbonActionContext` a custom ribbon button
// gets (the editor handle + macro helpers), so a macro is interchangeable between
// a button and a shortcut. Commands are also invokable programmatically via
// `handle.runCommand(id)`.
//
// This module is DOM-free so its chord parsing/matching/conflict logic is
// unit-testable; editorApp.ts owns the actual keydown wiring (see the
// `commands` block there). The built-in editing chords (Ctrl+B/I/U/Z/Y,
// Ctrl+Enter — see input/keymap.ts) always take precedence: the wiring skips a
// custom binding whenever the event was already handled (defaultPrevented).

import type { RibbonActionContext } from "./ribbon";

/** The context handed to a command's `run` — identical to a ribbon button's, so
 *  the same macro works from a keystroke, a button, or `handle.runCommand`. */
export type CommandContext = RibbonActionContext;

/** A registered command. `keybinding` is a chord string ("Mod+Shift+K") or an
 *  array of them; omit it for a command that is only invoked programmatically or
 *  wired to a ribbon button by id. */
export interface EditorCommand {
  /** Stable id (namespace it to avoid clashes, e.g. "myco.macros.upper"). */
  id: string;
  /** Human label (for docs / a future command palette). Optional. */
  label?: string;
  /** Keyboard shortcut(s). `Mod` = Ctrl on Windows/Linux, Cmd (Meta) on macOS.
   *  Modifiers: Mod, Ctrl, Alt, Shift, Meta. Key is the final segment, matched
   *  case-insensitively against `KeyboardEvent.key` (e.g. "k", "Enter", "F2",
   *  "ArrowUp", "/"). */
  keybinding?: string | string[];
  /** Invoked with the editor handle + macro helpers. May be async. Synchronous
   *  throws AND rejected promises are caught and logged by the wiring, so one bad
   *  command can't break the keymap. */
  run: (ctx: CommandContext) => void | Promise<void>;
}

/** A parsed keybinding chord. `mod` is the platform-agnostic Ctrl/Cmd modifier;
 *  when set it matches Ctrl OR Meta so a single binding works on every OS. */
export interface ParsedChord {
  key: string; // lower-cased KeyboardEvent.key
  mod: boolean;
  ctrl: boolean;
  alt: boolean;
  shift: boolean;
  meta: boolean;
}

const MODIFIERS = new Set(["mod", "ctrl", "control", "alt", "option", "shift", "meta", "cmd", "command", "win", "super"]);

/** Parse a chord string ("Mod+Shift+K") into a ParsedChord, or null if it has no
 *  non-modifier key. Whitespace and case are ignored; `+` separates segments. */
export function parseChord(binding: string): ParsedChord | null {
  const chord: ParsedChord = { key: "", mod: false, ctrl: false, alt: false, shift: false, meta: false };
  const parts = binding
    .split("+")
    .map((p) => p.trim())
    .filter((p) => p.length > 0);
  for (const part of parts) {
    const lower = part.toLowerCase();
    if (MODIFIERS.has(lower)) {
      if (lower === "mod") chord.mod = true;
      else if (lower === "ctrl" || lower === "control") chord.ctrl = true;
      else if (lower === "alt" || lower === "option") chord.alt = true;
      else if (lower === "shift") chord.shift = true;
      else chord.meta = true; // meta / cmd / command / win / super
    } else {
      // The final non-modifier segment is the key (last one wins).
      chord.key = lower;
    }
  }
  return chord.key ? chord : null;
}

/** A KeyboardEvent-shaped input, so matching stays testable without a real DOM.
 *  `code` is the physical-key position (layout-independent); optional so callers
 *  and older environments that lack it degrade to key-only matching. */
export interface KeyEventLike {
  key: string;
  code?: string;
  ctrlKey: boolean;
  altKey: boolean;
  shiftKey: boolean;
  metaKey: boolean;
}

/** The `KeyboardEvent.code` for a single-character shortcut target at its US-QWERTY
 *  physical position (letters → `KeyX`, digits → `DigitN`, common punctuation).
 *  Returns null for a target with no stable physical mapping. */
function physicalCodeFor(target: string): string | null {
  if (target >= "a" && target <= "z") return "Key" + target.toUpperCase();
  if (target >= "0" && target <= "9") return "Digit" + target;
  const PUNCT: Record<string, string> = {
    "/": "Slash", ".": "Period", ",": "Comma", ";": "Semicolon", "'": "Quote",
    "[": "BracketLeft", "]": "BracketRight", "\\": "Backslash",
    "-": "Minus", "=": "Equal", "`": "Backquote",
  };
  return PUNCT[target] ?? null;
}

/** Does a keyboard event's key match the shortcut target `target` (a lower-cased
 *  key name — "z", "/", "enter", "arrowup", "f2")? HYBRID layout handling, the
 *  crux of the fix for non-Latin keyboards:
 *
 *  - `KeyboardEvent.key` carries the LAYOUT-PRODUCED character, so on a Cyrillic
 *    layout the physical Z key yields `key === 'я'` and no ASCII shortcut matches —
 *    every Ctrl-chord was silently dead for those users.
 *  - We match the produced character when it is ASCII (a Latin layout, or a
 *    deliberate remap like Dvorak/AZERTY whose users expect the keycap legend), and
 *    fall back to the physical position via `event.code` (e.g. 'KeyZ') ONLY when the
 *    produced character is not ASCII. This is what VS Code / Word do, and why
 *    Cyrillic/Greek/… keycaps also carry the Latin legend.
 *  - Named keys (Enter, ArrowUp, F2, …) are layout-independent, so they always match
 *    on `key` directly.
 *  - If `code` is missing/empty (some virtual keyboards, older browsers) we degrade
 *    to key-only matching rather than failing — the exact pre-fix behaviour.
 *
 *  Note: this is for MODIFIER shortcuts. A bare character trigger (e.g. typing "/"
 *  to open the slash menu) must keep matching `key`, because there the user means the
 *  character they produced, not a physical position. */
export function keyMatches(target: string, ev: { key: string; code?: string }): boolean {
  const t = target.toLowerCase();
  const rawKey = ev.key ?? "";
  const key = rawKey.toLowerCase();
  // Named (multi-char) keys are layout-independent — compare directly.
  if (t.length !== 1) return key === t;
  // Single-char target: an ASCII produced character is authoritative (do NOT also
  // consult code, so a remapped 'z' matches key 'z' and not physical KeyZ).
  if (rawKey.length === 1 && rawKey.charCodeAt(0) < 128) return key === t;
  // Produced character is non-ASCII (or absent) → use the physical key position.
  const code = ev.code ?? "";
  if (!code) return key === t; // no code available → degrade to key-only
  const want = physicalCodeFor(t);
  return want !== null && code === want;
}

/** Resolve `mod` to the concrete modifier the active platform uses: Meta (⌘) on
 *  macOS, Ctrl elsewhere. Returns the effective {ctrl, meta} a chord requires. */
function effectiveModifiers(chord: ParsedChord, mac: boolean): { ctrl: boolean; meta: boolean } {
  return {
    ctrl: chord.ctrl || (chord.mod && !mac),
    meta: chord.meta || (chord.mod && mac),
  };
}

/** Does an event satisfy a parsed chord? `mod` resolves to Ctrl on Windows/Linux
 *  and Meta (⌘) on macOS (via `mac`), and the OPPOSITE modifier must be absent —
 *  so a Windows "Mod+S" fires on Ctrl+S but not on the Windows/Meta key. Explicit
 *  `ctrl`/`meta` still require that exact modifier. All four modifiers are matched
 *  exactly, so an extra held modifier never counts as a match. */
export function chordMatches(chord: ParsedChord, ev: KeyEventLike, mac = false): boolean {
  if (!keyMatches(chord.key, ev)) return false; // hybrid key/code match (non-Latin layouts)
  if (chord.shift !== ev.shiftKey) return false;
  if (chord.alt !== ev.altKey) return false;
  const { ctrl, meta } = effectiveModifiers(chord, mac);
  return ev.ctrlKey === ctrl && ev.metaKey === meta;
}

/** A canonical, order-independent string for a chord — used to detect two
 *  commands claiming the same shortcut. `mac` resolves `mod` to its concrete
 *  modifier so, e.g., "Mod+K" and "Ctrl+K" are detected as the same chord on
 *  Windows/Linux (and "Mod+K" and "Meta+K" on macOS). */
export function normalizeChord(chord: ParsedChord, mac = false): string {
  const { ctrl, meta } = effectiveModifiers(chord, mac);
  const mods: string[] = [];
  if (ctrl) mods.push("ctrl");
  if (meta) mods.push("meta");
  if (chord.alt) mods.push("alt");
  if (chord.shift) mods.push("shift");
  mods.sort();
  return [...mods, chord.key].join("+");
}

/** One resolved shortcut → command binding. */
export interface CommandBinding {
  chord: ParsedChord;
  commandId: string;
  /** The raw binding string, for diagnostics. */
  raw: string;
}

/** Flatten commands into an ordered binding list and report conflicts.
 *
 *  - Skips commands with a duplicate `id` (first registration wins).
 *  - A later binding on the same normalized chord is dropped and reported in
 *    `conflicts` so the wiring can warn — first-registered shortcut wins, which
 *    keeps behavior deterministic regardless of registration timing.
 *  Pure — no DOM, no console; the caller decides how to surface `conflicts`. */
export function resolveCommandBindings(
  commands: readonly EditorCommand[],
  mac = false,
): {
  bindings: CommandBinding[];
  conflicts: { raw: string; commandId: string; heldBy: string }[];
  duplicateIds: string[];
} {
  const bindings: CommandBinding[] = [];
  const conflicts: { raw: string; commandId: string; heldBy: string }[] = [];
  const duplicateIds: string[] = [];
  const seenIds = new Set<string>();
  const claimed = new Map<string, string>(); // normalized chord → commandId

  for (const cmd of commands) {
    if (seenIds.has(cmd.id)) {
      duplicateIds.push(cmd.id);
      continue;
    }
    seenIds.add(cmd.id);
    const raws = cmd.keybinding === undefined ? [] : Array.isArray(cmd.keybinding) ? cmd.keybinding : [cmd.keybinding];
    for (const raw of raws) {
      const chord = parseChord(raw);
      if (!chord) continue;
      const norm = normalizeChord(chord, mac);
      const heldBy = claimed.get(norm);
      if (heldBy !== undefined) {
        conflicts.push({ raw, commandId: cmd.id, heldBy });
        continue;
      }
      claimed.set(norm, cmd.id);
      bindings.push({ chord, commandId: cmd.id, raw });
    }
  }
  return { bindings, conflicts, duplicateIds };
}
