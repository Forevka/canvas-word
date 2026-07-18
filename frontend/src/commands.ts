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
  /** Invoked with the editor handle + macro helpers. Thrown errors are caught
   *  and logged by the wiring so one bad command can't break the keymap. */
  run: (ctx: CommandContext) => void;
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

/** A KeyboardEvent-shaped input, so matching stays testable without a real DOM. */
export interface KeyEventLike {
  key: string;
  ctrlKey: boolean;
  altKey: boolean;
  shiftKey: boolean;
  metaKey: boolean;
}

/** Does an event satisfy a parsed chord? `mod` matches when Ctrl OR Meta is held
 *  (whichever the OS uses); explicit `ctrl`/`meta` require that exact modifier. */
export function chordMatches(chord: ParsedChord, ev: KeyEventLike): boolean {
  if (ev.key.toLowerCase() !== chord.key) return false;
  if (chord.shift !== ev.shiftKey) return false;
  if (chord.alt !== ev.altKey) return false;
  if (chord.mod) {
    // Platform-agnostic: accept Ctrl (win/linux) or Meta (mac).
    if (!(ev.ctrlKey || ev.metaKey)) return false;
  } else {
    if (chord.ctrl !== ev.ctrlKey) return false;
    if (chord.meta !== ev.metaKey) return false;
  }
  return true;
}

/** A canonical, order-independent string for a chord — used to detect two
 *  commands claiming the same shortcut. */
export function normalizeChord(chord: ParsedChord): string {
  const mods: string[] = [];
  if (chord.mod) mods.push("mod");
  if (chord.ctrl) mods.push("ctrl");
  if (chord.alt) mods.push("alt");
  if (chord.shift) mods.push("shift");
  if (chord.meta) mods.push("meta");
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
export function resolveCommandBindings(commands: readonly EditorCommand[]): {
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
      const norm = normalizeChord(chord);
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
