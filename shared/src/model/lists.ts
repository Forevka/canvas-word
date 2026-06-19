// List definitions — a direct mirror of OOXML numbering.xml abstractNum levels,
// so the docx importer can feed Document.lists losslessly (id space = numId).
// Markers are PAINT-ONLY: they live in the hanging indent and never affect line
// breaking (matching Word's overlap behavior for very long markers).

import type { CharStyle } from "./document";

export type ListNumberFormat =
  | "bullet"
  | "decimal"
  | "lowerLetter"
  | "upperLetter"
  | "lowerRoman"
  | "upperRoman";

export interface ListLevel {
  format: ListNumberFormat;
  /** Marker pattern; %N is the counter of level N-1, formatted per THAT level
   *  (OOXML §17.9.11 semantics). E.g. "%1." or "%1.%2.%3". Ignored for bullets. */
  text: string;
  bulletChar?: string;
  /** Text indent for paragraphs at this level (added to ParaStyle.indentLeftPx). */
  indentLeftPx: number;
  /** Marker hangs this far left of the text origin. */
  hangingPx: number;
  start: number;
  markerStyle?: Partial<CharStyle>;
}

export interface ListDefinition {
  id: string;
  levels: ListLevel[]; // up to 9 (levels 0..8)
}

// ---------------------------------------------------------------------------
// Number formatting

const ROMAN: [number, string][] = [
  [1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"],
  [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"],
];

function toRoman(n: number): string {
  let out = "";
  let v = Math.max(1, n);
  for (const [value, sym] of ROMAN) {
    while (v >= value) {
      out += sym;
      v -= value;
    }
  }
  return out;
}

function toLetter(n: number): string {
  // 1->a .. 26->z, 27->aa (Word repeats: 27 is actually "aa")
  let out = "";
  let v = Math.max(1, n);
  while (v > 0) {
    v -= 1;
    out = String.fromCharCode(97 + (v % 26)) + out;
    v = Math.floor(v / 26);
  }
  return out;
}

export function formatListNumber(n: number, format: ListNumberFormat): string {
  switch (format) {
    case "decimal": return String(n);
    case "lowerLetter": return toLetter(n);
    case "upperLetter": return toLetter(n).toUpperCase();
    case "lowerRoman": return toRoman(n);
    case "upperRoman": return toRoman(n).toUpperCase();
    case "bullet": return "";
  }
}

/** Render a level's marker from the counter stack (counters[k] = level k's value). */
export function markerText(def: ListDefinition, level: number, counters: number[]): string {
  const lvl = def.levels[Math.min(level, def.levels.length - 1)];
  if (!lvl) return "";
  if (lvl.format === "bullet") return lvl.bulletChar ?? DEFAULT_BULLET;
  return lvl.text.replace(/%(\d)/g, (_, d: string) => {
    const k = Number(d) - 1;
    const refLevel = def.levels[Math.min(k, def.levels.length - 1)];
    const value = counters[k] ?? refLevel?.start ?? 1;
    return formatListNumber(value, refLevel?.format ?? "decimal");
  });
}

// ---------------------------------------------------------------------------
// Default definitions (what the toolbar buttons create)

export const DEFAULT_BULLET_LIST_ID = "bullets";
export const DEFAULT_NUMBER_LIST_ID = "numbers";
export const DEFAULT_MULTILEVEL_LIST_ID = "multilevel";

/** OOXML abstractNum holds up to 9 levels (0..8). */
const LIST_LEVELS = 9;
/** Per-level text indent + step (px): level i sits at LIST_INDENT_STEP_PX*(i+1). */
const LIST_INDENT_STEP_PX = 24;
/** How far the marker hangs left of the text origin, per list kind. */
const BULLET_HANGING_PX = 18;
const NUMBER_HANGING_PX = 22;

const DEFAULT_BULLET = "•";
const BULLETS = [DEFAULT_BULLET, "◦", "▪"];
const NUMBER_FORMATS: ListNumberFormat[] = ["decimal", "lowerLetter", "lowerRoman"];

const levelIndentPx = (i: number): number => LIST_INDENT_STEP_PX * (i + 1);

export function defaultListDefinition(kind: "bullet" | "decimal"): ListDefinition {
  const levels: ListLevel[] = [];
  for (let i = 0; i < LIST_LEVELS; i++) {
    if (kind === "bullet") {
      levels.push({
        format: "bullet",
        text: "",
        bulletChar: BULLETS[i % BULLETS.length]!,
        indentLeftPx: levelIndentPx(i),
        hangingPx: BULLET_HANGING_PX,
        start: 1,
      });
    } else {
      levels.push({
        format: NUMBER_FORMATS[i % NUMBER_FORMATS.length]!,
        text: `%${i + 1}.`,
        indentLeftPx: levelIndentPx(i),
        hangingPx: NUMBER_HANGING_PX,
        start: 1,
      });
    }
  }
  return { id: kind === "bullet" ? DEFAULT_BULLET_LIST_ID : DEFAULT_NUMBER_LIST_ID, levels };
}

/** A single-glyph bullet list (every level uses `char`) — what the bullet
 *  style picker applies. */
export function bulletListDefinition(id: string, char: string): ListDefinition {
  const levels: ListLevel[] = [];
  for (let i = 0; i < LIST_LEVELS; i++) {
    levels.push({ format: "bullet", text: "", bulletChar: char, indentLeftPx: levelIndentPx(i), hangingPx: BULLET_HANGING_PX, start: 1 });
  }
  return { id, levels };
}

/** A single-format numbered list (every level uses `format`), with `suffix`
 *  after each counter ("." or ")") — what the number style picker applies. */
export function numberListDefinition(id: string, format: ListNumberFormat, suffix: string): ListDefinition {
  const levels: ListLevel[] = [];
  for (let i = 0; i < LIST_LEVELS; i++) {
    levels.push({ format, text: `%${i + 1}${suffix}`, indentLeftPx: levelIndentPx(i), hangingPx: NUMBER_HANGING_PX, start: 1 });
  }
  return { id, levels };
}

/** Legal-style multilevel list: every level is decimal and its marker compounds
 *  all ancestors — level 0 "1.", level 1 "1.1.", level 2 "1.1.1.", … (Word's
 *  default Multilevel List). Tab/Shift+Tab walk the levels. */
export function multilevelListDefinition(): ListDefinition {
  const levels: ListLevel[] = [];
  for (let i = 0; i < LIST_LEVELS; i++) {
    const text = Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join(".") + ".";
    levels.push({
      format: "decimal",
      text, // "%1.", "%1.%2.", "%1.%2.%3.", …
      indentLeftPx: levelIndentPx(i),
      hangingPx: LIST_INDENT_STEP_PX + i * 8, // compound markers widen with depth
      start: 1,
    });
  }
  return { id: DEFAULT_MULTILEVEL_LIST_ID, levels };
}
