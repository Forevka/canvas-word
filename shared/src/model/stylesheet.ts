// Named styles (Word's style gallery). The editor model keeps runs CONCRETE —
// styles here are templates + references, not a live cascade at render time:
//   - applying a style patches exactly the fields the style defines (direct
//     formatting on other fields survives, like Word),
//   - modifying a style re-patches every paragraph that references it,
//   - resolution walks the basedOn chain (editor-simple override semantics;
//     OOXML XOR-toggle fidelity lives in import/docx/styles.ts, which resolves
//     to concrete values at import time and preserves the styleId reference —
//     the two systems share the same id space).

import type { CharStyle, Document, ParaStyle, Paragraph } from "./document";
import { paragraphsOf } from "./text";

export type NamedStyleType = "paragraph" | "character";

export interface NamedStyle {
  id: string; // shares the docx styleId space ("Normal", "Heading1", ...)
  name: string; // display name
  /** Paragraph vs character style (Word's two text-level style kinds). Optional
   *  for back-compat: legacy stylesheets and defaultStylesheet() omit it and read
   *  as "paragraph" via styleType(). Character styles store only `char` (empty
   *  `para`) and apply to runs, not paragraphs. */
  type?: NamedStyleType;
  basedOn?: string;
  char: Partial<CharStyle>;
  para: Partial<ParaStyle>;
}

export interface Stylesheet {
  styles: NamedStyle[];
  defaultStyleId: string;
}

/** A style's kind, defaulting untyped (legacy) styles to "paragraph". Read this
 *  rather than `.type` directly so old documents keep working. */
export function styleType(s: NamedStyle): NamedStyleType {
  return s.type ?? "paragraph";
}

export function styleById(sheet: Stylesheet, id: string): NamedStyle | undefined {
  return sheet.styles.find((s) => s.id === id);
}

/** Resolved templates: basedOn chain root→leaf, defined fields override. */
export function resolveStyle(
  sheet: Stylesheet,
  id: string,
): { char: Partial<CharStyle>; para: Partial<ParaStyle> } {
  const chain: NamedStyle[] = [];
  const seen = new Set<string>();
  for (let cur = styleById(sheet, id); cur && !seen.has(cur.id); cur = cur.basedOn ? styleById(sheet, cur.basedOn) : undefined) {
    seen.add(cur.id);
    chain.unshift(cur);
  }
  const char: Partial<CharStyle> = {};
  const para: Partial<ParaStyle> = {};
  for (const s of chain) {
    Object.assign(char, definedOnly(s.char));
    Object.assign(para, definedOnly(s.para));
  }
  return { char, para };
}

/** Resolved character template only: basedOn chain root→leaf, defined `char`
 *  fields override. Character styles have no `para`, so this is the run-level
 *  equivalent of resolveStyle for type==="character" styles. */
export function resolveCharStyle(sheet: Stylesheet, id: string): Partial<CharStyle> {
  const chain: NamedStyle[] = [];
  const seen = new Set<string>();
  for (let cur = styleById(sheet, id); cur && !seen.has(cur.id); cur = cur.basedOn ? styleById(sheet, cur.basedOn) : undefined) {
    seen.add(cur.id);
    chain.unshift(cur);
  }
  const char: Partial<CharStyle> = {};
  for (const s of chain) Object.assign(char, definedOnly(s.char));
  return char;
}

/** Every style id that (transitively) lists `id` as its basedOn ancestor. Used
 *  to keep a style out of its own descendants' "Based on" pickers. */
export function descendantsOf(sheet: Stylesheet, id: string): Set<string> {
  const out = new Set<string>();
  let grew = true;
  while (grew) {
    grew = false;
    for (const s of sheet.styles) {
      if (out.has(s.id)) continue;
      if (s.basedOn !== undefined && (s.basedOn === id || out.has(s.basedOn))) {
        out.add(s.id);
        grew = true;
      }
    }
  }
  return out;
}

/** Would setting `styleId`'s basedOn to `newBasedOn` create a cycle? True if
 *  newBasedOn is the style itself or one of its descendants. A backstop for the
 *  UI, which already excludes these from the picker. */
export function wouldCycle(sheet: Stylesheet, styleId: string, newBasedOn: string | undefined): boolean {
  if (!newBasedOn) return false;
  if (newBasedOn === styleId) return true;
  return descendantsOf(sheet, styleId).has(newBasedOn);
}

function definedOnly<T extends object>(o: T): Partial<T> {
  const out: Partial<T> = {};
  for (const [k, v] of Object.entries(o)) {
    if (v !== undefined) (out as Record<string, unknown>)[k] = v;
  }
  return out;
}

/** Every editable paragraph referencing a style — paragraphsOf already covers
 *  body, table cells, and both margin bands in document order. */
export function paragraphsWithStyle(doc: Document, styleId: string): Paragraph[] {
  return paragraphsOf(doc).filter((p) => p.style.namedStyle === styleId);
}

export function defaultStylesheet(): Stylesheet {
  const base: Partial<ParaStyle> = { lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 12 };
  return {
    defaultStyleId: "Normal",
    styles: [
      {
        id: "Normal",
        name: "Normal",
        char: { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, color: "#202124" },
        para: { ...base, align: "left", indentFirstLinePx: 0 },
      },
      {
        id: "Title",
        name: "Title",
        basedOn: "Normal",
        char: { fontFamily: "Arial, sans-serif", fontSizePx: 32, bold: true, color: "#1a1a2e" },
        para: { align: "center", spaceAfterPx: 4 },
      },
      {
        id: "Subtitle",
        name: "Subtitle",
        basedOn: "Normal",
        char: { italic: true, color: "#5f6368" },
        para: { align: "center", spaceAfterPx: 28 },
      },
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        char: { fontFamily: "Arial, sans-serif", fontSizePx: 24, bold: true, color: "#1a1a2e" },
        para: { spaceBeforePx: 18, spaceAfterPx: 8, keepWithNext: true },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Heading1",
        char: { fontSizePx: 19 },
        para: { spaceBeforePx: 14, spaceAfterPx: 6 },
      },
      {
        id: "Quote",
        name: "Quote",
        basedOn: "Normal",
        char: { italic: true, color: "#5f6368" },
        para: { indentLeftPx: 36, spaceBeforePx: 8, spaceAfterPx: 8 },
      },
      {
        id: "Code",
        name: "Code",
        basedOn: "Normal",
        char: { fontFamily: "Consolas, monospace", fontSizePx: 14, color: "#0b57d0" },
        para: { lineHeight: 1.35 },
      },
    ],
  };
}
