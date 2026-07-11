// Style delta helpers — the keys whose values differ between two styles, with
// the forward (new) values as `patch` and the old values as `inverse`. Mirrors
// frontend/src/review/intercept.ts styleDelta, kept here so the isomorphic diff
// core doesn't depend on the frontend. `patch` drives what the merged document
// carries (the revised style); `inverse` is what a reject re-applies to restore
// the original. Pure + DOM-free.

import type { CharStyle, ParaStyle } from "../model/document";

export interface StyleDelta<S> {
  patch: Partial<S>;
  inverse: Partial<S>;
}

function keysOf<S extends object>(a: S, b: S): (keyof S)[] {
  return [...new Set([...(Object.keys(a) as (keyof S)[]), ...(Object.keys(b) as (keyof S)[])])];
}

/** Shallow value equality good enough for style fields: primitives compare by
 *  ===; objects/arrays (list, tabStops, borders, sectionBreak, runBorder…) by a
 *  stable JSON serialization. */
function valueEq(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (a === undefined || b === undefined || a === null || b === null) return false;
  if (typeof a === "object" || typeof b === "object") return JSON.stringify(a) === JSON.stringify(b);
  return false;
}

function delta<S extends object>(oldS: S, newS: S): StyleDelta<S> {
  const patch: Partial<S> = {};
  const inverse: Partial<S> = {};
  for (const k of keysOf(oldS, newS)) {
    if (!valueEq(oldS[k], newS[k])) {
      patch[k] = newS[k];
      inverse[k] = oldS[k];
    }
  }
  return { patch, inverse };
}

/** Char-style delta between two runs' styles (oldS → newS). */
export function charStyleDelta(oldS: CharStyle, newS: CharStyle): StyleDelta<CharStyle> {
  return delta(oldS, newS);
}

/** Paragraph-style delta (oldS → newS). Empty patch ⇒ no paragraph-level change. */
export function paraStyleDelta(oldS: ParaStyle, newS: ParaStyle): StyleDelta<ParaStyle> {
  return delta(oldS, newS);
}
