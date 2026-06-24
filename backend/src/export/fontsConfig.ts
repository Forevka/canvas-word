// Validate + normalize a (possibly untrusted) fonts config into the resolved
// shape the export pipeline assumes. POST /docs and POST /render.pdf both accept a
// caller-supplied fontsConfig; resolveCustomFontBytes later trusts `family` +
// `faces.regular` + `sizing`, so we drop any entry that doesn't carry them rather
// than letting a malformed object reach the pipeline.

import type { ResolvedFontsConfig } from "@forevka/wordcanvas/export";

type FontDef = ResolvedFontsConfig["fonts"][number];

function isObject(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

function nonEmptyString(v: unknown): v is string {
  return typeof v === "string" && v.length > 0;
}

function normalizeFont(raw: unknown): FontDef | undefined {
  if (!isObject(raw)) return undefined;
  if (!nonEmptyString(raw.family)) return undefined;
  if (!isObject(raw.faces) || !nonEmptyString(raw.faces.regular)) return undefined;
  const sizing = raw.sizing;
  if (!isObject(sizing)) return undefined;
  const { ascent, descent } = sizing;
  if (typeof ascent !== "number" || !Number.isFinite(ascent) || ascent <= 0) return undefined;
  if (typeof descent !== "number" || !Number.isFinite(descent) || descent < 0) return undefined;

  const faces: FontDef["faces"] = { regular: raw.faces.regular };
  if (nonEmptyString(raw.faces.bold)) faces.bold = raw.faces.bold;
  if (nonEmptyString(raw.faces.italic)) faces.italic = raw.faces.italic;
  if (nonEmptyString(raw.faces.boldItalic)) faces.boldItalic = raw.faces.boldItalic;

  const def: FontDef = { family: raw.family, faces, sizing: { ascent, descent } };
  if (nonEmptyString(raw.label)) def.label = raw.label;
  return def;
}

/** Normalize a (possibly partial / untrusted) fonts config JSON into the resolved
 *  shape. Returns undefined for non-objects; drops any font entry missing a valid
 *  family / faces.regular / sizing so the export pipeline only ever sees valid defs. */
export function normalizeFontsConfig(raw: unknown): ResolvedFontsConfig | undefined {
  if (!isObject(raw)) return undefined;
  const disableBuiltin = Array.isArray(raw.disableBuiltin)
    ? raw.disableBuiltin.filter((x): x is string => typeof x === "string")
    : [];
  const fonts = Array.isArray(raw.fonts)
    ? raw.fonts.map(normalizeFont).filter((f): f is FontDef => f !== undefined)
    : [];
  return { disableBuiltin, fonts };
}
