// numbering.xml → IR list definitions, keyed by numId (the id space paragraphs
// reference via w:numPr/w:numId). Resolves w:num → w:abstractNum, applies
// w:lvlOverride. mapToModel converts these to model ListDefinitions (twips→px,
// marker rPr → CharStyle); the model renders markers paint-only.

import { decodeRunProps } from "./props";
import type { IRListDefinition, IRListLevel } from "./types";
import { attr, el, els, numAttr, parseXml, rootEl, val, type XmlNode } from "./xml";

export type NumberingData = Map<string, IRListDefinition>;

export const EMPTY_NUMBERING: NumberingData = new Map();

export function parseNumberingXml(xmlText: string): NumberingData {
  const out: NumberingData = new Map();
  const root = rootEl(parseXml(xmlText, "numbering.xml"), "w:numbering");
  if (!root) return out;

  // abstractNumId → its 0..8 levels.
  const abstracts = new Map<string, IRListLevel[]>();
  for (const abs of els(root, "w:abstractNum")) {
    const absId = attr(abs, "w:abstractNumId");
    if (absId === undefined) continue;
    abstracts.set(absId, parseLevels(abs));
  }

  // w:num maps a public numId onto an abstractNum, with optional level overrides.
  for (const num of els(root, "w:num")) {
    const numId = attr(num, "w:numId");
    if (numId === undefined) continue;
    const absId = val(num, "w:abstractNumId");
    const base = absId !== undefined ? abstracts.get(absId) : undefined;
    if (!base) continue; // numStyleLink-only definitions aren't resolved
    const levels = base.map((l) => ({ ...l })); // copy — overrides mutate per-num
    for (const ovr of els(num, "w:lvlOverride")) {
      const ilvl = numAttr(ovr, "w:ilvl") ?? 0;
      const startOverride = numAttr(el(ovr, "w:startOverride"), "w:val");
      const lvlEl = el(ovr, "w:lvl");
      if (lvlEl && levels[ilvl]) levels[ilvl] = parseLevel(lvlEl, ilvl);
      if (startOverride !== undefined && levels[ilvl]) levels[ilvl]!.start = startOverride;
    }
    out.set(numId, { id: numId, levels });
  }
  return out;
}

function parseLevels(abstractNum: XmlNode): IRListLevel[] {
  const levels: IRListLevel[] = [];
  for (const lvl of els(abstractNum, "w:lvl")) {
    const ilvl = numAttr(lvl, "w:ilvl") ?? levels.length;
    levels[ilvl] = parseLevel(lvl, ilvl);
  }
  return levels;
}

function parseLevel(lvl: XmlNode, ilvl: number): IRListLevel {
  // lvlText present but valueless ("") is meaningful (a "none" marker); only an
  // absent element falls back to the conventional "%n." pattern.
  const lvlTextEl = el(lvl, "w:lvlText");
  const level: IRListLevel = {
    format: val(lvl, "w:numFmt") ?? "decimal",
    lvlText: lvlTextEl ? (attr(lvlTextEl, "w:val") ?? "") : `%${ilvl + 1}.`,
    start: numAttr(el(lvl, "w:start"), "w:val") ?? 1,
  };

  const pPr = el(lvl, "w:pPr");
  const ind = pPr && el(pPr, "w:ind");
  if (ind) {
    const left = numAttr(ind, "w:left") ?? numAttr(ind, "w:start");
    if (left !== undefined) level.indentLeftTwips = left;
    const hanging = numAttr(ind, "w:hanging");
    if (hanging !== undefined) level.hangingTwips = hanging;
  }
  const rPr = el(lvl, "w:rPr");
  if (rPr) level.markerRunProps = decodeRunProps(rPr);
  return level;
}
