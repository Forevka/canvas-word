// numbering.xml from Document.lists. Each list becomes one w:abstractNum plus a
// w:num that maps an integer numId onto it. listIdMap gives Word-valid integer
// numIds (and is shared with the document's w:numPr references).

import type { ListDefinition, ListLevel } from "../../model/lists";
import { pxToTwips } from "../units";
import { partialRPrXml } from "./styleProps";
import { el, WML_NS, XML_DECL } from "./xmlWrite";

function levelXml(lvl: ListLevel, i: number): string {
  const c: string[] = [
    el("w:start", { "w:val": lvl.start }),
    el("w:numFmt", { "w:val": lvl.format === "bullet" ? "bullet" : lvl.format }),
    el("w:lvlText", { "w:val": lvl.format === "bullet" ? (lvl.bulletChar ?? "") : lvl.text }),
    el("w:pPr", undefined, el("w:ind", { "w:left": pxToTwips(lvl.indentLeftPx), "w:hanging": pxToTwips(lvl.hangingPx) })),
  ];
  if (lvl.markerStyle) {
    const rpr = partialRPrXml(lvl.markerStyle);
    if (rpr) c.push(rpr);
  }
  return el("w:lvl", { "w:ilvl": i }, c.join(""));
}

export function numberingXml(lists: Record<string, ListDefinition>, listIdMap: Map<string, number>): string {
  const abstracts: string[] = [];
  const nums: string[] = [];
  let absId = 0;
  for (const [listId, def] of Object.entries(lists)) {
    const numId = listIdMap.get(listId);
    if (numId === undefined) continue;
    const levels = def.levels.map((lvl, i) => levelXml(lvl, i)).join("");
    abstracts.push(el("w:abstractNum", { "w:abstractNumId": absId }, levels));
    nums.push(el("w:num", { "w:numId": numId }, el("w:abstractNumId", { "w:val": absId })));
    absId++;
  }
  return XML_DECL + el("w:numbering", WML_NS, abstracts.join("") + nums.join(""));
}
