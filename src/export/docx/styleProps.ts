// Partial rPr/pPr emitters for style-gallery patches (NamedStyle.char/.para) and
// list-marker run props. Only DEFINED fields are written — the importer reads
// each w:* element into a patch field, so absent elements stay absent.

import type { CharStyle, ParaStyle } from "../../model/document";
import { multiplierToLine, pxToTwips } from "../units";
import { el } from "./xmlWrite";

const HIGHLIGHT_NAME: Record<string, string> = {
  "#ffff00": "yellow", "#00ff00": "green", "#00ffff": "cyan", "#ff00ff": "magenta",
  "#0000ff": "blue", "#ff0000": "red", "#000080": "darkBlue", "#008080": "darkCyan",
  "#008000": "darkGreen", "#800080": "darkMagenta", "#800000": "darkRed", "#808000": "darkYellow",
  "#808080": "darkGray", "#c0c0c0": "lightGray", "#000000": "black", "#ffffff": "white",
};
const JC: Record<string, string> = { left: "left", center: "center", right: "right", justify: "both" };
const TAB_VAL: Record<string, string> = { left: "left", center: "center", right: "right", decimal: "decimal" };
const TAB_LEADER: Record<string, string> = { dot: "dot", dash: "hyphen", underscore: "underscore" };
const hex = (c: string): string => c.replace(/^#/, "").toLowerCase();

export function partialRPrXml(c: Partial<CharStyle>): string {
  const out: string[] = [];
  if (c.fontFamily) {
    const fam = (c.fontFamily.split(",")[0] ?? "").trim();
    if (fam) out.push(el("w:rFonts", { "w:ascii": fam, "w:hAnsi": fam, "w:cs": fam }));
  }
  if (c.bold !== undefined) out.push(el("w:b", { "w:val": c.bold ? "1" : "0" }));
  if (c.italic !== undefined) out.push(el("w:i", { "w:val": c.italic ? "1" : "0" }));
  if (c.strikethrough !== undefined) out.push(el("w:strike", { "w:val": c.strikethrough ? "1" : "0" }));
  if (c.color) out.push(el("w:color", { "w:val": hex(c.color) }));
  if (c.fontSizePx !== undefined) {
    const sz = Math.round(c.fontSizePx * 1.5);
    out.push(el("w:sz", { "w:val": sz }));
    out.push(el("w:szCs", { "w:val": sz }));
  }
  if (c.underline !== undefined) out.push(el("w:u", { "w:val": c.underline ? "single" : "none" }));
  if (c.highlightColor) {
    const name = HIGHLIGHT_NAME[c.highlightColor.toLowerCase()];
    if (name) out.push(el("w:highlight", { "w:val": name }));
  }
  if (c.verticalAlign) out.push(el("w:vertAlign", { "w:val": c.verticalAlign === "super" ? "superscript" : "subscript" }));
  return out.length > 0 ? el("w:rPr", undefined, out.join("")) : "";
}

export function partialPPrXml(p: Partial<ParaStyle>): string {
  const out: string[] = [];
  if (p.align) out.push(el("w:jc", { "w:val": JC[p.align] ?? "left" }));
  const sp: Record<string, number | string> = {};
  if (p.spaceBeforePx !== undefined) sp["w:before"] = pxToTwips(p.spaceBeforePx);
  if (p.spaceAfterPx !== undefined) sp["w:after"] = pxToTwips(p.spaceAfterPx);
  if (p.lineHeight !== undefined) {
    sp["w:line"] = multiplierToLine(p.lineHeight);
    sp["w:lineRule"] = "auto";
  }
  if (Object.keys(sp).length > 0) out.push(el("w:spacing", sp));
  const ind: Record<string, number> = {};
  if (p.indentLeftPx !== undefined) ind["w:left"] = pxToTwips(p.indentLeftPx);
  if (p.indentFirstLinePx !== undefined) {
    if (p.indentFirstLinePx >= 0) ind["w:firstLine"] = pxToTwips(p.indentFirstLinePx);
    else ind["w:hanging"] = pxToTwips(-p.indentFirstLinePx);
  }
  if (Object.keys(ind).length > 0) out.push(el("w:ind", ind));
  if (p.keepWithNext) out.push(el("w:keepNext"));
  if (p.keepLinesTogether) out.push(el("w:keepLines"));
  if (p.tabStops && p.tabStops.length > 0) {
    const tabs = p.tabStops
      .map((t) =>
        el("w:tab", {
          "w:val": t.align ? (TAB_VAL[t.align] ?? "left") : "left",
          "w:pos": pxToTwips(t.posPx),
          "w:leader": t.leader && t.leader !== "none" ? TAB_LEADER[t.leader] : undefined,
        }),
      )
      .join("");
    out.push(el("w:tabs", undefined, tabs));
  }
  return out.length > 0 ? el("w:pPr", undefined, out.join("")) : "";
}
