// Partial rPr/pPr emitters for style-gallery patches (NamedStyle.char/.para) and
// list-marker run props. Only DEFINED fields are written — the importer reads
// each w:* element into a patch field, so absent elements stay absent.

import type { CellBorders, CharStyle, ParaStyle, TableCond, TableCondProps, TableStyle } from "@cw/shared";
import { multiplierToLine, pxToEighthPoints, pxToHalfPoints, pxToTwips } from "../units";
import { HIGHLIGHT_NAME, hexColor as hex, JC, TAB_LEADER, TAB_VAL } from "./mappings";
import { el } from "./xmlWrite";

/** w:u attributes for a run: the line style (w:val) + optional color. underline
 *  off ⇒ w:val="none". Mirrors the importer (props.ts) and painter (paintStyle). */
function underlineAttrs(c: Pick<CharStyle, "underline" | "underlineStyle" | "underlineColor">): Record<string, string> {
  if (!c.underline) return { "w:val": "none" };
  const attrs: Record<string, string> = { "w:val": c.underlineStyle ?? "single" };
  if (c.underlineColor) attrs["w:color"] = hex(c.underlineColor);
  return attrs;
}

// Full rPr serializer: every toggle written explicitly (on/off) so a run's
// direct formatting fully overrides any inherited paragraph/named style on
// re-import. Shared by the main exporter (documentXml) and the headless TOC
// generator (recalc/generateTocDocx) so they can't drift.
export function runPropsXml(s: CharStyle): string {
  const family = (s.fontFamily.split(",")[0] ?? "").trim();
  const children: string[] = [];
  // w:rStyle (character-style reference) must be the first child of w:rPr; the
  // concrete props below still override it on re-import (Word semantics).
  if (s.charStyleId) children.push(el("w:rStyle", { "w:val": s.charStyleId }));
  if (family) children.push(el("w:rFonts", { "w:ascii": family, "w:hAnsi": family, "w:cs": family }));
  children.push(el("w:b", { "w:val": s.bold ? "1" : "0" }));
  children.push(el("w:i", { "w:val": s.italic ? "1" : "0" }));
  if (s.strikethrough) children.push(el("w:strike", { "w:val": "1" }));
  if (s.hidden) children.push(el("w:vanish", { "w:val": "1" })); // preserved hidden text
  children.push(el("w:color", { "w:val": hex(s.color) }));
  children.push(el("w:sz", { "w:val": pxToHalfPoints(s.fontSizePx) }));
  children.push(el("w:szCs", { "w:val": pxToHalfPoints(s.fontSizePx) }));
  children.push(el("w:u", underlineAttrs(s)));
  if (s.letterSpacingPx !== undefined) children.push(el("w:spacing", { "w:val": pxToTwips(s.letterSpacingPx) }));
  if (s.highlightColor) {
    const name = HIGHLIGHT_NAME[s.highlightColor.toLowerCase()];
    if (name) children.push(el("w:highlight", { "w:val": name }));
  }
  if (s.verticalAlign) children.push(el("w:vertAlign", { "w:val": s.verticalAlign === "super" ? "superscript" : "subscript" }));
  // Emit an explicit OFF (w:val="0") so an imported w:rtl="0" round-trips and can
  // still clear an inherited RTL run style; undefined stays absent.
  if (s.rtl === true) children.push(el("w:rtl"));
  else if (s.rtl === false) children.push(el("w:rtl", { "w:val": "0" }));
  return el("w:rPr", undefined, children.join(""));
}

// The spacing/indent/jc/tabs core shared by every full pPr: returned as
// concatenated child elements (NOT wrapped in w:pPr) so the main exporter can
// interleave it with pStyle/numPr/breaks/section props, while the TOC generator
// wraps it directly. Emission order matches OOXML's tolerant exporter convention.
export function paraCoreXml(style: ParaStyle): string {
  const c: string[] = [];
  // w:bidi precedes w:spacing/w:ind/w:jc in the CT_PPr schema sequence. An explicit
  // "ltr" emits w:bidi="0" so an imported w:bidi="0" round-trips (clears inherited
  // RTL); undefined stays absent.
  if (style.direction === "rtl") c.push(el("w:bidi"));
  else if (style.direction === "ltr") c.push(el("w:bidi", { "w:val": "0" }));
  c.push(el("w:spacing", {
    "w:before": pxToTwips(style.spaceBeforePx),
    "w:after": pxToTwips(style.spaceAfterPx),
    "w:line": multiplierToLine(style.lineHeight),
    "w:lineRule": "auto",
  }));
  const ind: Record<string, number> = {};
  if (style.indentLeftPx) ind["w:left"] = pxToTwips(style.indentLeftPx);
  if (style.indentRightPx) ind["w:right"] = pxToTwips(style.indentRightPx);
  if (style.indentFirstLinePx > 0) ind["w:firstLine"] = pxToTwips(style.indentFirstLinePx);
  else if (style.indentFirstLinePx < 0) ind["w:hanging"] = pxToTwips(-style.indentFirstLinePx);
  if (Object.keys(ind).length > 0) c.push(el("w:ind", ind));
  c.push(el("w:jc", { "w:val": JC[style.align] }));
  if (style.tabStops && style.tabStops.length > 0) {
    const tabs = style.tabStops
      .map((t) =>
        el("w:tab", {
          "w:val": t.align ? (TAB_VAL[t.align] ?? "left") : "left",
          "w:pos": pxToTwips(t.posPx),
          "w:leader": t.leader && t.leader !== "none" ? TAB_LEADER[t.leader] : undefined,
        }),
      )
      .join("");
    c.push(el("w:tabs", undefined, tabs));
  }
  return c.join("");
}

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
  if (c.underline !== undefined) out.push(el("w:u", underlineAttrs(c as Pick<CharStyle, "underline" | "underlineStyle" | "underlineColor">)));
  if (c.highlightColor) {
    const name = HIGHLIGHT_NAME[c.highlightColor.toLowerCase()];
    if (name) out.push(el("w:highlight", { "w:val": name }));
  }
  if (c.verticalAlign) out.push(el("w:vertAlign", { "w:val": c.verticalAlign === "super" ? "superscript" : "subscript" }));
  return out.length > 0 ? el("w:rPr", undefined, out.join("")) : "";
}

// w:pBdr for a paragraph-style patch (mirrors documentXml.paraBordersXml).
function paraBordersXml(b: NonNullable<ParaStyle["borders"]>): string {
  const edge = (name: string, spec: CellBorders["top"]): string => {
    if (!spec) return "";
    const val = spec.style === "double" ? "double" : spec.style === "dashed" ? "dashed" : spec.style === "dotted" ? "dotted" : "single";
    return el("w:" + name, { "w:val": val, "w:sz": pxToEighthPoints(spec.widthPx), "w:space": 0, "w:color": hex(spec.color) });
  };
  const inner =
    edge("top", b.top) + edge("left", b.left) + edge("bottom", b.bottom) + edge("right", b.right) + edge("between", b.between);
  return inner ? el("w:pBdr", undefined, inner) : "";
}

export function partialPPrXml(p: Partial<ParaStyle>): string {
  const out: string[] = [];
  if (p.direction === "rtl") out.push(el("w:bidi"));
  else if (p.direction === "ltr") out.push(el("w:bidi", { "w:val": "0" }));
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
  if (p.indentRightPx !== undefined) ind["w:right"] = pxToTwips(p.indentRightPx);
  if (p.indentFirstLinePx !== undefined) {
    if (p.indentFirstLinePx >= 0) ind["w:firstLine"] = pxToTwips(p.indentFirstLinePx);
    else ind["w:hanging"] = pxToTwips(-p.indentFirstLinePx);
  }
  if (Object.keys(ind).length > 0) out.push(el("w:ind", ind));
  if (p.keepWithNext) out.push(el("w:keepNext"));
  if (p.keepLinesTogether) out.push(el("w:keepLines"));
  if (p.borders) out.push(paraBordersXml(p.borders));
  if (p.shading) out.push(shdFillXml(p.shading));
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

// ---------------------------------------------------------------------------
// Table styles → w:style[@w:type="table"] with conditional formatting.

function condBordersXml(tag: string, b: CellBorders): string {
  const edge = (name: string, spec: CellBorders["top"]): string => {
    if (!spec) return "";
    const val = spec.style === "double" ? "double" : spec.style === "dashed" ? "dashed" : spec.style === "dotted" ? "dotted" : "single";
    return el("w:" + name, { "w:val": val, "w:sz": pxToEighthPoints(spec.widthPx), "w:space": 0, "w:color": hex(spec.color) });
  };
  return el(tag, undefined, edge("top", b.top) + edge("left", b.left) + edge("bottom", b.bottom) + edge("right", b.right));
}

const shdFillXml = (fill: string): string => el("w:shd", { "w:val": "clear", "w:color": "auto", "w:fill": hex(fill) });

function tblStylePrXml(cond: TableCond, props: TableCondProps): string {
  const tc: string[] = [];
  if (props.borders) tc.push(condBordersXml("w:tcBorders", props.borders));
  if (props.shading) tc.push(shdFillXml(props.shading));
  const tcPr = tc.length > 0 ? el("w:tcPr", undefined, tc.join("")) : "";
  return el("w:tblStylePr", { "w:type": cond }, partialRPrXml(props.char ?? {}) + partialPPrXml(props.para ?? {}) + tcPr);
}

/** One table style. Base (wholeTable) → rPr/pPr + tblPr(borders/shd/band sizes);
 *  each other condition → a w:tblStylePr. */
export function tableStyleXml(style: TableStyle): string {
  const base = style.conds.wholeTable ?? {};
  const tblPrKids: string[] = [];
  if (style.rowBandSize !== undefined) tblPrKids.push(el("w:tblStyleRowBandSize", { "w:val": style.rowBandSize }));
  if (style.colBandSize !== undefined) tblPrKids.push(el("w:tblStyleColBandSize", { "w:val": style.colBandSize }));
  if (base.borders) tblPrKids.push(condBordersXml("w:tblBorders", base.borders));
  if (base.shading) tblPrKids.push(shdFillXml(base.shading));
  const tblPr = tblPrKids.length > 0 ? el("w:tblPr", undefined, tblPrKids.join("")) : "";
  const conds = (Object.entries(style.conds) as [TableCond, TableCondProps][])
    .filter(([c]) => c !== "wholeTable")
    .map(([c, p]) => tblStylePrXml(c, p))
    .join("");
  const body =
    el("w:name", { "w:val": style.name }) +
    (style.basedOn ? el("w:basedOn", { "w:val": style.basedOn }) : "") +
    partialRPrXml(base.char ?? {}) +
    partialPPrXml(base.para ?? {}) +
    tblPr +
    conds;
  return el("w:style", { "w:type": "table", "w:styleId": style.id }, body);
}
