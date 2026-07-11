// Partial rPr/pPr emitters for style-gallery patches (NamedStyle.char/.para) and
// list-marker run props. Only DEFINED fields are written — the importer reads
// each w:* element into a patch field, so absent elements stay absent.

import type { CellBorders, CharStyle, ParaStyle, TableCond, TableCondProps, TableStyle, TabStop } from "@cw/shared";
import { multiplierToLine, pxToHalfPoints, pxToTwips } from "../units";
import { borderEdgeXml, HIGHLIGHT_NAME, hexColor as hex, JC, TAB_LEADER, TAB_VAL } from "./mappings";
import { el } from "./xmlWrite";

/** w:u attributes for a run: the line style (w:val) + optional color. underline
 *  off ⇒ w:val="none". Mirrors the importer (props.ts) and painter (paintStyle). */
function underlineAttrs(c: Pick<CharStyle, "underline" | "underlineStyle" | "underlineColor">): Record<string, string> {
  if (!c.underline) return { "w:val": "none" };
  const attrs: Record<string, string> = { "w:val": c.underlineStyle ?? "single" };
  if (c.underlineColor) attrs["w:color"] = hex(c.underlineColor);
  return attrs;
}

/** w:bdr element for a run border (the shared border-edge encoding). */
const runBorderXml = (b: NonNullable<CharStyle["runBorder"]>): string => borderEdgeXml("bdr", b);

/** Minor run typography & effects (OOXML w:rPr extras) common to the full and
 *  partial rPr serializers: double strike, the boolean text effects, character
 *  width scaling, kerning, baseline position, run border, fitText, and emphasis
 *  marks. Only DEFINED/truthy fields are emitted, so a run without them serializes
 *  exactly as before (no drift). */
function runEffectsXml(s: Partial<CharStyle>): string {
  const out: string[] = [];
  // Emit explicit on/off (w:val="1"/"0") whenever DEFINED — so a style patch's
  // `false` can clear an inherited w:rStyle value, mirroring w:b/w:i/w:strike.
  // Concrete runs leave these undefined unless set, so output stays drift-free.
  if (s.doubleStrikethrough !== undefined) out.push(el("w:dstrike", { "w:val": s.doubleStrikethrough ? "1" : "0" }));
  if (s.outline !== undefined) out.push(el("w:outline", { "w:val": s.outline ? "1" : "0" }));
  if (s.shadow !== undefined) out.push(el("w:shadow", { "w:val": s.shadow ? "1" : "0" }));
  if (s.emboss !== undefined) out.push(el("w:emboss", { "w:val": s.emboss ? "1" : "0" }));
  if (s.imprint !== undefined) out.push(el("w:imprint", { "w:val": s.imprint ? "1" : "0" }));
  if (s.snapToGrid !== undefined) out.push(el("w:snapToGrid", { "w:val": s.snapToGrid ? "1" : "0" }));
  // w:lang (issue #168) — emit whichever proofing-language attrs are set. Shared by
  // concrete runs and style/docDefaults rPr patches, so a run's/style's language
  // round-trips. Absent when the lang object has no attributes.
  if (s.lang) {
    const langAttrs: Record<string, string> = {};
    if (s.lang.val) langAttrs["w:val"] = s.lang.val;
    if (s.lang.eastAsia) langAttrs["w:eastAsia"] = s.lang.eastAsia;
    if (s.lang.bidi) langAttrs["w:bidi"] = s.lang.bidi;
    if (Object.keys(langAttrs).length > 0) out.push(el("w:lang", langAttrs));
  }
  if (s.widthScalePct !== undefined) out.push(el("w:w", { "w:val": s.widthScalePct }));
  if (s.kerningMinPx !== undefined) out.push(el("w:kern", { "w:val": pxToHalfPoints(s.kerningMinPx) }));
  if (s.positionPx !== undefined) out.push(el("w:position", { "w:val": pxToHalfPoints(s.positionPx) }));
  if (s.runBorder) out.push(runBorderXml(s.runBorder));
  if (s.fitTextPx !== undefined) out.push(el("w:fitText", { "w:val": pxToTwips(s.fitTextPx) }));
  if (s.emphasisMark) out.push(el("w:em", { "w:val": s.emphasisMark }));
  return out.join("");
}

// Full rPr serializer: every toggle written explicitly (on/off) so a run's
// direct formatting fully overrides any inherited paragraph/named style on
// re-import. Shared by the main exporter (documentXml) and the headless TOC
// generator (recalc/generateTocDocx) so they can't drift.
// First family of a CSS font stack ("Georgia, serif" → "Georgia"); "" if absent.
const firstFamily = (stack: string | undefined): string => (stack ? (stack.split(",")[0] ?? "").trim() : "");

/** A w:rFonts element covering the Latin (ascii/hAnsi), complex-script (cs) and
 *  East-Asian (eastAsia) slots, or "" when no slot is set. cs defaults to the
 *  Latin family (Word's convention) so a plain run keeps round-tripping its w:cs. */
function rFontsXml(fontFamily: string | undefined, complexScript?: string, eastAsia?: string): string {
  const latin = firstFamily(fontFamily);
  const cs = firstFamily(complexScript) || latin;
  const ea = firstFamily(eastAsia);
  if (!latin && !cs && !ea) return "";
  const attrs: Record<string, string> = {};
  if (latin) {
    attrs["w:ascii"] = latin;
    attrs["w:hAnsi"] = latin;
  }
  if (cs) attrs["w:cs"] = cs;
  if (ea) attrs["w:eastAsia"] = ea;
  return el("w:rFonts", attrs);
}

export function runPropsXml(s: CharStyle): string {
  const children: string[] = [];
  // w:rStyle (character-style reference) must be the first child of w:rPr; the
  // concrete props below still override it on re-import (Word semantics).
  if (s.charStyleId) children.push(el("w:rStyle", { "w:val": s.charStyleId }));
  const rFonts = rFontsXml(s.fontFamily, s.fontFamilyComplexScript, s.fontFamilyEastAsia);
  if (rFonts) children.push(rFonts);
  children.push(el("w:b", { "w:val": s.bold ? "1" : "0" }));
  children.push(el("w:i", { "w:val": s.italic ? "1" : "0" }));
  // w:caps / w:smallCaps precede w:strike in the CT_RPr schema sequence. Emit an
  // explicit OFF (w:val="0") for a false value so it clears an inherited (rStyle/
  // style-cascade) all-caps/small-caps on re-import; undefined stays absent.
  if (s.caps !== undefined) children.push(el("w:caps", { "w:val": s.caps ? "1" : "0" }));
  if (s.smallCaps !== undefined) children.push(el("w:smallCaps", { "w:val": s.smallCaps ? "1" : "0" }));
  // Always emit an explicit on/off (like w:b/w:i above) so a strike CLEARED against a
  // striking character style survives re-import instead of the style's strike returning
  // via toggle XOR (issue #155); undefined stays absent.
  children.push(el("w:strike", { "w:val": s.strikethrough ? "1" : "0" }));
  if (s.hidden) children.push(el("w:vanish", { "w:val": "1" })); // preserved hidden text

  children.push(el("w:color", { "w:val": hex(s.color) }));
  children.push(el("w:sz", { "w:val": pxToHalfPoints(s.fontSizePx) }));
  children.push(el("w:szCs", { "w:val": pxToHalfPoints(s.fontSizePx) }));
  children.push(el("w:u", underlineAttrs(s)));
  if (s.letterSpacingPx !== undefined) children.push(el("w:spacing", { "w:val": pxToTwips(s.letterSpacingPx) }));
  if (s.highlightColor) {
    const name = HIGHLIGHT_NAME[s.highlightColor.toLowerCase()];
    if (name) children.push(el("w:highlight", { "w:val": name }));
  } else if (s.highlightCleared) {
    // Explicit clear (issue #155) — beats a highlight the run's w:rStyle carries.
    children.push(el("w:highlight", { "w:val": "none" }));
  }
  if (s.verticalAlign) children.push(el("w:vertAlign", { "w:val": s.verticalAlign === "super" ? "superscript" : "subscript" }));
  const effects = runEffectsXml(s);
  if (effects) children.push(effects);
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
  // Fixed (exact/atLeast) spacing emits w:line in twips, but only with a height —
  // a rule without lineHeightPx would serialize w:line="0"; fall back to the
  // multiplier (240ths, lineRule="auto") instead.
  const fixed = style.lineRule !== undefined && style.lineHeightPx !== undefined;
  const spacingAttrs: Record<string, string | number> = {
    "w:before": pxToTwips(style.spaceBeforePx),
    "w:after": pxToTwips(style.spaceAfterPx),
    "w:line": fixed ? pxToTwips(style.lineHeightPx!) : multiplierToLine(style.lineHeight),
    "w:lineRule": fixed ? style.lineRule! : "auto",
  };
  // Auto-spacing (issue #160): re-emit the autospacing attributes; Word ignores the
  // explicit before/after when these are on, but we emit both (as Word does).
  if (style.spaceBeforeAuto) spacingAttrs["w:beforeAutospacing"] = "1";
  if (style.spaceAfterAuto) spacingAttrs["w:afterAutospacing"] = "1";
  c.push(el("w:spacing", spacingAttrs));
  const ind: Record<string, number> = {};
  if (style.indentLeftPx) ind["w:left"] = pxToTwips(style.indentLeftPx);
  if (style.indentRightPx) ind["w:right"] = pxToTwips(style.indentRightPx);
  if (style.indentFirstLinePx > 0) ind["w:firstLine"] = pxToTwips(style.indentFirstLinePx);
  else if (style.indentFirstLinePx < 0) ind["w:hanging"] = pxToTwips(-style.indentFirstLinePx);
  if (Object.keys(ind).length > 0) c.push(el("w:ind", ind));
  // w:contextualSpacing follows w:ind in the CT_PPr schema sequence. Emit an
  // explicit OFF (w:val="0") for `false` so it can clear an inherited style's
  // suppression on re-import; undefined stays absent.
  if (style.contextualSpacing === true) c.push(el("w:contextualSpacing"));
  else if (style.contextualSpacing === false) c.push(el("w:contextualSpacing", { "w:val": "0" }));
  c.push(el("w:jc", { "w:val": JC[style.align] }));
  if (style.tabStops && style.tabStops.length > 0) {
    c.push(el("w:tabs", undefined, style.tabStops.map(tabXml).join("")));
  }
  return c.join("");
}

/** One w:tab element — a real stop, or an explicit clear (issue #154) that re-emits
 *  `w:val="clear"` (position only) so a removed style tab stays removed on re-import. */
function tabXml(t: TabStop): string {
  if (t.cleared) return el("w:tab", { "w:val": "clear", "w:pos": pxToTwips(t.posPx) });
  return el("w:tab", {
    "w:val": t.align ? (TAB_VAL[t.align] ?? "left") : "left",
    "w:pos": pxToTwips(t.posPx),
    "w:leader": t.leader && t.leader !== "none" ? TAB_LEADER[t.leader] : undefined,
  });
}

export function partialRPrXml(c: Partial<CharStyle>): string {
  const out: string[] = [];
  if (c.fontFamily || c.fontFamilyComplexScript || c.fontFamilyEastAsia) {
    const rFonts = rFontsXml(c.fontFamily, c.fontFamilyComplexScript, c.fontFamilyEastAsia);
    if (rFonts) out.push(rFonts);
  }
  if (c.bold !== undefined) out.push(el("w:b", { "w:val": c.bold ? "1" : "0" }));
  if (c.italic !== undefined) out.push(el("w:i", { "w:val": c.italic ? "1" : "0" }));
  if (c.caps !== undefined) out.push(el("w:caps", { "w:val": c.caps ? "1" : "0" }));
  if (c.smallCaps !== undefined) out.push(el("w:smallCaps", { "w:val": c.smallCaps ? "1" : "0" }));
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
  const effects = runEffectsXml(c);
  if (effects) out.push(effects);
  return out.length > 0 ? el("w:rPr", undefined, out.join("")) : "";
}

// w:pBdr — paragraph borders. Only edges that are set are written (no "explicit
// empty" semantics — a paragraph with no borders omits w:pBdr). Shared by the
// full pPr emitter (documentXml) and the style-patch emitter below.
export function paraBordersXml(b: NonNullable<ParaStyle["borders"]>): string {
  const inner =
    borderEdgeXml("top", b.top) +
    borderEdgeXml("left", b.left) +
    borderEdgeXml("bottom", b.bottom) +
    borderEdgeXml("right", b.right) +
    borderEdgeXml("between", b.between);
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
  if (p.spaceBeforeAuto) sp["w:beforeAutospacing"] = "1"; // issue #160
  if (p.spaceAfterAuto) sp["w:afterAutospacing"] = "1";
  if (p.lineRule !== undefined && p.lineHeightPx !== undefined) {
    sp["w:line"] = pxToTwips(p.lineHeightPx);
    sp["w:lineRule"] = p.lineRule;
  } else if (p.lineHeight !== undefined) {
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
  if (p.contextualSpacing === true) out.push(el("w:contextualSpacing"));
  else if (p.contextualSpacing === false) out.push(el("w:contextualSpacing", { "w:val": "0" }));
  if (p.keepWithNext) out.push(el("w:keepNext"));
  if (p.keepLinesTogether) out.push(el("w:keepLines"));
  // Minor paragraph props (issue #62) — see documentXml.pPrXml. In a style patch
  // `false` is a real override, so emit each on/off element whenever it is DEFINED
  // (false → w:val="0") to keep the style contract lossless.
  if (p.widowControl !== undefined) out.push(el("w:widowControl", p.widowControl ? undefined : { "w:val": "0" }));
  if (p.suppressLineNumbers !== undefined) out.push(el("w:suppressLineNumbers", p.suppressLineNumbers ? undefined : { "w:val": "0" }));
  if (p.mirrorIndents !== undefined) out.push(el("w:mirrorIndents", p.mirrorIndents ? undefined : { "w:val": "0" }));
  if (p.adjustRightInd !== undefined) out.push(el("w:adjustRightInd", p.adjustRightInd ? undefined : { "w:val": "0" }));
  // Unmodeled CJK / hyphenation toggles (issue #161) — round-trip-only, same on/off encoding.
  if (p.snapToGrid !== undefined) out.push(el("w:snapToGrid", p.snapToGrid ? undefined : { "w:val": "0" }));
  if (p.suppressAutoHyphens !== undefined) out.push(el("w:suppressAutoHyphens", p.suppressAutoHyphens ? undefined : { "w:val": "0" }));
  if (p.kinsoku !== undefined) out.push(el("w:kinsoku", p.kinsoku ? undefined : { "w:val": "0" }));
  if (p.overflowPunct !== undefined) out.push(el("w:overflowPunct", p.overflowPunct ? undefined : { "w:val": "0" }));
  if (p.wordWrap !== undefined) out.push(el("w:wordWrap", p.wordWrap ? undefined : { "w:val": "0" }));
  if (p.topLinePunct !== undefined) out.push(el("w:topLinePunct", p.topLinePunct ? undefined : { "w:val": "0" }));
  if (p.autoSpaceDE !== undefined) out.push(el("w:autoSpaceDE", p.autoSpaceDE ? undefined : { "w:val": "0" }));
  if (p.autoSpaceDN !== undefined) out.push(el("w:autoSpaceDN", p.autoSpaceDN ? undefined : { "w:val": "0" }));
  if (p.textAlignment) out.push(el("w:textAlignment", { "w:val": p.textAlignment }));
  if (p.borders) out.push(paraBordersXml(p.borders));
  else if (p.bordersCleared) out.push(el("w:pBdr")); // issue #153: explicit "no box" delta
  if (p.shading) out.push(shdFillXml(p.shading));
  else if (p.shadingCleared) out.push(shdClearXml()); // issue #147: explicit "No Color" delta
  if (p.tabStops && p.tabStops.length > 0) {
    out.push(el("w:tabs", undefined, p.tabStops.map(tabXml).join("")));
  }
  return out.length > 0 ? el("w:pPr", undefined, out.join("")) : "";
}

// ---------------------------------------------------------------------------
// Table styles → w:style[@w:type="table"] with conditional formatting.

function condBordersXml(tag: string, b: CellBorders): string {
  return el(
    tag,
    undefined,
    borderEdgeXml("top", b.top) + borderEdgeXml("left", b.left) + borderEdgeXml("bottom", b.bottom) + borderEdgeXml("right", b.right),
  );
}

/** The w:shd fill element (clear pattern + explicit fill) — the one encoding of a
 *  model shading color, shared with documentXml's paragraph/cell/table emitters. */
export const shdFillXml = (fill: string): string => el("w:shd", { "w:val": "clear", "w:color": "auto", "w:fill": hex(fill) });

/** The empty w:shd element (issue #147) — Word's "Shading → No Color". Emitted for
 *  a `shadingCleared` paragraph/style so an explicit clear survives re-import
 *  instead of the named style's inherited fill silently returning. */
export const shdClearXml = (): string => el("w:shd", { "w:val": "clear", "w:color": "auto", "w:fill": "auto" });

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
