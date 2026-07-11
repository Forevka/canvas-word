// Decoders for w:rPr / w:pPr property bags — shared by documentParser.ts
// (direct formatting on runs/paragraphs) and styles.ts (the same bags appear
// inside w:style and w:docDefaults). Decode only; no resolution here.

import { decodeShdFill } from "./borders";
import type { IRLineNumbering, IRParaBorders, IRParaProps, IRRawBorder, IRRunProps } from "./types";
import { WarningSink } from "./types";
import { lineAutoToMultiplier } from "./units";
import { attr, el, els, indexChildren, numAttr, onOff, val, type XmlNode } from "./xml";

/** `val()` against a pre-built child index (see indexChildren). */
const bagVal = (bag: Map<string, XmlNode>, tagName: string): string | undefined => {
  const child = bag.get(tagName);
  return child ? attr(child, "w:val") : undefined;
};

export function decodeRunProps(rPr: XmlNode): IRRunProps {
  // ~26 tag probes per run bag — index the children once (O(1) per probe)
  // instead of a linear children scan per probe.
  const bag = indexChildren(rPr);
  const props: IRRunProps = {};
  const styleId = bagVal(bag, "w:rStyle");
  if (styleId) props.styleId = styleId;
  const bold = onOff(bag.get("w:b"));
  if (bold !== undefined) props.bold = bold;
  const italic = onOff(bag.get("w:i"));
  if (italic !== undefined) props.italic = italic;
  const strike = onOff(bag.get("w:strike"));
  if (strike !== undefined) props.strikethrough = strike;
  const u = bag.get("w:u");
  if (u) {
    const uVal = attr(u, "w:val");
    props.underline = uVal !== "none";
    // Carry the line style (anything other than a plain "single") + color so the
    // exporter/painter can reproduce double/dotted/wave/etc. and colored rules.
    if (props.underline && uVal && uVal !== "single") props.underlineStyle = uVal;
    const uColor = attr(u, "w:color");
    if (uColor) props.underlineColor = uColor;
    const uThemeColor = attr(u, "w:themeColor");
    if (uThemeColor) props.underlineColorTheme = uThemeColor;
  }
  const color = bag.get("w:color");
  if (color) {
    const hex = attr(color, "w:val");
    if (hex) props.color = hex;
    const themeColor = attr(color, "w:themeColor");
    if (themeColor) props.colorTheme = themeColor;
    const themeTint = attr(color, "w:themeTint");
    if (themeTint) props.colorThemeTint = themeTint;
    const themeShade = attr(color, "w:themeShade");
    if (themeShade) props.colorThemeShade = themeShade;
  }
  const sz = numAttr(bag.get("w:sz"), "w:val");
  if (sz !== undefined) props.sizeHalfPoints = sz;
  // w:spacing on a run is character tracking (twips); on a paragraph it is line
  // spacing — decodeParaProps reads the paragraph form separately.
  const spacing = numAttr(bag.get("w:spacing"), "w:val");
  if (spacing !== undefined) props.letterSpacingTwips = spacing;
  const rFonts = bag.get("w:rFonts");
  if (rFonts) {
    const font = attr(rFonts, "w:ascii");
    if (font) props.fontAscii = font;
    const themeFont = attr(rFonts, "w:asciiTheme");
    if (themeFont) props.fontThemeAscii = themeFont;
    const hAnsi = attr(rFonts, "w:hAnsi");
    if (hAnsi) props.fontHAnsi = hAnsi;
    const hAnsiTheme = attr(rFonts, "w:hAnsiTheme");
    if (hAnsiTheme) props.fontThemeHAnsi = hAnsiTheme;
    const cs = attr(rFonts, "w:cs");
    if (cs) props.fontCs = cs;
    const csTheme = attr(rFonts, "w:cstheme");
    if (csTheme) props.fontThemeCs = csTheme;
    const eastAsia = attr(rFonts, "w:eastAsia");
    if (eastAsia) props.fontEastAsia = eastAsia;
    const eastAsiaTheme = attr(rFonts, "w:eastAsiaTheme");
    if (eastAsiaTheme) props.fontThemeEastAsia = eastAsiaTheme;
  }
  // Tri-state (issue #155): "none" is an explicit clear (overrides a character
  // style's highlight), kept as null so mergeProps propagates it; absent = inherit.
  const highlight = bagVal(bag, "w:highlight");
  if (highlight === "none") props.highlight = null;
  else if (highlight) props.highlight = highlight;
  const vertAlign = bagVal(bag, "w:vertAlign");
  if (vertAlign) props.vertAlign = vertAlign;
  const vanish = onOff(bag.get("w:vanish"));
  if (vanish !== undefined) props.vanish = vanish;
  const caps = onOff(bag.get("w:caps"));
  if (caps !== undefined) props.caps = caps;
  const smallCaps = onOff(bag.get("w:smallCaps"));
  if (smallCaps !== undefined) props.smallCaps = smallCaps;
  const rtl = onOff(bag.get("w:rtl"));
  if (rtl !== undefined) props.rtl = rtl; // keep an explicit w:rtl="0" (clears inherited RTL)
  // Minor run typography & effects (w:rPr extras): double strike, baseline position,
  // kerning, character-width scaling, emphasis marks, the boolean text effects, a run
  // border, and fitText. Decode-only; mapToModel converts units / collapses borders.
  const dstrike = onOff(bag.get("w:dstrike"));
  if (dstrike !== undefined) props.doubleStrikethrough = dstrike;
  const position = numAttr(bag.get("w:position"), "w:val");
  if (position !== undefined) props.positionHalfPoints = position;
  const kern = numAttr(bag.get("w:kern"), "w:val");
  if (kern !== undefined) props.kerningHalfPoints = kern;
  const w = bag.get("w:w");
  if (w) {
    // w:w/@w:val is a percentage; older producers append "%". Parse the number out.
    const raw = attr(w, "w:val");
    if (raw !== undefined) {
      const pct = Number(raw.replace("%", ""));
      if (Number.isFinite(pct) && pct > 0) props.widthScalePct = pct;
    }
  }
  const em = bagVal(bag, "w:em");
  if (em && em !== "none") props.emphasisMark = em;
  const outline = onOff(bag.get("w:outline"));
  if (outline !== undefined) props.outline = outline;
  const shadow = onOff(bag.get("w:shadow"));
  if (shadow !== undefined) props.shadow = shadow;
  const emboss = onOff(bag.get("w:emboss"));
  if (emboss !== undefined) props.emboss = emboss;
  const imprint = onOff(bag.get("w:imprint"));
  if (imprint !== undefined) props.imprint = imprint;
  const snapToGrid = onOff(bag.get("w:snapToGrid"));
  if (snapToGrid !== undefined) props.snapToGrid = snapToGrid;
  const bdr = bag.get("w:bdr");
  if (bdr) {
    const raw: IRRawBorder = { val: attr(bdr, "w:val") ?? "single" };
    const sz = numAttr(bdr, "w:sz");
    if (sz !== undefined) raw.sizeEighthPt = sz;
    const color = attr(bdr, "w:color");
    if (color) raw.color = color;
    props.runBorder = raw;
  }
  const fitText = numAttr(bag.get("w:fitText"), "w:val");
  if (fitText !== undefined) props.fitTextTwips = fitText;
  return props;
}

const JC_MAP: Record<string, IRParaProps["align"]> = {
  left: "left",
  start: "left",
  center: "center",
  right: "right",
  end: "right",
  both: "justify",
  distribute: "justify",
};

export function decodeParaProps(pPr: XmlNode, warnings: WarningSink): IRParaProps {
  // Same one-pass child index as decodeRunProps — this bag is probed for ~20
  // tags and runs once per paragraph.
  const bag = indexChildren(pPr);
  const props: IRParaProps = {};
  const styleId = bagVal(bag, "w:pStyle");
  if (styleId) props.styleId = styleId;

  const jc = bagVal(bag, "w:jc");
  const align = jc !== undefined ? JC_MAP[jc] : undefined;
  if (align) props.align = align;

  const spacing = bag.get("w:spacing");
  if (spacing) {
    const before = numAttr(spacing, "w:before");
    if (before !== undefined) props.spaceBeforeTwips = before;
    const after = numAttr(spacing, "w:after");
    if (after !== undefined) props.spaceAfterTwips = after;
    // Auto-spacing (issue #160): when on, the explicit before/after is ignored and
    // Word computes the space itself. Kept as a flag; the mapper bakes a concrete px.
    // These are ATTRIBUTES of w:spacing (like w:before), not child on/off elements.
    const beforeAuto = attr(spacing, "w:beforeAutospacing");
    if (beforeAuto !== undefined) props.spaceBeforeAuto = beforeAuto !== "0" && beforeAuto !== "false";
    const afterAuto = attr(spacing, "w:afterAutospacing");
    if (afterAuto !== undefined) props.spaceAfterAuto = afterAuto !== "0" && afterAuto !== "false";
    const line = numAttr(spacing, "w:line");
    const rule = attr(spacing, "w:lineRule") ?? "auto";
    if (line !== undefined) {
      if (rule === "exact" || (rule === "atLeast" && line > 0)) {
        // Fixed point spacing: w:line is twips here, not 240ths.
        props.lineRule = rule;
        props.lineExactTwips = line;
      } else if (rule === "auto") {
        // Multiplier of the single line height. Record the explicit "auto" so it
        // overrides an inherited fixed rule through the cascade (mergeProps strips
        // undefined, so absence can't clear an inherited value).
        props.lineRule = "auto";
        props.lineHeight = lineAutoToMultiplier(line);
      }
      // else: a no-op atLeast=0 floor — leave line spacing untouched.
    }
  }

  const ind = bag.get("w:ind");
  if (ind) {
    const left = numAttr(ind, "w:left") ?? numAttr(ind, "w:start");
    if (left !== undefined) props.indentLeftTwips = left;
    const right = numAttr(ind, "w:right") ?? numAttr(ind, "w:end");
    if (right !== undefined) props.indentRightTwips = right;
    const firstLine = numAttr(ind, "w:firstLine");
    const hanging = numAttr(ind, "w:hanging");
    if (firstLine !== undefined) props.indentFirstLineTwips = firstLine;
    else if (hanging !== undefined) props.indentFirstLineTwips = -hanging;
  }

  const keepNext = onOff(bag.get("w:keepNext"));
  if (keepNext !== undefined) props.keepWithNext = keepNext;

  const keepLines = onOff(bag.get("w:keepLines"));
  if (keepLines !== undefined) props.keepLinesTogether = keepLines;

  const contextualSpacing = onOff(bag.get("w:contextualSpacing"));
  if (contextualSpacing !== undefined) props.contextualSpacing = contextualSpacing;

  const bidi = onOff(bag.get("w:bidi"));
  if (bidi !== undefined) props.direction = bidi ? "rtl" : "ltr"; // keep explicit w:bidi="0"

  const tabs = bag.get("w:tabs");
  if (tabs) {
    const stops: NonNullable<IRParaProps["tabStops"]> = [];
    for (const t of els(tabs, "w:tab")) {
      const pos = numAttr(t, "w:pos");
      const tabVal = attr(t, "w:val");
      // "bar" is a vertical rule, not a tab stop. "clear" is PRESERVED (issue #154):
      // it explicitly removes a tab an inherited paragraph style provides, so it must
      // round-trip instead of being silently dropped (the style tab would reappear).
      if (pos === undefined || tabVal === "bar") continue;
      const stop: { posTwips: number; val?: string; leader?: string } = { posTwips: pos };
      if (tabVal) stop.val = tabVal;
      const leader = attr(t, "w:leader");
      if (leader) stop.leader = leader;
      stops.push(stop);
    }
    if (stops.length > 0) props.tabStops = stops;
  }

  const pageBreakBefore = onOff(bag.get("w:pageBreakBefore"));
  if (pageBreakBefore !== undefined) props.pageBreakBefore = pageBreakBefore;

  // w:outlineLvl (0-8) — heading styles set it; absent = body text. Resolved
  // through the style cascade (mergeProps), so a heading paragraph inherits it.
  const outline = bagVal(bag, "w:outlineLvl");
  if (outline !== undefined) {
    const n = Number(outline);
    if (Number.isFinite(n) && n >= 0 && n <= 8) props.outlineLevel = n;
  }

  const numPr = bag.get("w:numPr");
  if (numPr) {
    const numId = val(numPr, "w:numId");
    if (numId !== undefined) {
      // numId 0 = Word's "remove numbering" sentinel (overrides an inherited list).
      props.list = numId === "0" ? null : { numId, level: numAttr(el(numPr, "w:ilvl"), "w:val") ?? 0 };
    }
  }

  // w:pBdr — paragraph borders (top/left/bottom/right/between/bar). Decoded raw
  // here; mapToModel collapses each edge to px (reusing the table border path).
  const pBdr = bag.get("w:pBdr");
  if (pBdr) {
    const borders: IRParaBorders = {};
    const edge = (key: keyof IRParaBorders, tag: string): void => {
      const e = el(pBdr, tag);
      if (!e) return;
      const raw: IRRawBorder = { val: attr(e, "w:val") ?? "single" };
      const sz = numAttr(e, "w:sz");
      if (sz !== undefined) raw.sizeEighthPt = sz;
      const color = attr(e, "w:color");
      if (color) raw.color = color;
      borders[key] = raw;
    };
    edge("top", "w:top");
    edge("left", "w:left");
    edge("bottom", "w:bottom");
    edge("right", "w:right");
    edge("between", "w:between");
    // Tri-state (issue #153): a w:pBdr element is present, so record `bordersSpecified`
    // and ALWAYS store the (possibly empty / all-nil) edge set. That defined-but-empty
    // object overrides an inherited style's borders through mergeProps — an empty w:pBdr
    // clears the style's box — where leaving `borders` undefined would let the style win.
    props.borders = borders;
    props.bordersSpecified = true;
  }

  // Paragraph-level w:shd (distinct from a run's or cell's shading).
  const shd = decodeShdFill(bag.get("w:shd"));
  if (shd !== undefined) props.shd = shd;

  // Minor paragraph props (issue #62): widow/orphan control, line-number
  // suppression, vertical line alignment, mirrored indents, right-indent adjust.
  const widow = onOff(bag.get("w:widowControl"));
  if (widow !== undefined) props.widowControl = widow;
  const suppressLn = onOff(bag.get("w:suppressLineNumbers"));
  if (suppressLn !== undefined) props.suppressLineNumbers = suppressLn;
  const textAlign = bagVal(bag, "w:textAlignment");
  if (textAlign === "top" || textAlign === "center" || textAlign === "bottom" || textAlign === "baseline") {
    props.textAlignment = textAlign;
  }
  const mirror = onOff(bag.get("w:mirrorIndents"));
  if (mirror !== undefined) props.mirrorIndents = mirror;
  const adjustRight = onOff(bag.get("w:adjustRightInd"));
  if (adjustRight !== undefined) props.adjustRightInd = adjustRight;
  // Unmodeled low-frequency CJK / hyphenation toggles (issue #161): round-trip-only,
  // no layout behavior. Each keeps an explicit w:val="0" so an OFF override survives.
  const snapToGrid = onOff(bag.get("w:snapToGrid"));
  if (snapToGrid !== undefined) props.snapToGrid = snapToGrid;
  const suppressAutoHyphens = onOff(bag.get("w:suppressAutoHyphens"));
  if (suppressAutoHyphens !== undefined) props.suppressAutoHyphens = suppressAutoHyphens;
  const kinsoku = onOff(bag.get("w:kinsoku"));
  if (kinsoku !== undefined) props.kinsoku = kinsoku;
  const overflowPunct = onOff(bag.get("w:overflowPunct"));
  if (overflowPunct !== undefined) props.overflowPunct = overflowPunct;
  const wordWrap = onOff(bag.get("w:wordWrap"));
  if (wordWrap !== undefined) props.wordWrap = wordWrap;
  const topLinePunct = onOff(bag.get("w:topLinePunct"));
  if (topLinePunct !== undefined) props.topLinePunct = topLinePunct;
  const autoSpaceDE = onOff(bag.get("w:autoSpaceDE"));
  if (autoSpaceDE !== undefined) props.autoSpaceDE = autoSpaceDE;
  const autoSpaceDN = onOff(bag.get("w:autoSpaceDN"));
  if (autoSpaceDN !== undefined) props.autoSpaceDN = autoSpaceDN;

  const rPr = bag.get("w:rPr");
  if (rPr) props.markRunProps = decodeRunProps(rPr);

  const sectPr = bag.get("w:sectPr");
  if (sectPr) {
    // Page geometry of non-last sections is still lossy (last wins), but the
    // page boundary the break implies IS respected via pageBreakBefore — when
    // the geometry actually changes (mapToModel compares sectionPgSize).
    const sectType = val(sectPr, "w:type");
    props.sectionBreak = sectType === "continuous" ? "continuous" : "page";
    if (sectType === "evenPage" || sectType === "oddPage") props.sectionBreakType = sectType;
    const lnNum = decodeLineNumbering(el(sectPr, "w:lnNumType"));
    if (lnNum) props.sectionLineNumbering = lnNum;
    const pgSz = el(sectPr, "w:pgSz");
    const w = numAttr(pgSz, "w:w");
    const h = numAttr(pgSz, "w:h");
    if (w !== undefined && h !== undefined) props.sectionPgSize = { w, h };
    const pgMar = el(sectPr, "w:pgMar");
    if (pgMar) {
      const top = numAttr(pgMar, "w:top");
      const right = numAttr(pgMar, "w:right");
      const bottom = numAttr(pgMar, "w:bottom");
      const left = numAttr(pgMar, "w:left");
      if (top !== undefined && right !== undefined && bottom !== undefined && left !== undefined) {
        props.sectionMarginTwips = { top, right, bottom, left };
      }
      const headerDist = numAttr(pgMar, "w:header");
      if (headerDist !== undefined) props.sectionHeaderDistTwips = headerDist;
      const footerDist = numAttr(pgMar, "w:footer");
      if (footerDist !== undefined) props.sectionFooterDistTwips = footerDist;
    }
    const cols = el(sectPr, "w:cols");
    if (cols) {
      const colEls = els(cols, "w:col");
      const colCount = numAttr(cols, "w:num") ?? (colEls.length > 1 ? colEls.length : 1);
      if (colCount > 1) {
        props.sectionColumns = { count: colCount };
        const space = numAttr(cols, "w:space");
        if (space !== undefined) props.sectionColumns.spaceTwips = space;
        if (attr(cols, "w:sep") === "1" || attr(cols, "w:sep") === "true") props.sectionColumns.sep = true;
        if (colEls.length === colCount) {
          const list = colEls.map((cel) => ({
            wTwips: numAttr(cel, "w:w") ?? 0,
            spaceTwips: numAttr(cel, "w:space") ?? 0,
          }));
          if (list.some((c) => c.wTwips > 0)) props.sectionColumns.cols = list;
        }
      }
    }
    const pgBorders = el(sectPr, "w:pgBorders");
    if (pgBorders) {
      const offsetFrom = attr(pgBorders, "w:offsetFrom");
      const borders: import("./types").IRPageBorders = { offsetFrom: offsetFrom === "text" ? "text" : "page" };
      const edge = (name: "top" | "right" | "bottom" | "left"): void => {
        const e = el(pgBorders, "w:" + name);
        if (!e) return;
        const sz = numAttr(e, "w:sz");
        const space = numAttr(e, "w:space");
        const color = attr(e, "w:color");
        borders[name] = {
          style: attr(e, "w:val") ?? "single",
          ...(sz !== undefined ? { sz } : {}),
          ...(space !== undefined ? { space } : {}),
          ...(color !== undefined ? { color } : {}),
        };
      };
      edge("top");
      edge("right");
      edge("bottom");
      edge("left");
      if (borders.top || borders.right || borders.bottom || borders.left) props.sectionPgBorders = borders;
    }
    const pgNumStart = numAttr(el(sectPr, "w:pgNumType"), "w:start");
    if (pgNumStart !== undefined) props.sectionPageNumberStart = pgNumStart;
    if (els(sectPr, "w:headerReference").length > 0 || els(sectPr, "w:footerReference").length > 0) {
      props.sectionHasBands = true;
    }
  }
  return props;
}

/** Decode w:lnNumType (line numbering) into raw OOXML units. Returns undefined
 *  when the element is absent so callers can leave the section unnumbered. */
export function decodeLineNumbering(lnNumType: XmlNode | undefined): IRLineNumbering | undefined {
  if (!lnNumType) return undefined;
  const out: IRLineNumbering = {};
  const countBy = numAttr(lnNumType, "w:countBy");
  if (countBy !== undefined) out.countBy = countBy;
  const start = numAttr(lnNumType, "w:start");
  if (start !== undefined) out.start = start;
  const distance = numAttr(lnNumType, "w:distance");
  if (distance !== undefined) out.distanceTwips = distance;
  const restart = attr(lnNumType, "w:restart");
  if (restart === "continuous" || restart === "newPage" || restart === "newSection") out.restart = restart;
  return out;
}
