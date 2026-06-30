// Decoders for w:rPr / w:pPr property bags — shared by documentParser.ts
// (direct formatting on runs/paragraphs) and styles.ts (the same bags appear
// inside w:style and w:docDefaults). Decode only; no resolution here.

import { decodeShdFill } from "./borders";
import type { IRLineNumbering, IRParaBorders, IRParaProps, IRRawBorder, IRRunProps } from "./types";
import { WarningSink } from "./types";
import { lineAutoToMultiplier } from "./units";
import { attr, el, els, numAttr, onOff, val, type XmlNode } from "./xml";

export function decodeRunProps(rPr: XmlNode): IRRunProps {
  const props: IRRunProps = {};
  const styleId = val(rPr, "w:rStyle");
  if (styleId) props.styleId = styleId;
  const bold = onOff(el(rPr, "w:b"));
  if (bold !== undefined) props.bold = bold;
  const italic = onOff(el(rPr, "w:i"));
  if (italic !== undefined) props.italic = italic;
  const strike = onOff(el(rPr, "w:strike"));
  if (strike !== undefined) props.strikethrough = strike;
  const u = el(rPr, "w:u");
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
  const color = el(rPr, "w:color");
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
  const sz = numAttr(el(rPr, "w:sz"), "w:val");
  if (sz !== undefined) props.sizeHalfPoints = sz;
  // w:spacing on a run is character tracking (twips); on a paragraph it is line
  // spacing — decodeParaProps reads the paragraph form separately.
  const spacing = numAttr(el(rPr, "w:spacing"), "w:val");
  if (spacing !== undefined) props.letterSpacingTwips = spacing;
  const rFonts = el(rPr, "w:rFonts");
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
  const highlight = val(rPr, "w:highlight");
  if (highlight && highlight !== "none") props.highlight = highlight;
  const vertAlign = val(rPr, "w:vertAlign");
  if (vertAlign) props.vertAlign = vertAlign;
  const vanish = onOff(el(rPr, "w:vanish"));
  if (vanish !== undefined) props.vanish = vanish;
  const caps = onOff(el(rPr, "w:caps"));
  if (caps !== undefined) props.caps = caps;
  const smallCaps = onOff(el(rPr, "w:smallCaps"));
  if (smallCaps !== undefined) props.smallCaps = smallCaps;
  const rtl = onOff(el(rPr, "w:rtl"));
  if (rtl !== undefined) props.rtl = rtl; // keep an explicit w:rtl="0" (clears inherited RTL)
  // Minor run typography & effects (w:rPr extras): double strike, baseline position,
  // kerning, character-width scaling, emphasis marks, the boolean text effects, a run
  // border, and fitText. Decode-only; mapToModel converts units / collapses borders.
  const dstrike = onOff(el(rPr, "w:dstrike"));
  if (dstrike !== undefined) props.doubleStrikethrough = dstrike;
  const position = numAttr(el(rPr, "w:position"), "w:val");
  if (position !== undefined) props.positionHalfPoints = position;
  const kern = numAttr(el(rPr, "w:kern"), "w:val");
  if (kern !== undefined) props.kerningHalfPoints = kern;
  const w = el(rPr, "w:w");
  if (w) {
    // w:w/@w:val is a percentage; older producers append "%". Parse the number out.
    const raw = attr(w, "w:val");
    if (raw !== undefined) {
      const pct = Number(raw.replace("%", ""));
      if (Number.isFinite(pct) && pct > 0) props.widthScalePct = pct;
    }
  }
  const em = val(rPr, "w:em");
  if (em && em !== "none") props.emphasisMark = em;
  const outline = onOff(el(rPr, "w:outline"));
  if (outline !== undefined) props.outline = outline;
  const shadow = onOff(el(rPr, "w:shadow"));
  if (shadow !== undefined) props.shadow = shadow;
  const emboss = onOff(el(rPr, "w:emboss"));
  if (emboss !== undefined) props.emboss = emboss;
  const imprint = onOff(el(rPr, "w:imprint"));
  if (imprint !== undefined) props.imprint = imprint;
  const bdr = el(rPr, "w:bdr");
  if (bdr) {
    const raw: IRRawBorder = { val: attr(bdr, "w:val") ?? "single" };
    const sz = numAttr(bdr, "w:sz");
    if (sz !== undefined) raw.sizeEighthPt = sz;
    const color = attr(bdr, "w:color");
    if (color) raw.color = color;
    props.runBorder = raw;
  }
  const fitText = numAttr(el(rPr, "w:fitText"), "w:val");
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
  const props: IRParaProps = {};
  const styleId = val(pPr, "w:pStyle");
  if (styleId) props.styleId = styleId;

  const jc = val(pPr, "w:jc");
  const align = jc !== undefined ? JC_MAP[jc] : undefined;
  if (align) props.align = align;

  const spacing = el(pPr, "w:spacing");
  if (spacing) {
    const before = numAttr(spacing, "w:before");
    if (before !== undefined) props.spaceBeforeTwips = before;
    const after = numAttr(spacing, "w:after");
    if (after !== undefined) props.spaceAfterTwips = after;
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

  const ind = el(pPr, "w:ind");
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

  const keepNext = onOff(el(pPr, "w:keepNext"));
  if (keepNext !== undefined) props.keepWithNext = keepNext;

  const keepLines = onOff(el(pPr, "w:keepLines"));
  if (keepLines !== undefined) props.keepLinesTogether = keepLines;

  const contextualSpacing = onOff(el(pPr, "w:contextualSpacing"));
  if (contextualSpacing !== undefined) props.contextualSpacing = contextualSpacing;

  const bidi = onOff(el(pPr, "w:bidi"));
  if (bidi !== undefined) props.direction = bidi ? "rtl" : "ltr"; // keep explicit w:bidi="0"

  const tabs = el(pPr, "w:tabs");
  if (tabs) {
    const stops: NonNullable<IRParaProps["tabStops"]> = [];
    for (const t of els(tabs, "w:tab")) {
      const pos = numAttr(t, "w:pos");
      const tabVal = attr(t, "w:val");
      // "clear" removes a stop; "bar" is a vertical rule, not a tab stop.
      if (pos === undefined || tabVal === "clear" || tabVal === "bar") continue;
      const stop: { posTwips: number; val?: string; leader?: string } = { posTwips: pos };
      if (tabVal) stop.val = tabVal;
      const leader = attr(t, "w:leader");
      if (leader) stop.leader = leader;
      stops.push(stop);
    }
    if (stops.length > 0) props.tabStops = stops;
  }

  const pageBreakBefore = onOff(el(pPr, "w:pageBreakBefore"));
  if (pageBreakBefore !== undefined) props.pageBreakBefore = pageBreakBefore;

  // w:outlineLvl (0-8) — heading styles set it; absent = body text. Resolved
  // through the style cascade (mergeProps), so a heading paragraph inherits it.
  const outline = val(pPr, "w:outlineLvl");
  if (outline !== undefined) {
    const n = Number(outline);
    if (Number.isFinite(n) && n >= 0 && n <= 8) props.outlineLevel = n;
  }

  const numPr = el(pPr, "w:numPr");
  if (numPr) {
    const numId = val(numPr, "w:numId");
    if (numId !== undefined) {
      // numId 0 = Word's "remove numbering" sentinel (overrides an inherited list).
      props.list = numId === "0" ? null : { numId, level: numAttr(el(numPr, "w:ilvl"), "w:val") ?? 0 };
    }
  }

  // w:pBdr — paragraph borders (top/left/bottom/right/between/bar). Decoded raw
  // here; mapToModel collapses each edge to px (reusing the table border path).
  const pBdr = el(pPr, "w:pBdr");
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
    if (borders.top || borders.left || borders.bottom || borders.right || borders.between) {
      props.borders = borders;
    }
  }

  // Paragraph-level w:shd (distinct from a run's or cell's shading).
  const shd = decodeShdFill(el(pPr, "w:shd"));
  if (shd !== undefined) props.shd = shd;

  // Minor paragraph props (issue #62): widow/orphan control, line-number
  // suppression, vertical line alignment, mirrored indents, right-indent adjust.
  const widow = onOff(el(pPr, "w:widowControl"));
  if (widow !== undefined) props.widowControl = widow;
  const suppressLn = onOff(el(pPr, "w:suppressLineNumbers"));
  if (suppressLn !== undefined) props.suppressLineNumbers = suppressLn;
  const textAlign = val(pPr, "w:textAlignment");
  if (textAlign === "top" || textAlign === "center" || textAlign === "bottom" || textAlign === "baseline") {
    props.textAlignment = textAlign;
  }
  const mirror = onOff(el(pPr, "w:mirrorIndents"));
  if (mirror !== undefined) props.mirrorIndents = mirror;
  const adjustRight = onOff(el(pPr, "w:adjustRightInd"));
  if (adjustRight !== undefined) props.adjustRightInd = adjustRight;

  const rPr = el(pPr, "w:rPr");
  if (rPr) props.markRunProps = decodeRunProps(rPr);

  const sectPr = el(pPr, "w:sectPr");
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
