// Decoders for w:rPr / w:pPr property bags — shared by documentParser.ts
// (direct formatting on runs/paragraphs) and styles.ts (the same bags appear
// inside w:style and w:docDefaults). Decode only; no resolution here.

import type { IRParaProps, IRRunProps } from "./types";
import { WarningSink } from "./types";
import { lineAutoToMultiplier } from "./units";
import { attr, el, numAttr, onOff, val, type XmlNode } from "./xml";

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
  if (u) props.underline = attr(u, "w:val") !== "none";
  const color = el(rPr, "w:color");
  if (color) {
    const hex = attr(color, "w:val");
    if (hex) props.color = hex;
    const themeColor = attr(color, "w:themeColor");
    if (themeColor) props.colorTheme = themeColor;
  }
  const sz = numAttr(el(rPr, "w:sz"), "w:val");
  if (sz !== undefined) props.sizeHalfPoints = sz;
  const rFonts = el(rPr, "w:rFonts");
  if (rFonts) {
    const font = attr(rFonts, "w:ascii");
    if (font) props.fontAscii = font;
    const themeFont = attr(rFonts, "w:asciiTheme");
    if (themeFont) props.fontThemeAscii = themeFont;
  }
  const highlight = val(rPr, "w:highlight");
  if (highlight && highlight !== "none") props.highlight = highlight;
  const vertAlign = val(rPr, "w:vertAlign");
  if (vertAlign) props.vertAlign = vertAlign;
  const vanish = onOff(el(rPr, "w:vanish"));
  if (vanish !== undefined) props.vanish = vanish;
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
      if (rule === "auto") props.lineHeight = lineAutoToMultiplier(line);
      else warnings.add("line-rule-exact", "Exact/atLeast line spacing was ignored (model is multiplier-only).");
    }
  }

  const ind = el(pPr, "w:ind");
  if (ind) {
    const left = numAttr(ind, "w:left") ?? numAttr(ind, "w:start");
    if (left !== undefined) props.indentLeftTwips = left;
    const firstLine = numAttr(ind, "w:firstLine");
    const hanging = numAttr(ind, "w:hanging");
    if (firstLine !== undefined) props.indentFirstLineTwips = firstLine;
    else if (hanging !== undefined) props.indentFirstLineTwips = -hanging;
  }

  const keepNext = onOff(el(pPr, "w:keepNext"));
  if (keepNext !== undefined) props.keepWithNext = keepNext;

  const keepLines = onOff(el(pPr, "w:keepLines"));
  if (keepLines !== undefined) props.keepLinesTogether = keepLines;

  const pageBreakBefore = onOff(el(pPr, "w:pageBreakBefore"));
  if (pageBreakBefore !== undefined) props.pageBreakBefore = pageBreakBefore;

  const numPr = el(pPr, "w:numPr");
  if (numPr) {
    const numId = val(numPr, "w:numId");
    if (numId !== undefined) {
      // numId 0 = Word's "remove numbering" sentinel (overrides an inherited list).
      props.list = numId === "0" ? null : { numId, level: numAttr(el(numPr, "w:ilvl"), "w:val") ?? 0 };
    }
  }

  const rPr = el(pPr, "w:rPr");
  if (rPr) props.markRunProps = decodeRunProps(rPr);

  const sectPr = el(pPr, "w:sectPr");
  if (sectPr) {
    // Page geometry of non-last sections is still lossy (last wins), but the
    // page boundary the break implies IS respected via pageBreakBefore.
    warnings.add("multiple-sections", "Multiple sections found — only the last section's page setup is used.");
    props.sectionBreak = val(sectPr, "w:type") === "continuous" ? "continuous" : "page";
  }
  return props;
}
