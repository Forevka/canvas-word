// Run/paragraph property decoders. Pure: they take a parsed w:rPr/w:pPr node and
// return an IR props bag, with the WarningSink capturing lossy decisions. We focus
// on edge cases the roundtrip integration tests don't pin: on/off toggles,
// underline none, justification aliases, hanging vs firstLine indent, outline-level
// bounds, numId=0 sentinel, and fixed (exact/atLeast) line spacing.

import { describe, expect, it } from "vitest";
import { decodeParaProps, decodeRunProps } from "./props";
import { WarningSink } from "./types";
import { parseXml } from "./xml";

const rPr = (inner: string) => parseXml(`<w:rPr xmlns:w="urn">${inner}</w:rPr>`, "t")[0]!;
const pPr = (inner: string) => parseXml(`<w:pPr xmlns:w="urn">${inner}</w:pPr>`, "t")[0]!;

describe("decodeRunProps", () => {
  it("decodes a bare toggle as on", () => {
    expect(decodeRunProps(rPr(`<w:b/>`)).bold).toBe(true);
    expect(decodeRunProps(rPr(`<w:i/>`)).italic).toBe(true);
    expect(decodeRunProps(rPr(`<w:strike/>`)).strikethrough).toBe(true);
  });

  it("decodes w:caps and w:smallCaps toggles", () => {
    expect(decodeRunProps(rPr(`<w:caps/>`)).caps).toBe(true);
    expect(decodeRunProps(rPr(`<w:smallCaps/>`)).smallCaps).toBe(true);
    expect(decodeRunProps(rPr(`<w:caps w:val="0"/>`)).caps).toBe(false);
    expect(decodeRunProps(rPr(`<w:smallCaps w:val="false"/>`)).smallCaps).toBe(false);
  });

  it("decodes a toggle turned off via w:val='0'/'false'", () => {
    expect(decodeRunProps(rPr(`<w:b w:val="0"/>`)).bold).toBe(false);
    expect(decodeRunProps(rPr(`<w:i w:val="false"/>`)).italic).toBe(false);
  });

  it("treats w:u w:val='none' as no underline, anything else as underline", () => {
    expect(decodeRunProps(rPr(`<w:u w:val="none"/>`)).underline).toBe(false);
    expect(decodeRunProps(rPr(`<w:u w:val="single"/>`)).underline).toBe(true);
  });

  it("captures the underline line style (but not a plain 'single')", () => {
    expect(decodeRunProps(rPr(`<w:u w:val="double"/>`)).underlineStyle).toBe("double");
    expect(decodeRunProps(rPr(`<w:u w:val="wavyHeavy"/>`)).underlineStyle).toBe("wavyHeavy");
    expect(decodeRunProps(rPr(`<w:u w:val="single"/>`)).underlineStyle).toBeUndefined();
    expect(decodeRunProps(rPr(`<w:u w:val="none"/>`)).underlineStyle).toBeUndefined();
  });

  it("captures underline color and theme color", () => {
    const p = decodeRunProps(rPr(`<w:u w:val="dash" w:color="FF0000" w:themeColor="accent1"/>`));
    expect(p.underlineColor).toBe("FF0000");
    expect(p.underlineColorTheme).toBe("accent1");
  });

  it("captures color hex and theme color", () => {
    const p = decodeRunProps(rPr(`<w:color w:val="FF0000" w:themeColor="accent1"/>`));
    expect(p.color).toBe("FF0000");
    expect(p.colorTheme).toBe("accent1");
  });

  it("captures font size in half-points", () => {
    expect(decodeRunProps(rPr(`<w:sz w:val="24"/>`)).sizeHalfPoints).toBe(24);
  });

  it("captures rFonts ascii and asciiTheme", () => {
    const p = decodeRunProps(rPr(`<w:rFonts w:ascii="Calibri" w:asciiTheme="minorHAnsi"/>`));
    expect(p.fontAscii).toBe("Calibri");
    expect(p.fontThemeAscii).toBe("minorHAnsi");
  });

  it("captures the complex-script and East-Asian rFonts slots (+ themes)", () => {
    const p = decodeRunProps(
      rPr(`<w:rFonts w:hAnsi="Calibri" w:cs="Arial" w:eastAsia="SimSun" w:eastAsiaTheme="minorEastAsia"/>`),
    );
    expect(p.fontHAnsi).toBe("Calibri");
    expect(p.fontCs).toBe("Arial");
    expect(p.fontEastAsia).toBe("SimSun");
    expect(p.fontThemeEastAsia).toBe("minorEastAsia");
  });

  it("captures run-level character tracking (w:spacing)", () => {
    expect(decodeRunProps(rPr(`<w:spacing w:val="40"/>`)).letterSpacingTwips).toBe(40);
    expect(decodeRunProps(rPr(`<w:spacing w:val="-20"/>`)).letterSpacingTwips).toBe(-20);
  });

  it("captures theme tint/shade alongside the theme color", () => {
    const tinted = decodeRunProps(rPr(`<w:color w:themeColor="accent1" w:themeTint="99"/>`));
    expect(tinted.colorTheme).toBe("accent1");
    expect(tinted.colorThemeTint).toBe("99");
    const shaded = decodeRunProps(rPr(`<w:color w:themeColor="accent1" w:themeShade="80"/>`));
    expect(shaded.colorThemeShade).toBe("80");
  });

  it("ignores highlight='none' but keeps a real highlight", () => {
    expect(decodeRunProps(rPr(`<w:highlight w:val="none"/>`)).highlight).toBeUndefined();
    expect(decodeRunProps(rPr(`<w:highlight w:val="yellow"/>`)).highlight).toBe("yellow");
  });

  it("returns an empty bag for empty rPr", () => {
    expect(decodeRunProps(rPr(``))).toEqual({});
  });

  it("captures the rStyle reference and vertAlign", () => {
    const p = decodeRunProps(rPr(`<w:rStyle w:val="Emphasis"/><w:vertAlign w:val="superscript"/>`));
    expect(p.styleId).toBe("Emphasis");
    expect(p.vertAlign).toBe("superscript");
  });

  it("decodes double strikethrough (w:dstrike) and the text-effect toggles", () => {
    const p = decodeRunProps(rPr(`<w:dstrike/><w:outline/><w:shadow/><w:emboss/><w:imprint/>`));
    expect(p.doubleStrikethrough).toBe(true);
    expect(p.outline).toBe(true);
    expect(p.shadow).toBe(true);
    expect(p.emboss).toBe(true);
    expect(p.imprint).toBe(true);
  });

  it("decodes signed baseline position and kerning in half-points", () => {
    expect(decodeRunProps(rPr(`<w:position w:val="6"/>`)).positionHalfPoints).toBe(6);
    expect(decodeRunProps(rPr(`<w:position w:val="-6"/>`)).positionHalfPoints).toBe(-6);
    expect(decodeRunProps(rPr(`<w:kern w:val="18"/>`)).kerningHalfPoints).toBe(18);
  });

  it("decodes character width scaling (w:w), tolerating a trailing '%'", () => {
    expect(decodeRunProps(rPr(`<w:w w:val="150"/>`)).widthScalePct).toBe(150);
    expect(decodeRunProps(rPr(`<w:w w:val="66%"/>`)).widthScalePct).toBe(66);
  });

  it("decodes emphasis marks (w:em), dropping 'none'", () => {
    expect(decodeRunProps(rPr(`<w:em w:val="dot"/>`)).emphasisMark).toBe("dot");
    expect(decodeRunProps(rPr(`<w:em w:val="none"/>`)).emphasisMark).toBeUndefined();
  });

  it("decodes a run border (w:bdr) and fitText width (w:fitText)", () => {
    const p = decodeRunProps(rPr(`<w:bdr w:val="dashed" w:sz="12" w:color="1A73E8"/><w:fitText w:val="900"/>`));
    expect(p.runBorder).toEqual({ val: "dashed", sizeEighthPt: 12, color: "1A73E8" });
    expect(p.fitTextTwips).toBe(900);
  });
});

describe("decodeParaProps", () => {
  const decode = (inner: string) => {
    const warnings = new WarningSink();
    const props = decodeParaProps(pPr(inner), warnings);
    return { props, warnings };
  };

  it("maps justification aliases to model alignment", () => {
    expect(decode(`<w:jc w:val="start"/>`).props.align).toBe("left");
    expect(decode(`<w:jc w:val="end"/>`).props.align).toBe("right");
    expect(decode(`<w:jc w:val="both"/>`).props.align).toBe("justify");
    expect(decode(`<w:jc w:val="distribute"/>`).props.align).toBe("justify");
    expect(decode(`<w:jc w:val="center"/>`).props.align).toBe("center");
  });

  it("ignores an unknown jc value", () => {
    expect(decode(`<w:jc w:val="bogus"/>`).props.align).toBeUndefined();
  });

  it("decodes auto line spacing to a multiplier", () => {
    expect(decode(`<w:spacing w:line="360" w:lineRule="auto"/>`).props.lineHeight).toBe(1.5);
  });

  it("defaults absent lineRule to auto (multiplier)", () => {
    expect(decode(`<w:spacing w:line="240"/>`).props.lineHeight).toBe(1);
  });

  it("decodes exact line spacing as fixed twips (no warning, no multiplier)", () => {
    const { props, warnings } = decode(`<w:spacing w:line="420" w:lineRule="exact"/>`);
    expect(props.lineHeight).toBeUndefined();
    expect(props.lineRule).toBe("exact");
    expect(props.lineExactTwips).toBe(420);
    expect(warnings.list).toEqual([]);
  });

  it("decodes atLeast line spacing as fixed twips", () => {
    const { props } = decode(`<w:spacing w:line="360" w:lineRule="atLeast"/>`);
    expect(props.lineRule).toBe("atLeast");
    expect(props.lineExactTwips).toBe(360);
  });

  it("ignores a zero atLeast floor (no-op)", () => {
    const { props } = decode(`<w:spacing w:line="0" w:lineRule="atLeast"/>`);
    expect(props.lineRule).toBeUndefined();
    expect(props.lineExactTwips).toBeUndefined();
  });

  it("captures before/after spacing in twips", () => {
    const p = decode(`<w:spacing w:before="120" w:after="240"/>`).props;
    expect(p.spaceBeforeTwips).toBe(120);
    expect(p.spaceAfterTwips).toBe(240);
  });

  it("prefers w:start/w:end over absent w:left/w:right for indents", () => {
    const p = decode(`<w:ind w:start="720" w:end="360"/>`).props;
    expect(p.indentLeftTwips).toBe(720);
    expect(p.indentRightTwips).toBe(360);
  });

  it("records a hanging indent as a NEGATIVE first-line indent", () => {
    expect(decode(`<w:ind w:hanging="360"/>`).props.indentFirstLineTwips).toBe(-360);
  });

  it("prefers firstLine over hanging when both are present", () => {
    expect(decode(`<w:ind w:firstLine="240" w:hanging="360"/>`).props.indentFirstLineTwips).toBe(240);
  });

  it("accepts an in-range outline level and rejects out-of-range", () => {
    expect(decode(`<w:outlineLvl w:val="0"/>`).props.outlineLevel).toBe(0);
    expect(decode(`<w:outlineLvl w:val="8"/>`).props.outlineLevel).toBe(8);
    expect(decode(`<w:outlineLvl w:val="9"/>`).props.outlineLevel).toBeUndefined();
    expect(decode(`<w:outlineLvl w:val="-1"/>`).props.outlineLevel).toBeUndefined();
  });

  it("maps a real list reference to numId+level", () => {
    const p = decode(`<w:numPr><w:numId w:val="3"/><w:ilvl w:val="2"/></w:numPr>`).props;
    expect(p.list).toEqual({ numId: "3", level: 2 });
  });

  it("defaults the list level to 0 when w:ilvl is absent", () => {
    expect(decode(`<w:numPr><w:numId w:val="3"/></w:numPr>`).props.list).toEqual({
      numId: "3",
      level: 0,
    });
  });

  it("treats numId=0 as the remove-numbering sentinel (null)", () => {
    expect(decode(`<w:numPr><w:numId w:val="0"/></w:numPr>`).props.list).toBeNull();
  });

  it("decodes keepNext / keepLines / pageBreakBefore toggles", () => {
    const p = decode(`<w:keepNext/><w:keepLines/><w:pageBreakBefore/>`).props;
    expect(p.keepWithNext).toBe(true);
    expect(p.keepLinesTogether).toBe(true);
    expect(p.pageBreakBefore).toBe(true);
  });

  it("decodes w:contextualSpacing (and its explicit off)", () => {
    expect(decode(`<w:contextualSpacing/>`).props.contextualSpacing).toBe(true);
    expect(decode(`<w:contextualSpacing w:val="0"/>`).props.contextualSpacing).toBe(false);
    expect(decode(``).props.contextualSpacing).toBeUndefined();
  });

  it("collects tab stops, dropping 'clear' and 'bar' entries", () => {
    const p = decode(
      `<w:tabs>` +
        `<w:tab w:val="left" w:pos="720"/>` +
        `<w:tab w:val="clear" w:pos="1440"/>` +
        `<w:tab w:val="bar" w:pos="2160"/>` +
        `<w:tab w:val="right" w:pos="2880" w:leader="dot"/>` +
        `</w:tabs>`,
    ).props;
    expect(p.tabStops).toEqual([
      { posTwips: 720, val: "left" },
      { posTwips: 2880, val: "right", leader: "dot" },
    ]);
  });

  it("decodes mark run props from a paragraph-mark rPr", () => {
    const p = decode(`<w:rPr><w:b/></w:rPr>`).props;
    expect(p.markRunProps).toEqual({ bold: true });
  });
});
