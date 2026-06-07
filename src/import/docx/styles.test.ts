// StyleResolver tests: cascade order, basedOn chains, toggle-XOR semantics,
// theme indirection, docDefaults — exercised end-to-end through the pipeline
// (real zip → rels → styles.xml → resolution) so the wiring is covered too.

import { describe, expect, it } from "vitest";
import type { Paragraph } from "../../model/document";
import { runImport } from "./pipeline";
import { stylesPartXml, styledDocx } from "./fixture";

// Word-like styles.xml: docDefaults (theme body font, 11pt, 1.08 line), Normal
// as the w:default paragraph style, Heading1 based on Normal, Strong char style.
const STYLES = stylesPartXml(
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
     <w:name w:val="Normal"/>
     <w:pPr><w:ind w:left="720"/></w:pPr>
   </w:style>
   <w:style w:type="paragraph" w:styleId="Heading1">
     <w:basedOn w:val="Normal"/>
     <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="0"/></w:pPr>
     <w:rPr><w:rFonts w:asciiTheme="majorHAnsi"/><w:b/><w:sz w:val="32"/><w:color w:val="2F5496" w:themeColor="accent1"/></w:rPr>
   </w:style>
   <w:style w:type="paragraph" w:styleId="Heading2">
     <w:basedOn w:val="Heading1"/>
     <w:rPr><w:sz w:val="26"/></w:rPr>
   </w:style>
   <w:style w:type="character" w:styleId="Strong">
     <w:rPr><w:b/></w:rPr>
   </w:style>
   <w:style w:type="character" w:styleId="Accent">
     <w:rPr><w:color w:themeColor="accent2"/></w:rPr>
   </w:style>`,
  `<w:docDefaults>
     <w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
     <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259"/></w:pPr></w:pPrDefault>
   </w:docDefaults>`,
);

const importStyled = (body: string) => runImport(styledDocx(body, STYLES));

const para = (b: unknown): Paragraph => {
  expect((b as Paragraph).kind).toBe("paragraph");
  return b as Paragraph;
};

describe("StyleResolver — docDefaults and default style", () => {
  it("applies docDefaults run props (theme font + size) to plain runs", () => {
    const r = importStyled(`<w:p><w:r><w:t>plain</w:t></w:r></w:p>`);
    const run = para(r.doc.blocks[0]).runs[0]!;
    expect(run.style.fontFamily).toBe("Calibri, serif"); // minorHAnsi via theme
    expect(run.style.fontSizePx).toBeCloseTo(14.67, 2); // 11pt
  });

  it("applies docDefaults paragraph props (spacing, line height)", () => {
    const r = importStyled(`<w:p><w:r><w:t>plain</w:t></w:r></w:p>`);
    const style = para(r.doc.blocks[0]).style;
    expect(style.spaceAfterPx).toBeCloseTo(10.67, 2); // 160 twips
    expect(style.lineHeight).toBeCloseTo(1.08, 2); // 259/240
  });

  it("applies the w:default paragraph style to paragraphs without pStyle", () => {
    const r = importStyled(`<w:p><w:r><w:t>plain</w:t></w:r></w:p>`);
    expect(para(r.doc.blocks[0]).style.indentLeftPx).toBe(48); // Normal's 720 twips
  });

  it("styles EMPTY paragraphs with the document's resolved defaults, not the editor's", () => {
    // Users leave empty paragraphs as intentional spacing — their height in
    // Word comes from the document default font/size, which must survive.
    const r = importStyled(`<w:p/>`);
    const run = para(r.doc.blocks[0]).runs[0]!;
    expect(run.text).toBe("");
    expect(run.style.fontFamily).toBe("Calibri, serif"); // minorHAnsi via docDefaults
    expect(run.style.fontSizePx).toBeCloseTo(14.67, 2); // docDefaults 11pt
  });

  it("respects the paragraph mark's own formatting (w:pPr/w:rPr) on empty paragraphs", () => {
    const r = importStyled(`<w:p><w:pPr><w:rPr><w:sz w:val="48"/></w:rPr></w:pPr></w:p>`);
    expect(para(r.doc.blocks[0]).runs[0]!.style.fontSizePx).toBe(32); // 24pt mark
  });

  it("styles empty paragraphs through their paragraph style too", () => {
    const r = importStyled(`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:p>`);
    const run = para(r.doc.blocks[0]).runs[0]!;
    expect(run.style.bold).toBe(true);
    expect(run.style.fontFamily).toBe("Calibri Light, serif"); // majorHAnsi
  });
});

describe("StyleResolver — style chains", () => {
  it("resolves a named paragraph style (run + para props)", () => {
    const r = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>title</w:t></w:r></w:p>`,
    );
    const p = para(r.doc.blocks[0]);
    expect(p.runs[0]!.style).toMatchObject({
      bold: true,
      fontFamily: "Calibri Light, serif", // majorHAnsi
      fontSizePx: round(32 / 1.5), // 21.33 — 16pt
      color: "#2f5496", // concrete w:val wins over themeColor
    });
    expect(p.style).toMatchObject({
      keepWithNext: true,
      spaceBeforePx: 16,
      spaceAfterPx: 0, // Heading1 overrides docDefaults' 160
      indentLeftPx: 48, // inherited from Normal via basedOn
      namedStyle: "Heading1",
    });
  });

  it("inherits through basedOn chains with child overrides", () => {
    const r = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>sub</w:t></w:r></w:p>`,
    );
    const run = para(r.doc.blocks[0]).runs[0]!;
    expect(run.style.bold).toBe(true); // from Heading1
    expect(run.style.fontSizePx).toBeCloseTo(17.33, 2); // Heading2's own 13pt
  });

  it("applies character styles via w:rStyle", () => {
    const r = importStyled(
      `<w:p><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>strong</w:t></w:r></w:p>`,
    );
    expect(para(r.doc.blocks[0]).runs[0]!.style.bold).toBe(true);
  });

  it("resolves theme-only colors (w:themeColor without w:val)", () => {
    const r = importStyled(
      `<w:p><w:r><w:rPr><w:rStyle w:val="Accent"/></w:rPr><w:t>orange</w:t></w:r></w:p>`,
    );
    expect(para(r.doc.blocks[0]).runs[0]!.style.color).toBe("#ed7d31"); // accent2
  });
});

describe("StyleResolver — toggle XOR (§17.7.3)", () => {
  it("paragraph-style bold XOR character-style bold cancels", () => {
    const r = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>` +
        `<w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>not bold</w:t></w:r></w:p>`,
    );
    expect(para(r.doc.blocks[0]).runs[0]!.style.bold).toBe(false);
  });

  it("direct formatting is absolute, exempt from XOR", () => {
    const r = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>` +
        `<w:r><w:rPr><w:b/></w:rPr><w:t>still bold</w:t></w:r></w:p>`,
    );
    expect(para(r.doc.blocks[0]).runs[0]!.style.bold).toBe(true);
  });

  it("direct off-toggle wins over style-driven bold", () => {
    const r = importStyled(
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>` +
        `<w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>unbold</w:t></w:r></w:p>`,
    );
    expect(para(r.doc.blocks[0]).runs[0]!.style.bold).toBe(false);
  });
});

describe("StyleResolver — robustness", () => {
  it("survives basedOn cycles", () => {
    const cyclic = stylesPartXml(
      `<w:style w:type="paragraph" w:styleId="A"><w:basedOn w:val="B"/><w:rPr><w:b/></w:rPr></w:style>
       <w:style w:type="paragraph" w:styleId="B"><w:basedOn w:val="A"/><w:rPr><w:i/></w:rPr></w:style>`,
    );
    const r = runImport(styledDocx(`<w:p><w:pPr><w:pStyle w:val="A"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`, cyclic));
    const run = para(r.doc.blocks[0]).runs[0]!;
    expect(run.style.bold).toBe(true); // A's own props still apply
  });

  it("falls back to Word's bare-document defaults without styles.xml", () => {
    const r = runImport(styledDocx(`<w:p><w:r><w:t>x</w:t></w:r></w:p>`));
    const p = para(r.doc.blocks[0]);
    const run = p.runs[0]!;
    expect(run.style.fontFamily).toBe("Times New Roman, serif");
    expect(run.style.fontSizePx).toBe(16); // 12pt
    expect(p.style.lineHeight).toBe(1); // single — Word's default, not the editor's
    expect(p.style.spaceAfterPx).toBe(0);
  });
});

const round = (v: number): number => Math.round(v * 100) / 100;
