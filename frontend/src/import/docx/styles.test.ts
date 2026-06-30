// StyleResolver tests: cascade order, basedOn chains, toggle-XOR semantics,
// theme indirection, docDefaults — exercised end-to-end through the pipeline
// (real zip → rels → styles.xml → resolution) so the wiring is covered too.

import { describe, expect, it } from "vitest";
import type { Paragraph } from "@cw/shared";
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

describe("StyleResolver — theme tint/shade", () => {
  // accent1 = 4472C4 (68,114,196). Tint 0x99 (0.6): c·0.6 + 255·0.4. Shade 0x80
  // (≈0.502): c·0.502. Both applied per channel (the linear w:color interpretation).
  it("lightens a theme color by w:themeTint instead of flattening to the base hue", () => {
    const r = importStyled(
      `<w:p><w:r><w:rPr><w:color w:themeColor="accent1" w:themeTint="99"/></w:rPr><w:t>tint</w:t></w:r></w:p>`,
    );
    expect(para(r.doc.blocks[0]).runs[0]!.style.color).toBe("#8faadc");
  });

  it("darkens a theme color by w:themeShade", () => {
    const r = importStyled(
      `<w:p><w:r><w:rPr><w:color w:themeColor="accent1" w:themeShade="80"/></w:rPr><w:t>shade</w:t></w:r></w:p>`,
    );
    expect(para(r.doc.blocks[0]).runs[0]!.style.color).toBe("#223962");
  });

  it("keeps the concrete w:val when present (Word pre-bakes the tinted hex there)", () => {
    const r = importStyled(
      `<w:p><w:r><w:rPr><w:color w:val="9CC3E5" w:themeColor="accent1" w:themeTint="99"/></w:rPr><w:t>x</w:t></w:r></w:p>`,
    );
    expect(para(r.doc.blocks[0]).runs[0]!.style.color).toBe("#9cc3e5");
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

describe("StyleResolver — heading-led section breaks", () => {
  // A geometry-preserving "page" section break breaks only when the new section
  // opens with a heading (a real chapter boundary), else it flows — past the
  // empty/hidden paragraphs the generator inserts ahead of the heading.
  it("page-breaks a section break that starts a heading section", () => {
    const r = importStyled(
      `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p>` +
        `<w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr><w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>CHAPTER TWO</w:t></w:r></w:p>`,
    );
    const heading = r.doc.blocks.find((b) => para(b).runs.some((run) => run.text === "CHAPTER TWO"))!;
    expect(para(heading).style.pageBreakBefore).toBe(true);
  });

  it("flows a section break followed by ordinary content", () => {
    const r = importStyled(
      `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p>` +
        `<w:p><w:r><w:t>just more body text</w:t></w:r></w:p>`,
    );
    const body = r.doc.blocks.find((b) => para(b).runs.some((run) => run.text === "just more body text"))!;
    expect(para(body).style.pageBreakBefore).toBeUndefined();
  });

  it("treats a heading-NAMED custom style (e.g. \"TOC Heading Custom\") as a heading", () => {
    const styles = stylesPartXml(
      `<w:style w:type="paragraph" w:styleId="TocH"><w:name w:val="TOC Heading Custom"/></w:style>
       <w:style w:type="paragraph" w:styleId="TocE"><w:name w:val="toc 1"/></w:style>`,
    );
    const sect = `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p>`;
    const tocTitle = runImport(styledDocx(`${sect}<w:p><w:pPr><w:pStyle w:val="TocH"/></w:pPr><w:r><w:t>TABLE OF CONTENTS</w:t></w:r></w:p>`, styles));
    expect(para(tocTitle.doc.blocks.find((b) => para(b).runs.some((r) => r.text === "TABLE OF CONTENTS"))!).style.pageBreakBefore).toBe(true);
    // the entry style "toc 1" must NOT count as a heading (entries don't break)
    const tocEntry = runImport(styledDocx(`${sect}<w:p><w:pPr><w:pStyle w:val="TocE"/></w:pPr><w:r><w:t>An entry</w:t></w:r></w:p>`, styles));
    expect(para(tocEntry.doc.blocks.find((b) => para(b).runs.some((r) => r.text === "An entry"))!).style.pageBreakBefore).toBeUndefined();
  });

  it("page-breaks a heading section sitting behind a continuous break (full-page map)", () => {
    // Real appraisal-report shape: ZONING (nextPage) → a full-page Zoning Map section
    // (CONTINUOUS — flows on the same page) → HIGHEST AND BEST USE (nextPage). A
    // sectPr describes the section it CLOSES, so HIGHEST's new page comes from the
    // sectPr AFTER it; the map's continuous break must not swallow it.
    const sect = `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p>`;
    const cont = `<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p>`;
    const h1 = (t: string): string => `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
    const r = importStyled(
      `<w:p><w:r><w:t>COVER</w:t></w:r></w:p>` +
        sect +
        h1("ZONING") +
        sect +
        `<w:p><w:r><w:t>Zoning Map</w:t></w:r></w:p>` +
        cont +
        h1("HIGHEST AND BEST USE") +
        sect +
        `<w:p><w:r><w:t>tail</w:t></w:r></w:p>`,
    );
    const find = (t: string): Paragraph => para(r.doc.blocks.find((b) => para(b).runs.some((rn) => rn.text === t))!);
    expect(find("ZONING").style.pageBreakBefore).toBe(true);
    expect(find("Zoning Map").style.pageBreakBefore).toBeUndefined(); // continuous → flows
    expect(find("HIGHEST AND BEST USE").style.pageBreakBefore).toBe(true); // regression: was undefined
  });
});

describe("StyleResolver — fixed line spacing cascade (w:lineRule)", () => {
  // A paragraph style with EXACT 24pt (480 twips) line spacing, plus a child
  // based on it — so we can check both inheritance and explicit override-to-auto.
  const FIXED_STYLES = stylesPartXml(
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
     <w:style w:type="paragraph" w:styleId="Tight">
       <w:basedOn w:val="Normal"/>
       <w:pPr><w:spacing w:line="480" w:lineRule="exact"/></w:pPr>
     </w:style>`,
  );
  const importFixed = (body: string): Paragraph =>
    para(runImport(styledDocx(body, FIXED_STYLES)).doc.blocks[0]);

  it("inherits a paragraph style's exact spacing", () => {
    const p = importFixed(`<w:p><w:pPr><w:pStyle w:val="Tight"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(p.style.lineRule).toBe("exact");
    expect(p.style.lineHeightPx).toBeCloseTo(32, 1); // 480 twips = 32px
  });

  it("an explicit direct w:lineRule=\"auto\" clears the inherited fixed rule", () => {
    const p = importFixed(
      `<w:p><w:pPr><w:pStyle w:val="Tight"/><w:spacing w:line="360" w:lineRule="auto"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
    );
    expect(p.style.lineRule).toBeUndefined(); // reverted to multiplier
    expect(p.style.lineHeightPx).toBeUndefined();
    expect(p.style.lineHeight).toBeCloseTo(1.5, 2); // 360/240
  });
});

const round = (v: number): number => Math.round(v * 100) / 100;
