// Underline style + color (OOXML w:u/@w:val + @w:color) must survive a full .docx
// round-trip: export emits the real w:val (not always "single") and the optional
// color; re-import rebuilds CharStyle.underlineStyle + underlineColor. Plus an
// import-only check that a themed underline color (w:u/@w:themeColor) resolves
// through theme1.xml, mirroring the run-color theme path.
import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, Run, SectionProps, UnderlineStyle } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { styledDocx } from "../../import/docx/fixture";
import { writeDocx } from "./writeDocx";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#202124" };
const PARA = { align: "left" as const, lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

const STYLES: UnderlineStyle[] = ["single", "double", "thick", "dotted", "dash", "dotDash", "dotDotDash", "wave"];

const styledRun = (text: string, patch: Partial<CharStyle>): Run => ({ text, style: { ...CHAR, ...patch } });

describe("underline style + color .docx round-trip", () => {
  it("preserves every underline line style through export → re-import", () => {
    const runs: Run[] = STYLES.map((s, i) =>
      styledRun(`u${i}`, s === "single" ? { underline: true } : { underline: true, underlineStyle: s }),
    );
    const doc: Document = {
      section: SECTION,
      blocks: [{ kind: "paragraph", id: "p1", revision: 0, style: { ...PARA }, runs }],
    };

    const back = runImport(writeDocx(doc).bytes).doc;
    const p = back.blocks[0] as Paragraph;
    STYLES.forEach((s, i) => {
      const r = p.runs.find((rr) => rr.text === `u${i}`);
      expect(r, `run u${i}`).toBeTruthy();
      expect(r!.style.underline).toBe(true);
      // "single" carries no explicit style (a plain solid line); the rest survive.
      expect(r!.style.underlineStyle).toBe(s === "single" ? undefined : s);
    });
  });

  it("preserves an explicit underline color (and pairs it with a style)", () => {
    const doc: Document = {
      section: SECTION,
      blocks: [{
        kind: "paragraph", id: "p1", revision: 0, style: { ...PARA },
        runs: [
          styledRun("red wavy", { underline: true, underlineStyle: "wave", underlineColor: "#d93025" }),
          styledRun("blue double", { underline: true, underlineStyle: "double", underlineColor: "#1a73e8" }),
          styledRun("auto single", { underline: true }),
        ],
      }],
    };

    const back = runImport(writeDocx(doc).bytes).doc;
    const p = back.blocks[0] as Paragraph;
    const red = p.runs.find((r) => r.text === "red wavy");
    expect(red!.style.underlineStyle).toBe("wave");
    expect(red!.style.underlineColor).toBe("#d93025");
    const blue = p.runs.find((r) => r.text === "blue double");
    expect(blue!.style.underlineStyle).toBe("double");
    expect(blue!.style.underlineColor).toBe("#1a73e8");
    // A plain underline with no color stays color-less (paints in the text color).
    const plain = p.runs.find((r) => r.text === "auto single");
    expect(plain!.style.underline).toBe(true);
    expect(plain!.style.underlineColor).toBeUndefined();
  });

  it("folds Word's heavy/long underline variants onto the model's base styles", () => {
    const doc: Document = {
      section: SECTION,
      blocks: [{
        kind: "paragraph", id: "p1", revision: 0, style: { ...PARA },
        // dottedHeavy/dashedHeavy aren't model styles; the importer folds them.
        runs: [styledRun("heavy", { underline: true, underlineStyle: "dottedHeavy" as UnderlineStyle })],
      }],
    };
    const back = runImport(writeDocx(doc).bytes).doc;
    const r = (back.blocks[0] as Paragraph).runs[0]!;
    expect(r.style.underline).toBe(true);
    expect(r.style.underlineStyle).toBe("dotted");
  });

  it("resolves a themed underline color (w:u/@w:themeColor) through theme1.xml", () => {
    // accent1 = 4472C4 in the bundled THEME_XML fixture.
    const body = `<w:p><w:r><w:rPr><w:u w:val="single" w:themeColor="accent1"/></w:rPr><w:t>themed</w:t></w:r></w:p>`;
    const back = runImport(styledDocx(body)).doc;
    const r = (back.blocks[0] as Paragraph).runs[0]!;
    expect(r.style.underline).toBe(true);
    expect(r.style.underlineColor).toBe("#4472c4");
  });
});
