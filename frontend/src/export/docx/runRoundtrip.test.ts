// Run-level round-trip fidelity (issue #57): character tracking (w:spacing) and the
// complex-script (w:cs) + East-Asian (w:eastAsia) font slots must survive a full
// export → re-import cycle. Tracking import was previously export-only, and the
// CS/EA rFonts slots were dropped on import (only w:ascii was read).
import { describe, expect, it } from "vitest";
import type { CharStyle, Document, Paragraph, SectionProps, Stylesheet } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { writeDocx } from "./writeDocx";

const CHAR: CharStyle = { fontFamily: "Georgia, serif", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#202124" };
const PARA = { align: "left" as const, lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };
const sheet: Stylesheet = { defaultStyleId: "Normal", styles: [{ id: "Normal", name: "Normal", type: "paragraph", char: { ...CHAR }, para: { ...PARA } }] };

const tracked: CharStyle = { ...CHAR, letterSpacingPx: 2 };
const cjkRun: CharStyle = { ...CHAR, fontFamilyComplexScript: "Amiri, serif", fontFamilyEastAsia: "SimSun, serif" };

const doc: Document = {
  section: SECTION,
  stylesheet: sheet,
  blocks: [{
    kind: "paragraph", id: "p1", revision: 0, style: { ...PARA, namedStyle: "Normal" },
    runs: [
      { text: "plain ", style: { ...CHAR } },
      { text: "tracked", style: tracked },
      { text: " 日本語", style: cjkRun },
    ],
  }],
};

const reimport = (): Paragraph => (runImport(writeDocx(doc).bytes).doc.blocks[0] as Paragraph);

describe("run-level round-trip (issue #57)", () => {
  it("re-imports run character tracking (w:spacing)", () => {
    const runs = reimport().runs;
    expect(runs.find((r) => r.text === "tracked")?.style.letterSpacingPx).toBe(2);
    // A plain run gains no spurious tracking.
    expect(runs.find((r) => r.text.includes("plain"))?.style.letterSpacingPx).toBeUndefined();
  });

  it("re-imports the complex-script and East-Asian font slots", () => {
    const cjk = reimport().runs.find((r) => r.text.includes("日本語"))!;
    expect(cjk.style.fontFamilyComplexScript).toBe("Amiri, serif");
    expect(cjk.style.fontFamilyEastAsia).toBe("SimSun, serif");
    // The Latin face is untouched by the CS/EA slots.
    expect(cjk.style.fontFamily).toBe("Georgia, serif");
  });

  it("does not invent CS/EA slots for a plain run (w:cs == w:ascii)", () => {
    const plain = reimport().runs.find((r) => r.text.includes("plain"))!;
    expect(plain.style.fontFamilyComplexScript).toBeUndefined();
    expect(plain.style.fontFamilyEastAsia).toBeUndefined();
  });
});
