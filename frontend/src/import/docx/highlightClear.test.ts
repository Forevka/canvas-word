// Ground-truth import + round-trip test for issue #155: a run's highlight is
// tri-state — absent (inherit the character style's highlight), a named color, and
// explicit `<w:highlight w:val="none"/>` (an override that CLEARS a highlight
// inherited from a character style). The importer used to drop "none" and the
// cascade merge couldn't override the inherited color, so a cleared run kept the
// style's highlight. Asserted against the SOURCE XML's meaning plus export re-emission.

import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import type { Document, Paragraph, Run } from "@cw/shared";
import { writeDocx } from "../../export/docx/writeDocx";
import { stylesPartXml, styledDocx } from "./fixture";
import { runImport } from "./pipeline";

// A CHARACTER style that carries a yellow highlight (like a "Marker" style).
const STYLES = stylesPartXml(
  `<w:style w:type="character" w:styleId="HL">
     <w:name w:val="Marker"/>
     <w:rPr><w:highlight w:val="yellow"/></w:rPr>
   </w:style>`,
);
const rStyle = (extra: string, text: string): string =>
  `<w:r><w:rPr><w:rStyle w:val="HL"/>${extra}</w:rPr><w:t>${text}</w:t></w:r>`;
const importRuns = (runsXml: string): Run[] => {
  const d: Document = runImport(styledDocx(`<w:p>${runsXml}</w:p>`, STYLES)).doc;
  const p = d.blocks[0] as Paragraph;
  expect(p.kind).toBe("paragraph");
  return p.runs;
};
const exportedXml = (runsXml: string): string => {
  const d: Document = runImport(styledDocx(`<w:p>${runsXml}</w:p>`, STYLES)).doc;
  return strFromU8(unzipSync(writeDocx(d).bytes)["word/document.xml"]!);
};

describe("run w:highlight clear tri-state (issue #155)", () => {
  it("inherits the character style's highlight when the run doesn't clear it (sanity)", () => {
    const runs = importRuns(rStyle("", "marked"));
    expect(runs[0]!.style.highlightColor).toBeDefined();
    expect(runs[0]!.style.highlightCleared).toBeUndefined();
  });

  it("clears the inherited highlight on an explicit w:val=none", () => {
    const runs = importRuns(rStyle(`<w:highlight w:val="none"/>`, "plain"));
    expect(runs[0]!.style.highlightColor).toBeUndefined();
    expect(runs[0]!.style.highlightCleared).toBe(true);
  });
});

describe("run w:highlight clear export/round-trip (issue #155)", () => {
  it("re-emits w:highlight w:val=none for a cleared run under a highlighted char style", () => {
    expect(exportedXml(rStyle(`<w:highlight w:val="none"/>`, "plain"))).toContain(`<w:highlight w:val="none"/>`);
  });

  it("survives import → export → import (the clear is not lost)", () => {
    const a: Document = runImport(styledDocx(`<w:p>${rStyle(`<w:highlight w:val="none"/>`, "plain")}</w:p>`, STYLES)).doc;
    const b = runImport(writeDocx(a).bytes).doc;
    const run = (b.blocks[0] as Paragraph).runs[0]!;
    expect(run.style.highlightColor).toBeUndefined(); // did NOT re-inherit the style highlight
    expect(run.style.highlightCleared).toBe(true);
  });
});
