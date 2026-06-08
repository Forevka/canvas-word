// keepLines, newspaper columns, pageNumberStart, and footnotes.

import { describe, expect, it } from "vitest";
import type { Paragraph } from "../../model/document";
import { runImport } from "./pipeline";
import { CONTENT_TYPES_XML, documentXml, makeDocx, REL_TYPES, relsXml, simpleDocx } from "./fixture";

const para = (b: unknown): Paragraph => {
  expect((b as Paragraph).kind).toBe("paragraph");
  return b as Paragraph;
};
const importBody = (body: string) => runImport(simpleDocx(body));

describe("paragraph keepLines", () => {
  it("maps w:keepLines to keepLinesTogether", () => {
    const r = importBody(`<w:p><w:pPr><w:keepLines/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(para(r.doc.blocks[0]).style.keepLinesTogether).toBe(true);
  });
  it("leaves it unset when absent", () => {
    const r = importBody(`<w:p><w:r><w:t>x</w:t></w:r></w:p>`);
    expect(para(r.doc.blocks[0]).style.keepLinesTogether).toBeUndefined();
  });
});

describe("section columns and page numbering", () => {
  const withSect = (sect: string) =>
    importBody(`<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr>${sect}<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`);

  it("maps multi-column w:cols to SectionProps.columns with a px gap", () => {
    const r = withSect(`<w:cols w:num="2" w:space="720"/>`);
    expect(r.doc.section.columns).toEqual({ count: 2, gapPx: 48 });
  });
  it("ignores single-column w:cols", () => {
    const r = withSect(`<w:cols w:num="1"/>`);
    expect(r.doc.section.columns).toBeUndefined();
  });
  it("defaults the column gap to 0.5in when w:space is absent", () => {
    const r = withSect(`<w:cols w:num="3"/>`);
    expect(r.doc.section.columns).toEqual({ count: 3, gapPx: 48 });
  });
  it("maps w:pgNumType/@w:start to pageNumberStart", () => {
    const r = withSect(`<w:pgNumType w:start="5"/>`);
    expect(r.doc.section.pageNumberStart).toBe(5);
  });
});

describe("footnotes", () => {
  function docxWithFootnotes(body: string, notes: string): Uint8Array {
    return makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(body),
      "word/_rels/document.xml.rels": relsXml([
        { id: "rF", type: REL_TYPES.footnotes, target: "footnotes.xml" },
      ]),
      "word/footnotes.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  ${notes}
</w:footnotes>`,
    });
  }

  const ref = (id: string) =>
    `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
  const note = (id: string, text: string) =>
    `<w:footnote w:id="${id}"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:footnote>`;

  it("maps a footnote reference and its body", () => {
    const r = runImport(
      docxWithFootnotes(`<w:p><w:r><w:t>see</w:t></w:r>${ref("2")}</w:p>`, note("2", "the note text")),
    );
    const markerRun = para(r.doc.blocks[0]).runs.find((x) => x.style.footnoteRef);
    expect(markerRun).toBeDefined();
    expect(markerRun!.text).toBe("1"); // first reference → number 1
    expect(markerRun!.style.verticalAlign).toBe("super");
    expect(markerRun!.style.footnoteRef).toBe("fn2");
    // Body stored under the same key; the w:footnoteRef placeholder is dropped.
    expect(r.doc.footnotes?.fn2).toBeDefined();
    expect(r.doc.footnotes!.fn2!.map((p) => p.runs.map((x) => x.text).join("")).join("")).toContain(
      "the note text",
    );
  });

  it("numbers multiple footnotes sequentially in document order", () => {
    const r = runImport(
      docxWithFootnotes(
        `<w:p><w:r><w:t>a</w:t></w:r>${ref("2")}<w:r><w:t>b</w:t></w:r>${ref("3")}</w:p>`,
        note("2", "first") + note("3", "second"),
      ),
    );
    const markers = para(r.doc.blocks[0]).runs.filter((x) => x.style.footnoteRef);
    expect(markers.map((m) => m.text)).toEqual(["1", "2"]);
    expect(markers.map((m) => m.style.footnoteRef)).toEqual(["fn2", "fn3"]);
    expect(Object.keys(r.doc.footnotes!)).toEqual(["fn2", "fn3"]);
  });

  it("ignores separator/continuation pseudo-notes", () => {
    const r = runImport(docxWithFootnotes(`<w:p><w:r><w:t>x</w:t></w:r>${ref("2")}</w:p>`, note("2", "real")));
    // Only the real note is in the registry (separator id -1 excluded).
    expect(Object.keys(r.doc.footnotes!)).toEqual(["fn2"]);
  });
});
