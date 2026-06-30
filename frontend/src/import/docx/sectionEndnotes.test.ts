// Endnote import (endnotes.xml + w:endnoteReference). Mirror of the footnote
// import tests, but the bodies land in Document.endnotes (document-end placement).

import { describe, expect, it } from "vitest";
import type { Paragraph } from "@cw/shared";
import { runImport } from "./pipeline";
import { CONTENT_TYPES_XML, documentXml, makeDocx, REL_TYPES, relsXml } from "./fixture";

const para = (b: unknown): Paragraph => {
  expect((b as Paragraph).kind).toBe("paragraph");
  return b as Paragraph;
};

describe("endnotes", () => {
  function docxWithEndnotes(body: string, notes: string): Uint8Array {
    return makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(body),
      "word/_rels/document.xml.rels": relsXml([
        { id: "rE", type: REL_TYPES.endnotes, target: "endnotes.xml" },
      ]),
      "word/endnotes.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
  ${notes}
</w:endnotes>`,
    });
  }

  const ref = (id: string) =>
    `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`;
  const note = (id: string, text: string) =>
    `<w:endnote w:id="${id}"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:endnote>`;

  it("maps an endnote reference and its body", () => {
    const r = runImport(
      docxWithEndnotes(`<w:p><w:r><w:t>see</w:t></w:r>${ref("2")}</w:p>`, note("2", "the note text")),
    );
    const markerRun = para(r.doc.blocks[0]).runs.find((x) => x.style.endnoteRef);
    expect(markerRun).toBeDefined();
    expect(markerRun!.text).toBe("1"); // first reference → number 1
    expect(markerRun!.style.verticalAlign).toBe("super");
    expect(markerRun!.style.endnoteRef).toBe("en2");
    // Body stored under the same key; the w:endnoteRef placeholder is dropped.
    expect(r.doc.endnotes?.en2).toBeDefined();
    expect(r.doc.endnotes!.en2!.map((p) => p.runs.map((x) => x.text).join("")).join("")).toContain(
      "the note text",
    );
  });

  it("numbers multiple endnotes sequentially in document order", () => {
    const r = runImport(
      docxWithEndnotes(
        `<w:p><w:r><w:t>a</w:t></w:r>${ref("2")}<w:r><w:t>b</w:t></w:r>${ref("3")}</w:p>`,
        note("2", "first") + note("3", "second"),
      ),
    );
    const markers = para(r.doc.blocks[0]).runs.filter((x) => x.style.endnoteRef);
    expect(markers.map((m) => m.text)).toEqual(["1", "2"]);
    expect(markers.map((m) => m.style.endnoteRef)).toEqual(["en2", "en3"]);
    expect(Object.keys(r.doc.endnotes!)).toEqual(["en2", "en3"]);
  });

  it("ignores separator/continuation pseudo-notes", () => {
    const r = runImport(docxWithEndnotes(`<w:p><w:r><w:t>x</w:t></w:r>${ref("2")}</w:p>`, note("2", "real")));
    // Only the real note is in the registry (separator id -1 excluded).
    expect(Object.keys(r.doc.endnotes!)).toEqual(["en2"]);
  });
});
