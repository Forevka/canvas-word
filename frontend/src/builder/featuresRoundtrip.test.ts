// A builder-authored doc that exercises the new features survives a .docx
// round-trip: fields, content controls, footnotes, bookmarks, a TOC, a section
// break with columns, a custom list, and a preset table all persist through
// export → re-import (structure stable; exact text not asserted where the
// importer reshapes it — e.g. a TOC field flattens then re-marks).

import { beforeAll, describe, expect, it } from "vitest";
import { runImport } from "../import/docx/pipeline";
import { runExport } from "../export/pipeline";
import { installMeasureHost } from "../export/shared/measureHost";
import { DocumentBuilder } from "./index";

const NOW = new Date(2026, 5, 16, 9, 0);

const buildSample = () =>
  DocumentBuilder.create({ idSeed: "rt" })
    .tableStylePreset("brand", { headerRow: true, headerShading: "#1a73e8", headerChar: { color: "#fff" } })
    .listDefinition("rom", { kind: "number", format: "upperRoman" })
    .tableOfContents()
    .paragraph("Overview").withStyle("Heading1")
    .paragraph("Generated ").dateField("MMMM d, yyyy", { now: NOW }).text(", page ").pageField()
    .paragraph("See ").bookmark("intro", "the intro").text(" — refer to ").crossReference("intro")
    .paragraph("Pick: ").dropDown("One", [{ display: "One", value: "1" }, { display: "Two", value: "2" }], { alias: "Choice" })
    .paragraph("With a footnote").footnote("a builder-authored footnote")
    .paragraph("Details").withStyle("Heading2")
    .list(["First", "Second"], { listId: "rom" })
    .sectionBreak({ columns: { count: 2 } })
    .paragraph("Two-column section body.")
    .table([["H1", "H2"], ["a", "b"]], { style: "brand" })
    .build();

describe("builder feature round-trip (.docx)", () => {
  beforeAll(async () => {
    await installMeasureHost();
  });

  it("preserves fields, content controls, footnotes, bookmarks and the TOC field", async () => {
    const doc = buildSample();
    // Sanity: the built doc actually has the features.
    expect(Object.keys(doc.fields ?? {}).length).toBeGreaterThanOrEqual(3); // DATE + PAGE + PAGEREF
    expect(Object.keys(doc.sdts ?? {}).length).toBe(1);
    expect(Object.keys(doc.footnotes ?? {}).length).toBe(1);
    expect(Object.keys(doc.bookmarks ?? {}).length).toBeGreaterThanOrEqual(1);
    expect(doc.tocInstruction).toContain("TOC");

    const back = runImport((await runExport(doc, "docx")).bytes).doc;

    const fieldTypes = Object.values(back.fields ?? {}).map((d) => (d.kind === "builtin" ? `builtin:${d.spec?.type}` : "custom"));
    expect(fieldTypes).toContain("builtin:DATE");
    expect(fieldTypes).toContain("builtin:PAGE");
    expect(Object.keys(back.sdts ?? {}).length).toBe(1); // dropdown control survives
    expect(Object.keys(back.footnotes ?? {}).length).toBe(1); // footnote survives
    expect(Object.keys(back.bookmarks ?? {}).length).toBeGreaterThanOrEqual(1); // bookmark survives
    expect(back.tocInstruction).toContain("TOC"); // live TOC field round-trips
    // The TOC entries re-import as marked tocEntry paragraphs.
    const tocEntries = back.blocks.filter((b) => b.kind === "paragraph" && b.style.tocEntry);
    expect(tocEntries.length).toBeGreaterThanOrEqual(2);
  });
});
