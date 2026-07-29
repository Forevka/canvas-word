// A builder-authored doc that exercises the new features survives a .docx
// round-trip: fields, content controls, footnotes, bookmarks, a TOC, a section
// break with columns, a custom list, and a preset table all persist through
// export → re-import (structure stable; exact text not asserted where the
// importer reshapes it — e.g. a TOC field flattens then re-marks).

import { beforeAll, describe, expect, it } from "vitest";
import { AUTO_PARA_SPACING_PX, type Block } from "@cw/shared";
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
    .paragraph("See ").bookmark("intro", "the intro").text(" — refer to ").crossReference("intro").text(" on page ").crossReference("intro", { kind: "pageRef" })
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
    expect(Object.keys(doc.fields ?? {}).length).toBeGreaterThanOrEqual(4); // DATE + PAGE + REF + PAGEREF
    expect(Object.keys(doc.sdts ?? {}).length).toBe(1);
    expect(Object.keys(doc.footnotes ?? {}).length).toBe(1);
    expect(Object.keys(doc.bookmarks ?? {}).length).toBeGreaterThanOrEqual(1);
    expect(doc.tocInstruction).toContain("TOC");

    const back = runImport((await runExport(doc, "docx")).bytes).doc;

    const fieldTypes = Object.values(back.fields ?? {}).map((d) => (d.kind === "builtin" ? `builtin:${d.spec?.type}` : "custom"));
    expect(fieldTypes).toContain("builtin:DATE");
    expect(fieldTypes).toContain("builtin:PAGE");
    expect(fieldTypes).toContain("builtin:REF"); // cross-reference (text) survives
    expect(fieldTypes).toContain("builtin:PAGEREF"); // cross-reference (page) survives
    // The REF/PAGEREF fields target exactly `intro` — asserted on the parsed spec,
    // not a substring of the instruction, so a stray `intro2` can't satisfy it. The
    // length check keeps `every`/Set from passing vacuously on an empty list.
    const xrefTargets = Object.values(back.fields ?? {})
      .map((d) => d.spec)
      .filter((s) => s?.type === "REF" || s?.type === "PAGEREF")
      .map((s) => (s && (s.type === "REF" || s.type === "PAGEREF") ? s.bookmark : undefined));
    expect(xrefTargets.length).toBeGreaterThanOrEqual(2); // the REF and the PAGEREF
    expect(new Set(xrefTargets)).toEqual(new Set(["intro"]));
    expect(Object.keys(back.sdts ?? {}).length).toBe(1); // dropdown control survives
    expect(Object.keys(back.footnotes ?? {}).length).toBe(1); // footnote survives
    expect(Object.keys(back.bookmarks ?? {}).length).toBeGreaterThanOrEqual(1); // bookmark survives
    expect(back.tocInstruction).toContain("TOC"); // live TOC field round-trips
    // The TOC entries re-import as marked tocEntry paragraphs.
    const tocEntries = back.blocks.filter((b) => b.kind === "paragraph" && b.style.tocEntry);
    expect(tocEntries.length).toBeGreaterThanOrEqual(2);
  });

  it("preserves an explicit shading clear against a shaded named style (issue #147)", async () => {
    const doc = DocumentBuilder.create({ idSeed: "shd" })
      .style({ id: "Callout", name: "Callout", type: "paragraph", char: {}, para: { shading: "#d9ead3" } })
      .paragraph("inherits the callout fill").withStyle("Callout")
      .paragraph("clears the callout fill").withStyle("Callout").clearShading()
      .build();

    // The builder bakes the style fill onto the first paragraph and marks the second cleared.
    const built = doc.blocks.filter((b): b is Extract<Block, { kind: "paragraph" }> => b.kind === "paragraph");
    expect(built[0]!.style.shading).toBe("#d9ead3");
    expect(built[1]!.style.shading).toBeUndefined();
    expect(built[1]!.style.shadingCleared).toBe(true);

    const back = runImport((await runExport(doc, "docx")).bytes).doc;
    const paras = back.blocks.filter((b): b is Extract<Block, { kind: "paragraph" }> => b.kind === "paragraph");
    expect(paras[0]!.style.shading).toBe("#d9ead3"); // still inherits the fill
    expect(paras[1]!.style.shading).toBeUndefined(); // clear survived — the fill did NOT return
    expect(paras[1]!.style.shadingCleared).toBe(true);
  });

  it("keeps shading and the clear marker mutually exclusive across builder order (issue #147)", () => {
    const doc = DocumentBuilder.create({ idSeed: "shd2" })
      .style({ id: "Callout", name: "Callout", type: "paragraph", char: {}, para: { shading: "#d9ead3" } })
      // clearShading THEN withStyle: the later style's fill wins, the stale marker is dropped.
      .paragraph("a").clearShading().withStyle("Callout")
      // withStyle THEN clearShading: the later clear wins, the style fill is dropped.
      .paragraph("b").withStyle("Callout").clearShading()
      .build();
    const p = doc.blocks.filter((b): b is Extract<Block, { kind: "paragraph" }> => b.kind === "paragraph");
    expect(p[0]!.style.shading).toBe("#d9ead3");
    expect(p[0]!.style.shadingCleared).toBeUndefined();
    expect(p[1]!.style.shading).toBeUndefined();
    expect(p[1]!.style.shadingCleared).toBe(true);
  });

  it("preserves a cell shading clear against a table-level fill (issue #150)", async () => {
    const doc = DocumentBuilder.create({ idSeed: "cellshd" })
      // Table default shading #eeeeee; the first cell explicitly clears it.
      .table([[{ text: "cleared", shadingCleared: true }, "inherits"]], { shading: "#eeeeee" })
      .build();
    const t0 = doc.blocks.find((b): b is Extract<Block, { kind: "table" }> => b.kind === "table")!;
    // The cleared cell does NOT pick up the table default; the sibling does.
    expect(t0.rows[0]!.cells[0]!.shading).toBeUndefined();
    expect(t0.rows[0]!.cells[0]!.shadingCleared).toBe(true);
    expect(t0.rows[0]!.cells[1]!.shading).toBe("#eeeeee");

    const back = runImport((await runExport(doc, "docx")).bytes).doc;
    const t = back.blocks.find((b): b is Extract<Block, { kind: "table" }> => b.kind === "table")!;
    expect(t.rows[0]!.cells[0]!.shading).toBeUndefined(); // clear survived
    expect(t.rows[0]!.cells[0]!.shadingCleared).toBe(true);
    expect(t.rows[0]!.cells[1]!.shading).toBe("#eeeeee"); // sibling still shaded
  });

  it("preserves a paragraph border-box clear against a boxed named style (issue #153)", async () => {
    const box = { top: { color: "#1a73e8", widthPx: 2 }, bottom: { color: "#1a73e8", widthPx: 2 } };
    const doc = DocumentBuilder.create({ idSeed: "pbdr" })
      .style({ id: "Boxed", name: "Boxed", type: "paragraph", char: {}, para: { borders: box } })
      .paragraph("keeps the box").withStyle("Boxed")
      .paragraph("clears the box").withStyle("Boxed").clearBorders()
      .build();
    const p0 = doc.blocks.filter((b): b is Extract<Block, { kind: "paragraph" }> => b.kind === "paragraph");
    expect(p0[0]!.style.borders).toBeDefined();
    expect(p0[1]!.style.borders).toBeUndefined();
    expect(p0[1]!.style.bordersCleared).toBe(true);

    const back = runImport((await runExport(doc, "docx")).bytes).doc;
    const p = back.blocks.filter((b): b is Extract<Block, { kind: "paragraph" }> => b.kind === "paragraph");
    expect(p[0]!.style.borders).toBeDefined(); // still inherits the box
    expect(p[1]!.style.borders).toBeUndefined(); // clear survived
    expect(p[1]!.style.bordersCleared).toBe(true);
  });

  it("preserves a cleared tab stop against a tabbed named style (issue #154)", async () => {
    const doc = DocumentBuilder.create({ idSeed: "tab" })
      .style({ id: "Tabbed", name: "Tabbed", type: "paragraph", char: {}, para: { tabStops: [{ posPx: 96 }] } })
      .paragraph("keeps the tab").withStyle("Tabbed")
      .paragraph("clears the tab").withStyle("Tabbed").tabStops([{ posPx: 96, cleared: true }])
      .build();
    const p0 = doc.blocks.filter((b): b is Extract<Block, { kind: "paragraph" }> => b.kind === "paragraph");
    expect(p0[0]!.style.tabStops).toEqual([{ posPx: 96 }]);
    expect(p0[1]!.style.tabStops).toEqual([{ posPx: 96, cleared: true }]);

    const back = runImport((await runExport(doc, "docx")).bytes).doc;
    const p = back.blocks.filter((b): b is Extract<Block, { kind: "paragraph" }> => b.kind === "paragraph");
    expect((p[0]!.style.tabStops ?? []).some((t) => !t.cleared && Math.round(t.posPx) === 96)).toBe(true);
    const cleared = p[1]!.style.tabStops ?? [];
    expect(cleared.some((t) => t.cleared && Math.round(t.posPx) === 96)).toBe(true); // clear survived
    expect(cleared.some((t) => !t.cleared && Math.round(t.posPx) === 96)).toBe(false);
  });

  it("preserves a run highlight clear against a highlighted character style (issue #155)", async () => {
    const doc = DocumentBuilder.create({ idSeed: "hl" })
      .style({ id: "HL", name: "Marker", type: "character", char: { highlightColor: "#ffff00" }, para: {} })
      .paragraph("").text("inherits", { charStyleId: "HL" })
      .paragraph("").text("cleared", { charStyleId: "HL" }).clearHighlight()
      .build();
    const built = doc.blocks.filter((b): b is Extract<Block, { kind: "paragraph" }> => b.kind === "paragraph");
    expect(built[1]!.runs[0]!.style.highlightCleared).toBe(true);

    const back = runImport((await runExport(doc, "docx")).bytes).doc;
    const p = back.blocks.filter((b): b is Extract<Block, { kind: "paragraph" }> => b.kind === "paragraph");
    expect(p[0]!.runs[0]!.style.highlightColor).toBeDefined(); // inherits the style highlight
    expect(p[1]!.runs[0]!.style.highlightColor).toBeUndefined(); // clear survived
    expect(p[1]!.runs[0]!.style.highlightCleared).toBe(true);
  });

  it("preserves paragraph auto-spacing (issue #160)", async () => {
    const doc = DocumentBuilder.create({ idSeed: "auto" })
      .paragraph("autospaced").spacing({ beforeAuto: true, afterAuto: true })
      .build();
    const built = doc.blocks.find((b): b is Extract<Block, { kind: "paragraph" }> => b.kind === "paragraph")!;
    expect(built.style.spaceBeforeAuto).toBe(true);
    expect(built.style.spaceBeforePx).toBe(AUTO_PARA_SPACING_PX);

    const back = runImport((await runExport(doc, "docx")).bytes).doc;
    const p = back.blocks.find((b): b is Extract<Block, { kind: "paragraph" }> => b.kind === "paragraph")!;
    expect(p.style.spaceBeforeAuto).toBe(true);
    expect(p.style.spaceAfterAuto).toBe(true);
    expect(p.style.spaceBeforePx).toBe(AUTO_PARA_SPACING_PX);
  });
});
