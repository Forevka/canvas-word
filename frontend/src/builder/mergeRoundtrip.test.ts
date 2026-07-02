// mergeDocuments survives a .docx round-trip: two builder-authored parts — given
// the SAME id seed on purpose, so their block ids, bookmark names and note refs
// all collide — merge into one document that exports and re-imports with both
// parts' content, sections, footnotes and bookmarks intact.

import { beforeAll, describe, expect, it } from "vitest";
import { mergeAll, mergeDocuments, resolveSections } from "@cw/shared";
import { runImport } from "../import/docx/pipeline";
import { runExport } from "../export/pipeline";
import { installMeasureHost } from "../export/shared/measureHost";
import { DocumentBuilder } from "./index";

// Same idSeed for every part → identical id spaces → maximal collisions for merge.
// `pageSize` lets a part carry distinct page GEOMETRY so the section seam between
// parts is MATERIAL: the importer intentionally flows a same-page-size ("footer
// only") break, and materializes one that changes page size/columns/numbering, so
// a distinct page size keeps the section boundary observable through re-import.
const part = (heading: string, body: string, pageSize: "Letter" | "A4" | "Legal" = "Letter") =>
  DocumentBuilder.create({ idSeed: "rt", pageSize })
    .paragraph(heading).withStyle("Heading1")
    .paragraph(body).footnote(`${heading} footnote`)
    .paragraph("See ").bookmark("intro", "the intro")
    .build();

const textOfDoc = (blocks: { kind: string }[]): string =>
  blocks
    .filter((b): b is { kind: "paragraph"; runs: { text: string }[] } => b.kind === "paragraph")
    .map((p) => p.runs.map((r) => r.text).join(""))
    .join("\n");

describe("mergeDocuments round-trip (.docx)", () => {
  beforeAll(async () => {
    await installMeasureHost();
  });

  it("merges two colliding parts and survives export → re-import", async () => {
    const a = part("Part A", "Alpha body.", "Letter");
    const b = part("Part B", "Beta body.", "A4");

    const { doc: merged, warnings } = mergeDocuments(a, b, { sectionBreak: "nextPage" });

    // Pre-export invariants: both parts present, ids disjoint, collisions handled.
    const ids = merged.blocks.map((x) => x.id);
    expect(new Set(ids).size).toBe(ids.length);
    expect(Object.keys(merged.footnotes ?? {})).toHaveLength(2);
    expect(Object.keys(merged.bookmarks ?? {})).toHaveLength(2); // intro + intro__2
    expect(warnings.some((w) => w.code === "bookmark-renamed")).toBe(true);
    // The seam split the document into two sections.
    expect(resolveSections(merged)).toHaveLength(2);

    // Round-trip through the real exporter + importer.
    const back = runImport((await runExport(merged, "docx")).bytes).doc;

    const text = textOfDoc(back.blocks);
    expect(text).toContain("Part A");
    expect(text).toContain("Part B");
    expect(text).toContain("Alpha body.");
    expect(text).toContain("Beta body.");
    // Both footnotes and both bookmarks survive the round-trip.
    expect(Object.keys(back.footnotes ?? {}).length).toBeGreaterThanOrEqual(2);
    expect(Object.keys(back.bookmarks ?? {}).length).toBeGreaterThanOrEqual(2);
    // The section break persists (≥2 sections after re-import).
    expect(resolveSections(back).length).toBeGreaterThanOrEqual(2);
  });

  it("mergeAll folds three parts into one document that round-trips", async () => {
    const docs = [part("Part A", "A."), part("Part B", "B."), part("Part C", "C.")];
    const { doc: merged } = mergeAll(docs, { sectionBreak: "nextPage" });

    expect(resolveSections(merged)).toHaveLength(3);

    const back = runImport((await runExport(merged, "docx")).bytes).doc;
    const text = textOfDoc(back.blocks);
    for (const h of ["Part A", "Part B", "Part C"]) expect(text).toContain(h);
    expect(Object.keys(back.footnotes ?? {}).length).toBeGreaterThanOrEqual(3);
  });
});
