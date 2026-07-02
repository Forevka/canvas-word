import { describe, expect, it } from "vitest";
import type { Block, Document, Paragraph, ParaStyle, Run, SectionProps, TableBlock } from "./document";
import { DEFAULT_CHAR_STYLE, DEFAULT_PARA_STYLE } from "./defaults";
import type { Stylesheet } from "./stylesheet";
import { mergeAll, mergeDocuments } from "./merge";

// ---- tiny concrete-model builders -----------------------------------------

function sec(overrides: Partial<SectionProps> = {}): SectionProps {
  return { pageWidthPx: 794, pageHeightPx: 1123, marginPx: { top: 96, right: 96, bottom: 96, left: 96 }, ...overrides };
}
function run(text: string, style: Partial<Run["style"]> = {}): Run {
  return { text, style: { ...DEFAULT_CHAR_STYLE, ...style } };
}
function para(id: string, text: string, style: Partial<ParaStyle> = {}, runs?: Run[]): Paragraph {
  return { kind: "paragraph", id, revision: 0, runs: runs ?? [run(text)], style: { ...DEFAULT_PARA_STYLE, ...style } };
}
function doc(blocks: Block[], extra: Partial<Document> = {}): Document {
  return { section: sec(), blocks, ...extra };
}

/** Every block id in a document (recursing tables + note bodies + section bands). */
function allBlockIds(d: Document): string[] {
  const ids: string[] = [];
  const walk = (blocks: readonly Block[]): void => {
    for (const b of blocks) {
      ids.push(b.id);
      if (b.kind === "table") for (const row of b.rows) for (const cell of row.cells) walk(cell.blocks);
    }
  };
  walk(d.blocks);
  for (const paras of Object.values(d.footnotes ?? {})) walk(paras);
  for (const paras of Object.values(d.endnotes ?? {})) walk(paras);
  return ids;
}

// ---------------------------------------------------------------------------

describe("mergeDocuments", () => {
  it("appends blocks with a section-break seam by default (nextPage)", () => {
    const dest = doc([para("d1", "Dest")]);
    const source: Document = { section: sec({ pageWidthPx: 500 }), blocks: [para("s1", "Src")] };

    const { doc: m } = mergeDocuments(dest, source);

    expect(m.blocks).toHaveLength(3); // dest, seam carrier, source
    expect(m.blocks[0]!.id).toBe("d1"); // destination ids never change
    const seam = m.blocks[1] as Paragraph;
    expect(seam.kind).toBe("paragraph");
    expect(seam.style.sectionBreak?.type).toBe("nextPage");
    // The seam fully describes the DESTINATION's section (so dest content keeps it).
    expect(seam.style.sectionBreak?.props.pageWidthPx).toBe(794);
    // The source's section governs the trailing (source) section.
    expect(m.section.pageWidthPx).toBe(500);
    // The source paragraph was cloned with a fresh id.
    expect(m.blocks[2]!.id).not.toBe("s1");
    expect((m.blocks[2] as Paragraph).runs[0]!.text).toBe("Src");
  });

  it("continuous/none appends inline with no seam and keeps the destination section", () => {
    const dest = doc([para("d1", "Dest")]);
    const source: Document = { section: sec({ pageWidthPx: 500 }), blocks: [para("s1", "Src")] };

    const { doc: m, warnings } = mergeDocuments(dest, source, { sectionBreak: "none" });

    expect(m.blocks).toHaveLength(2); // no seam carrier
    expect(m.section.pageWidthPx).toBe(794); // destination governs the whole tail
    expect(warnings.some((w) => w.code === "source-section-dropped")).toBe(true);
  });

  it("adopts the source section when the destination is empty", () => {
    const dest = doc([]);
    const source: Document = { section: sec({ pageWidthPx: 321 }), blocks: [para("s1", "Src")] };

    const { doc: m } = mergeDocuments(dest, source);

    expect(m.blocks).toHaveLength(1);
    expect(m.section.pageWidthPx).toBe(321);
  });

  it("remaps every source id so nothing collides with the destination", () => {
    const dest = doc([para("p1", "D"), para("p2", "D2")]);
    const source = doc([para("p1", "S"), para("p2", "S2")]); // same ids on purpose

    const { doc: m, idMap } = mergeDocuments(dest, source, { sectionBreak: "none" });

    const ids = allBlockIds(m);
    expect(new Set(ids).size).toBe(ids.length); // all unique
    expect(idMap.blocks["p1"]).toBeDefined();
    expect(idMap.blocks["p1"]).not.toBe("p1");
  });

  it("remaps ids inside nested table cells", () => {
    const cellPara = para("cp", "cell");
    const table: TableBlock = { kind: "table", id: "t1", revision: 0, rows: [{ cells: [{ id: "c1", blocks: [cellPara] }] }] };
    const dest = doc([para("cp", "dest reuse")]); // collides with the source cell paragraph id
    const source = doc([table]);

    const { doc: m } = mergeDocuments(dest, source, { sectionBreak: "none" });

    const mergedTable = m.blocks.find((b) => b.kind === "table") as TableBlock;
    expect(mergedTable.id).not.toBe("t1");
    const mergedCell = mergedTable.rows[0]!.cells[0]!;
    expect(mergedCell.id).not.toBe("c1"); // cell ids are freshly minted
    expect(mergedCell.blocks[0]!.id).not.toBe("cp"); // and its block was remapped
  });

  it("useDestination repoints a same-named style; source-only styles are added", () => {
    const destSheet: Stylesheet = {
      defaultStyleId: "Normal",
      styles: [
        { id: "Normal", name: "Normal", char: {}, para: {} },
        { id: "H", name: "Heading 1", char: { bold: true }, para: {} },
      ],
    };
    const srcSheet: Stylesheet = {
      defaultStyleId: "Normal",
      styles: [
        { id: "Normal", name: "Normal", char: {}, para: {} },
        { id: "H", name: "Heading 1", char: { bold: false }, para: {} }, // same name, different look
        { id: "Cap", name: "Caption", char: { italic: true }, para: {} }, // source-only
      ],
    };
    const dest = doc([para("d1", "D")], { stylesheet: destSheet });
    const source = doc([para("s1", "S", { namedStyle: "H" }), para("s2", "C", { namedStyle: "Cap" })], { stylesheet: srcSheet });

    const { doc: m, idMap } = mergeDocuments(dest, source, { styles: "useDestination", sectionBreak: "none" });

    // "Heading 1" collapsed onto the destination's; "Caption" added.
    const names = m.stylesheet!.styles.map((s) => s.name).sort();
    expect(names).toEqual(["Caption", "Heading 1", "Normal"]);
    expect(idMap.styles["H"]).toBe("H"); // repointed at the destination style
    // The source paragraph that used "H" now references the destination's "H".
    const sPara = m.blocks.find((b) => b.kind === "paragraph" && b.runs[0]?.text === "S") as Paragraph;
    expect(sPara.style.namedStyle).toBe("H");
    // Every referenced style resolves in the merged sheet.
    for (const b of m.blocks) {
      if (b.kind === "paragraph" && b.style.namedStyle) {
        expect(m.stylesheet!.styles.some((s) => s.id === b.style.namedStyle)).toBe(true);
      }
    }
  });

  it("keepSource renames a colliding style id and repoints the source's references", () => {
    const destSheet: Stylesheet = { defaultStyleId: "Normal", styles: [{ id: "H", name: "Heading 1", char: { bold: true }, para: {} }] };
    const srcSheet: Stylesheet = { defaultStyleId: "Normal", styles: [{ id: "H", name: "Heading 1", char: { bold: false }, para: {} }] };
    const dest = doc([para("d1", "D")], { stylesheet: destSheet });
    const source = doc([para("s1", "S", { namedStyle: "H" })], { stylesheet: srcSheet });

    const { doc: m, idMap } = mergeDocuments(dest, source, { styles: "keepSource", sectionBreak: "none" });

    expect(m.stylesheet!.styles).toHaveLength(2); // both kept
    expect(idMap.styles["H"]).not.toBe("H");
    const sPara = m.blocks.find((b) => b.kind === "paragraph" && b.runs[0]?.text === "S") as Paragraph;
    expect(sPara.style.namedStyle).toBe(idMap.styles["H"]);
  });

  it("always remaps colliding list ids and rewrites list references", () => {
    const dest = doc([para("d1", "D", { list: { listId: "1", level: 0 } })], {
      lists: { "1": { id: "1", levels: [{ format: "decimal", text: "%1.", indentLeftPx: 24, hangingPx: 24, start: 1 }] } },
    });
    const source = doc([para("s1", "S", { list: { listId: "1", level: 0 } })], {
      lists: { "1": { id: "1", levels: [{ format: "bullet", text: "", bulletChar: "•", indentLeftPx: 24, hangingPx: 24, start: 1 }] } },
    });

    const { doc: m, idMap } = mergeDocuments(dest, source, { sectionBreak: "none" });

    expect(Object.keys(m.lists!)).toHaveLength(2);
    const newListId = idMap.lists["1"];
    expect(newListId).toBeDefined();
    expect(newListId).not.toBe("1");
    expect(m.lists![newListId!]).toBeDefined();
    const sPara = m.blocks.find((b) => b.kind === "paragraph" && b.runs[0]?.text === "S") as Paragraph;
    expect(sPara.style.list?.listId).toBe(newListId);
  });

  it("remaps sdt ids on both block and inline paths, preserving nesting", () => {
    const dest = doc([para("d1", "D")], { sdts: { a: { type: "richText" } } });
    const srcPara = para("s1", "S", {}, [run("S", { sdtPath: ["a", "b"] })]);
    srcPara.sdtPath = ["a"]; // block-level control (Block.sdtPath, not ParaStyle)
    const source = doc([srcPara], { sdts: { a: { type: "plainText" }, b: { type: "richText" } } });

    const { doc: m, idMap } = mergeDocuments(dest, source, { sectionBreak: "none" });

    expect(Object.keys(m.sdts!).sort()).toHaveLength(3); // dest a + source a→renamed + b
    const sPara = m.blocks.find((b) => b.kind === "paragraph" && b.runs[0]?.text === "S") as Paragraph;
    const newA = idMap.sdts["a"];
    const newB = idMap.sdts["b"];
    expect(sPara.sdtPath).toEqual([newA]);
    expect(sPara.runs[0]!.style.sdtPath).toEqual([newA, newB]);
    for (const id of [newA!, newB!]) expect(m.sdts![id]).toBeDefined();
  });

  it("remaps footnote refs and carries the note bodies", () => {
    const dest = doc([para("d1", "D")], { footnotes: { "1": [para("dn", "dest note")] } });
    const source = doc([para("s1", "S", {}, [run("S", { footnoteRef: "1" })])], { footnotes: { "1": [para("sn", "src note")] } });

    const { doc: m, idMap } = mergeDocuments(dest, source, { sectionBreak: "none" });

    expect(Object.keys(m.footnotes!)).toHaveLength(2);
    const newRef = idMap.footnotes["1"];
    expect(newRef).not.toBe("1");
    const sPara = m.blocks.find((b) => b.kind === "paragraph" && b.runs[0]?.text === "S") as Paragraph;
    expect(sPara.runs[0]!.style.footnoteRef).toBe(newRef);
    expect(m.footnotes![newRef!]).toBeDefined();
  });

  it("rebases bookmark positions and renames colliding names (or drops when told)", () => {
    const dest = doc([para("d1", "D")], { bookmarks: { bm: { start: { blockId: "d1", offset: 0 }, end: { blockId: "d1", offset: 1 } } } });
    const source = doc([para("s1", "S")], { bookmarks: { bm: { start: { blockId: "s1", offset: 0 }, end: { blockId: "s1", offset: 1 } } } });

    const renamed = mergeDocuments(dest, source, { sectionBreak: "none" });
    expect(Object.keys(renamed.doc.bookmarks!)).toHaveLength(2);
    const newName = renamed.idMap.bookmarks["bm"];
    expect(newName).not.toBe("bm");
    // The renamed bookmark points at the (remapped) source block, which exists.
    const pos = renamed.doc.bookmarks![newName!]!.start;
    expect(allBlockIds(renamed.doc)).toContain(pos.blockId);

    const dropped = mergeDocuments(dest, source, { sectionBreak: "none", renameBookmarksOnCollision: false });
    expect(Object.keys(dropped.doc.bookmarks!)).toHaveLength(1);
    expect(dropped.warnings.some((w) => w.code === "bookmark-dropped")).toBe(true);
  });

  it("remaps a tocEntry target to the merged heading id", () => {
    const source = doc([
      para("h1", "Heading", { outlineLevel: 0 }),
      para("toc", "Heading", { tocEntry: { targetId: "h1", level: 0 } }),
    ]);
    const dest = doc([para("d1", "D")]);

    const { doc: m, idMap } = mergeDocuments(dest, source, { sectionBreak: "none" });

    const tocPara = m.blocks.find((b) => b.kind === "paragraph" && b.style.tocEntry) as Paragraph;
    expect(tocPara.style.tocEntry!.targetId).toBe(idMap.blocks["h1"]);
    expect(allBlockIds(m)).toContain(tocPara.style.tocEntry!.targetId);
  });

  it("does not mutate the inputs", () => {
    const dest = doc([para("d1", "D")], { sdts: { a: { type: "richText" } } });
    const sp = para("d1", "S");
    sp.sdtPath = ["a"];
    const source = doc([sp], { sdts: { a: { type: "plainText" } } });
    const destBefore = JSON.stringify(dest);
    const sourceBefore = JSON.stringify(source);

    mergeDocuments(dest, source);

    expect(JSON.stringify(dest)).toBe(destBefore);
    expect(JSON.stringify(source)).toBe(sourceBefore);
  });
});

describe("mergeAll", () => {
  it("folds N documents left-to-right with a seam between each", () => {
    const a = doc([para("a1", "A")]);
    const b = doc([para("b1", "B")]);
    const c = doc([para("c1", "C")]);

    const { doc: m } = mergeAll([a, b, c]);

    // a, seam, b, seam, c
    expect(m.blocks).toHaveLength(5);
    const texts = m.blocks.filter((x) => x.kind === "paragraph").map((p) => (p as Paragraph).runs.map((r) => r.text).join(""));
    expect(texts).toEqual(["A", "", "B", "", "C"]);
  });

  it("throws on an empty list", () => {
    expect(() => mergeAll([])).toThrow();
  });
});
