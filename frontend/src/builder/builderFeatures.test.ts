// Builder feature surface added on top of the core: inline fields, content
// controls, footnotes, bookmarks, cross-references, rich-text run gaps, TOC,
// sections/columns, custom styles + list definitions, and table-style presets.
// Pure Node — no DOM. `now` is injected so date/time fields are deterministic.

import { describe, expect, it } from "vitest";
import type { FieldDef, Paragraph, TableBlock } from "@cw/shared";
import { formatFieldDate } from "@cw/shared";
import { DocumentBuilder } from "./index";

const para = (b: unknown): Paragraph => b as Paragraph;
const table = (b: unknown): TableBlock => b as TableBlock;
const NOW = new Date(2026, 5, 16, 14, 30); // 2026-06-16 14:30 local

describe("inline fields", () => {
  it("pageField emits a {page} token tagged with a builtin PAGE field def", () => {
    const doc = DocumentBuilder.create({ idSeed: "t" }).paragraph("p. ").pageField().build();
    const runs = para(doc.blocks[0]).runs;
    const tok = runs.find((r) => r.text === "{page}")!;
    expect(tok.style.fieldId).toBeDefined();
    const def = doc.fields![tok.style.fieldId!] as FieldDef;
    expect(def.kind).toBe("builtin");
    expect(def.spec).toEqual({ type: "PAGE" });
    expect(def.instruction).toContain("PAGE");
  });

  it("pageField(roman) uses the {page:roman} token", () => {
    const doc = DocumentBuilder.create().paragraph().pageField("roman").build();
    expect(para(doc.blocks[0]).runs[0]!.text).toBe("{page:roman}");
  });

  it("dateField materializes via the shared formatter (deterministic now)", () => {
    const doc = DocumentBuilder.create().paragraph().dateField("MMMM d, yyyy", { now: NOW }).build();
    expect(para(doc.blocks[0]).runs[0]!.text).toBe(formatFieldDate(NOW, "MMMM d, yyyy"));
  });

  it("ifField materializes the chosen plain-text branch", () => {
    const doc = DocumentBuilder.create().paragraph().ifField("2", ">", "1", "held", "nope").build();
    const run = para(doc.blocks[0]).runs[0]!;
    expect(run.text).toBe("held");
    const def = doc.fields![run.style.fieldId!] as FieldDef;
    expect(def.spec!.type).toBe("IF");
  });

  it("customField registers a non-builtin field with a verbatim instruction", () => {
    const doc = DocumentBuilder.create().paragraph().customField(" SEQ Figure ", "3", { name: "SEQ" }).build();
    const def = doc.fields![para(doc.blocks[0]).runs[0]!.style.fieldId!] as FieldDef;
    expect(def.kind).toBe("custom");
    expect(def.name).toBe("SEQ");
    expect(def.spec).toBeUndefined();
  });

  it("crossReference emits a PAGEREF field over a bookmark", () => {
    const doc = DocumentBuilder.create().paragraph().crossReference("intro", { kind: "pageRef" }).build();
    const def = doc.fields![para(doc.blocks[0]).runs[0]!.style.fieldId!] as FieldDef;
    expect(def.kind).toBe("custom");
    expect(def.instruction).toContain("PAGEREF intro");
  });
});

describe("content controls", () => {
  it("dropDown registers SdtProps and tags the run", () => {
    const doc = DocumentBuilder.create()
      .paragraph()
      .dropDown("One", [{ display: "One", value: "1" }, { display: "Two", value: "2" }], { alias: "Choice" })
      .build();
    const run = para(doc.blocks[0]).runs[0]!;
    expect(run.text).toBe("One");
    const props = doc.sdts![run.style.sdtPath!.at(-1)!]!;
    expect(props.type).toBe("dropDown");
    expect(props.listItems).toHaveLength(2);
    expect(props.alias).toBe("Choice");
  });

  it("checkbox glyph reflects the checked state", () => {
    const doc = DocumentBuilder.create().paragraph().checkbox(true).checkbox(false).build();
    const [a, b] = para(doc.blocks[0]).runs;
    expect(a!.text).toBe("☒");
    expect(b!.text).toBe("☐");
  });
});

describe("rich-text run gaps", () => {
  it("superscript/subscript/letterSpacing/hidden patch the run style", () => {
    const doc = DocumentBuilder.create()
      .paragraph().text("x").superscript()
      .paragraph().text("y").letterSpacing(2).hidden()
      .build();
    expect(para(doc.blocks[0]).runs[0]!.style.verticalAlign).toBe("super");
    expect(para(doc.blocks[1]).runs[0]!.style.letterSpacingPx).toBe(2);
    expect(para(doc.blocks[1]).runs[0]!.style.hidden).toBe(true);
  });

  it("toggling superscript off clears it", () => {
    const doc = DocumentBuilder.create().paragraph().text("x").superscript().superscript(false).build();
    expect(para(doc.blocks[0]).runs[0]!.style.verticalAlign).toBeUndefined();
  });
});

describe("footnotes", () => {
  it("auto-numbers markers and registers note bodies", () => {
    const doc = DocumentBuilder.create()
      .paragraph("First").footnote("first note")
      .paragraph("Second").footnote((s) => s.paragraph("a").paragraph("b"))
      .build();
    const m1 = para(doc.blocks[0]).runs.find((r) => r.style.footnoteRef)!;
    const m2 = para(doc.blocks[1]).runs.find((r) => r.style.footnoteRef)!;
    expect(m1.text).toBe("1");
    expect(m2.text).toBe("2");
    expect(m1.style.verticalAlign).toBe("super");
    expect(doc.footnotes![m1.style.footnoteRef!]).toHaveLength(1);
    expect(doc.footnotes![m2.style.footnoteRef!]).toHaveLength(2); // multi-paragraph callback
  });
});

describe("bookmarks", () => {
  it("bookmark spans exactly the appended run's offsets", () => {
    const doc = DocumentBuilder.create().paragraph("hello ").bookmark("mark", "TARGET").text(" tail").build();
    const p = para(doc.blocks[0]);
    expect(doc.bookmarks!.mark).toEqual({
      start: { blockId: p.id, offset: 6 },
      end: { blockId: p.id, offset: 12 },
    });
  });

  it("bookmarkRange spans the paragraphs the callback adds", () => {
    const doc = DocumentBuilder.create()
      .paragraph("before")
      .bookmarkRange("sec", (s) => s.paragraph("one").paragraph("two and a bit"))
      .paragraph("after")
      .build();
    const ps = doc.blocks.map(para);
    const one = ps.find((p) => p.runs[0]!.text === "one")!;
    const two = ps.find((p) => p.runs[0]!.text === "two and a bit")!;
    expect(doc.bookmarks!.sec!.start).toEqual({ blockId: one.id, offset: 0 });
    expect(doc.bookmarks!.sec!.end).toEqual({ blockId: two.id, offset: "two and a bit".length });
  });
});

describe("table of contents", () => {
  it("collects headings added AFTER the call and replaces the anchor", () => {
    const b = DocumentBuilder.create({ idSeed: "t" })
      .tableOfContents()
      .paragraph("Alpha").withStyle("Heading1")
      .paragraph("Beta").withStyle("Heading2")
      .paragraph("body");
    const doc = b.build();
    const entries = doc.blocks.map(para).filter((p) => p.style.tocEntry);
    expect(entries.map((p) => p.runs.map((r) => r.text).join(""))).toEqual(["Alpha", "Beta"]);
    expect(doc.tocInstruction).toContain("TOC");
    // anchor placeholder is gone (replaced by the entries)
    expect(doc.blocks.map(para).some((p) => p.id === doc.tocAnchorBlockId && !p.style.tocEntry)).toBe(false);
    // idempotent: a second build regenerates identically
    const doc2 = b.build();
    expect(doc2.blocks.map(para).filter((p) => p.style.tocEntry)).toHaveLength(2);
  });

  it("no headings → does not crash, leaves a paragraph", () => {
    const doc = DocumentBuilder.create().tableOfContents().paragraph("plain").build();
    expect(doc.blocks.length).toBeGreaterThan(0);
  });
});

describe("sections and columns", () => {
  it("sectionBreak sets the break + column patch on the last paragraph", () => {
    const doc = DocumentBuilder.create()
      .paragraph("section one")
      .sectionBreak({ columns: { count: 2 }, pageNumberStart: 1 })
      .paragraph("section two")
      .build();
    const br = para(doc.blocks[0]).style.sectionBreak!;
    expect(br.type).toBe("nextPage");
    expect(br.props.columns).toEqual({ count: 2, gapPx: 48 });
    expect(br.props.pageNumberStart).toBe(1);
  });

  it("columnBreak sets columnBreakBefore on the next paragraph", () => {
    const doc = DocumentBuilder.create().paragraph("a").columnBreak().paragraph("b").build();
    expect(para(doc.blocks[0]).style.columnBreakBefore).toBeUndefined();
    expect(para(doc.blocks[1]).style.columnBreakBefore).toBe(true);
  });
});

describe("custom styles and list definitions", () => {
  it("defaultStyle makes later content inherit the chosen style", () => {
    const doc = DocumentBuilder.create().defaultStyle("Heading1").paragraph("x").build();
    expect(para(doc.blocks[0]).runs[0]!.style.bold).toBe(true); // Heading1 char
    expect(para(doc.blocks[0]).runs[0]!.style.fontSizePx).toBe(24);
  });

  it("listDefinition registers a custom numbered list referenced by id", () => {
    const b = DocumentBuilder.create().listDefinition("rom", { kind: "number", format: "upperRoman" });
    b.paragraph("intro").list(["one", "two"], { listId: "rom" });
    const doc = b.build();
    expect(doc.lists!.rom!.levels[0]!.format).toBe("upperRoman");
    expect(b.warnings.some((w) => w.code.startsWith("list-missing"))).toBe(false);
  });
});

describe("table-style presets", () => {
  it("applies header styling, borders and shading; explicit cell wins; unknown warns", () => {
    const b = DocumentBuilder.create().tableStylePreset("brand", {
      headerRow: true,
      headerShading: "#001a66",
      headerChar: { color: "#ffffff" },
      borders: { top: { color: "#ccc", widthPx: 1 } },
    });
    b.table(
      [
        ["H1", "H2"],
        [{ text: "own", shading: "#eee" }, "plain"],
      ],
      { style: "brand" },
    );
    const t = table(b.build().blocks[0]);
    const [header, body] = t.rows;
    expect(header!.cells[0]!.shading).toBe("#001a66"); // header shading
    expect(para(header!.cells[0]!.blocks[0]).runs[0]!.style.color).toBe("#ffffff"); // header char
    expect(para(header!.cells[0]!.blocks[0]).runs[0]!.style.bold).toBe(true);
    expect(body!.cells[0]!.shading).toBe("#eee"); // explicit CellSpec shading wins
    expect(body!.cells[0]!.borders).toBeDefined(); // preset borders fill in
    expect(body!.cells[1]!.borders).toBeDefined();

    b.table([["x"]], { style: "ghost" });
    expect(b.warnings.some((w) => w.code === "table-style-missing:ghost")).toBe(true);
  });
});
