// Builder core: shape of the built Document, style resolution + precedence,
// chaining scope semantics, and build() isolation. Pure Node — no DOM.

import { describe, expect, it } from "vitest";
import type { Paragraph, TableBlock } from "@cw/shared";
import { DEFAULT_BULLET_LIST_ID, DEFAULT_NUMBER_LIST_ID } from "@cw/shared";
import { DocumentBuilder } from "./index";

const para = (b: unknown): Paragraph => b as Paragraph;
const table = (b: unknown): TableBlock => b as TableBlock;

describe("paragraphs and runs", () => {
  it("builds concrete runs with chain-order precedence", () => {
    const doc = DocumentBuilder.create({ idSeed: "t" })
      .paragraph("Title").withStyle("Heading1")
      .paragraph("Body text")
      .paragraph("Late override").withStyle("Heading1").fontSize(13)
      .build();

    expect(doc.blocks).toHaveLength(3);
    const [h1, body, overridden] = doc.blocks.map(para);

    expect(h1!.style.namedStyle).toBe("Heading1");
    expect(h1!.runs[0]!.style.fontSizePx).toBe(24); // resolved through basedOn chain
    expect(h1!.runs[0]!.style.bold).toBe(true);
    expect(h1!.style.keepWithNext).toBe(true);

    expect(body!.style.namedStyle).toBeUndefined();
    expect(body!.runs[0]!.text).toBe("Body text");

    // direct formatting AFTER withStyle wins
    expect(overridden!.runs[0]!.style.fontSizePx).toBe(13);
    expect(overridden!.runs[0]!.style.bold).toBe(true); // untouched style field survives
  });

  it("withStyle on an unknown id warns and is a no-op", () => {
    const b = DocumentBuilder.create();
    b.paragraph("x").withStyle("NoSuchStyle");
    const doc = b.build();
    expect(para(doc.blocks[0]).style.namedStyle).toBeUndefined();
    expect(b.warnings.some((w) => w.code === "style-missing:NoSuchStyle")).toBe(true);
  });

  it("text() replaces the empty placeholder run and inherits scope formatting", () => {
    const doc = DocumentBuilder.create()
      .paragraph().bold().text("first").text("second", { bold: false })
      .build();
    const runs = para(doc.blocks[0]).runs;
    expect(runs.map((r) => r.text)).toEqual(["first", "second"]);
    expect(runs[0]!.style.bold).toBe(true); // inherited from .bold()
    expect(runs[1]!.style.bold).toBe(false); // explicit patch wins
  });

  it(".shape() with a freeform path builds a custom geometry (and preset is the fallback)", () => {
    const doc = DocumentBuilder.create()
      .shape("rect", {
        widthPx: 100, heightPx: 100,
        path: [
          { type: "moveTo", x: 0, y: 0 },
          { type: "lineTo", x: 1, y: 1 },
          { type: "close" },
        ],
      })
      .build();
    const shape = doc.blocks[0] as { geometry: { preset: string; custom?: { segments: unknown[] } } };
    expect(shape.geometry.preset).toBe("rect");
    expect(shape.geometry.custom?.segments).toEqual([
      { type: "moveTo", x: 0, y: 0 },
      { type: "lineTo", x: 1, y: 1 },
      { type: "close" },
    ]);
  });

  it(".shape() prepends an implicit moveTo(0,0) when a path doesn't start with one", () => {
    const doc = DocumentBuilder.create()
      .shape("rect", {
        widthPx: 100, heightPx: 100,
        path: [
          { type: "lineTo", x: 1, y: 0 },
          { type: "lineTo", x: 1, y: 1 },
          { type: "close" },
        ],
      })
      .build();
    const shape = doc.blocks[0] as { geometry: { custom?: { segments: Array<{ type: string }> } } };
    expect(shape.geometry.custom?.segments[0]).toEqual({ type: "moveTo", x: 0, y: 0 });
    expect(shape.geometry.custom?.segments.map((s) => s.type)).toEqual(["moveTo", "lineTo", "lineTo", "close"]);
  });

  it(".shape() authors an absolute float anchor (behind/in-front + z-order)", () => {
    const doc = DocumentBuilder.create()
      .shape("rect", {
        widthPx: 150, heightPx: 90,
        anchor: { behind: true, offsetXPx: 20, offsetYPx: 120, relFromH: "margin", relFromV: "paragraph", z: -1 },
      })
      .shape("ellipse", {
        widthPx: 80, heightPx: 80,
        anchor: { behind: false, offsetXPx: 230, offsetYPx: 110, relFromH: "margin", relFromV: "paragraph", z: 3 },
      })
      .build();
    const behind = doc.blocks[0] as { anchor?: { behind: boolean; z?: number }; wrap?: string };
    const front = doc.blocks[1] as { anchor?: { behind: boolean; z?: number } };
    expect(behind.anchor).toMatchObject({ behind: true, offsetXPx: 20, offsetYPx: 120, z: -1 });
    expect(behind.wrap).toBeUndefined(); // anchor and wrap are mutually exclusive
    expect(front.anchor).toMatchObject({ behind: false, z: 3 });
  });

  it(".shape() anchor wins over wrap when both are set", () => {
    const doc = DocumentBuilder.create()
      .shape("rect", {
        widthPx: 100, heightPx: 60, wrap: "square",
        anchor: { behind: false, offsetXPx: 0, offsetYPx: 0, relFromH: "margin", relFromV: "paragraph" },
      })
      .build();
    const s = doc.blocks[0] as { anchor?: unknown; wrap?: string };
    expect(s.anchor).toBeDefined();
    expect(s.wrap).toBeUndefined();
  });

  it(".shapeGroup() authors a NESTED group (a group inside a group)", () => {
    const doc = DocumentBuilder.create()
      .shapeGroup({
        widthPx: 300, heightPx: 120, align: "center",
        children: [
          { preset: "rect", xPx: 0, yPx: 0, widthPx: 120, heightPx: 60, stroke: { color: "#333333", widthPt: 1 } },
          {
            preset: "rect", xPx: 140, yPx: 10, widthPx: 120, heightPx: 60,
            group: {
              children: [
                { preset: "ellipse", xPx: 0, yPx: 0, widthPx: 50, heightPx: 50, fill: { color: "#ead1dc" } },
                { preset: "triangle", xPx: 60, yPx: 0, widthPx: 50, heightPx: 50 },
              ],
            },
          },
        ],
      })
      .build();
    const outer = doc.blocks[0] as {
      group?: { children: { shape: { geometry: { preset: string }; group?: { children: { shape: { geometry: { preset: string } } }[] } } }[] };
    };
    expect(outer.group?.children).toHaveLength(2);
    const nested = outer.group!.children[1]!.shape;
    expect(nested.group).toBeDefined();
    expect(nested.group!.children.map((c) => c.shape.geometry.preset)).toEqual(["ellipse", "triangle"]);
    // the nested group's own box is the child's widthPx×heightPx.
    expect(nested).toMatchObject({ widthPx: 120, heightPx: 60 });
  });

  it("char formatting patches every run already in the paragraph", () => {
    const doc = DocumentBuilder.create()
      .paragraph("a").text("b").color("#ff0000")
      .build();
    for (const r of para(doc.blocks[0]).runs) expect(r.style.color).toBe("#ff0000");
  });

  it("paragraph-level styling: align, spacing, indent, link", () => {
    const doc = DocumentBuilder.create()
      .paragraph("x").align("center").spacing({ before: 10, after: 20, lineHeight: 2 }).indent({ left: 30, firstLine: 15 }).link("https://example.com")
      .build();
    const p = para(doc.blocks[0]);
    expect(p.style).toMatchObject({ align: "center", spaceBeforePx: 10, spaceAfterPx: 20, lineHeight: 2, indentLeftPx: 30, indentFirstLinePx: 15 });
    expect(p.runs[0]!.style.link).toBe("https://example.com");
  });

  it("spacing() sets fixed line spacing (lineRule + lineHeightPx)", () => {
    const doc = DocumentBuilder.create()
      .paragraph("exact").spacing({ lineRule: "exact", lineHeightPx: 28 })
      .paragraph("atLeast").spacing({ lineRule: "atLeast", lineHeightPx: 24 })
      .build();
    expect(para(doc.blocks[0]).style).toMatchObject({ lineRule: "exact", lineHeightPx: 28 });
    expect(para(doc.blocks[1]).style).toMatchObject({ lineRule: "atLeast", lineHeightPx: 24 });
  });

  it("spacing({ lineRule }) without lineHeightPx warns and is a no-op", () => {
    const b = DocumentBuilder.create();
    b.paragraph("x").spacing({ lineRule: "exact" });
    const p = para(b.build().blocks[0]);
    expect(p.style.lineRule).toBeUndefined();
    expect(p.style.lineHeightPx).toBeUndefined();
    expect(b.warnings.some((w) => w.code === "spacing-line-height-missing")).toBe(true);
  });

  it("reverting to multiplier spacing clears a previously-set fixed rule", () => {
    const doc = DocumentBuilder.create()
      .paragraph("x").spacing({ lineRule: "exact", lineHeightPx: 28 }).spacing({ lineHeight: 1.5 })
      .build();
    const p = para(doc.blocks[0]);
    expect(p.style.lineRule).toBeUndefined();
    expect(p.style.lineHeightPx).toBeUndefined();
    expect(p.style.lineHeight).toBe(1.5);
  });

  it("mints unique ids; idSeed makes them deterministic", () => {
    const build = (): string[] =>
      DocumentBuilder.create({ idSeed: "seed" }).paragraph("a").paragraph("b").table([["c"]]).build().blocks.map((b) => b.id);
    const ids = build();
    expect(new Set(ids).size).toBe(ids.length);
    expect(build()).toEqual(ids); // same seed → same ids
    for (const id of ids) expect(id.startsWith("seed-")).toBe(true);
  });
});

describe("tables", () => {
  it("data-driven shape with headerRow and normalized colFractions", () => {
    const doc = DocumentBuilder.create()
      .table([["Name", "Qty"], ["Widget", { text: "3", align: "right" }]], { headerRow: true, colFractions: [3, 1] })
      .build();
    const t = table(doc.blocks[0]);
    expect(t.rows).toHaveLength(2);
    const headerCell = para(t.rows[0]!.cells[0]!.blocks[0]);
    expect(headerCell.runs[0]!.style.bold).toBe(true);
    const qty = para(t.rows[1]!.cells[1]!.blocks[0]);
    expect(qty.style.align).toBe("right");
    expect(t.colFractions).toEqual([0.75, 0.25]);
  });

  it("structural shape: spans, shading, and full block stories in cells", () => {
    const doc = DocumentBuilder.create()
      .table((t) =>
        t
          .row((r) => r.cell("spans two", { colSpan: 2, shading: "#eeeeee" }))
          .row((r) => r.cell("a").cell((s) => s.paragraph("one").paragraph("two").bulletList(["x"]))),
      )
      .build();
    const t = table(doc.blocks[0]);
    expect(t.rows[0]!.cells[0]!.colSpan).toBe(2);
    expect(t.rows[0]!.cells[0]!.shading).toBe("#eeeeee");
    expect(t.rows[1]!.cells[1]!.blocks).toHaveLength(3);
  });

  it("records table-level defaults and cascades shading + margin onto cells (issue #48)", () => {
    const rule = { color: "#1a73e8", widthPx: 1 };
    const doc = DocumentBuilder.create()
      .table([["A", "B"], ["c", { text: "d", shading: "#ffffff" }]], {
        borders: { top: rule, right: rule, bottom: rule, left: rule, insideH: rule, insideV: rule },
        shading: "#eef5ff",
        cellMargin: { top: 6, right: 10, bottom: 6, left: 10 },
      })
      .build();
    const t = table(doc.blocks[0]);
    // Table-level defaults live on the model (so export re-emits them at tblPr).
    expect(t.defaultBorders?.insideH?.color).toBe("#1a73e8");
    expect(t.defaultShading).toBe("#eef5ff");
    expect(t.defaultCellMargin).toEqual({ top: 6, right: 10, bottom: 6, left: 10 });
    // Shading + margin + borders cascade onto cells that didn't set their own…
    expect(t.rows[0]!.cells[0]!.shading).toBe("#eef5ff");
    expect(t.rows[0]!.cells[0]!.margin).toEqual({ top: 6, right: 10, bottom: 6, left: 10 });
    const corner = t.rows[0]!.cells[0]!.borders!;
    expect(corner.top).toEqual(rule);
    expect(corner.left).toEqual(rule);
    expect(corner.right).toEqual(rule); // interior insideV (same rule here)
    expect(corner.bottom).toEqual(rule); // interior insideH
    // …but an explicit per-cell value still wins.
    expect(t.rows[1]!.cells[1]!.shading).toBe("#ffffff");
  });

  it("an empty table warns and self-heals to one cell", () => {
    const b = DocumentBuilder.create();
    b.table((t) => void t);
    const doc = b.build();
    expect(table(doc.blocks[0]).rows[0]!.cells).toHaveLength(1);
    expect(b.warnings.some((w) => w.code === "table-empty")).toBe(true);
  });
});

describe("images", () => {
  it("inlines raw bytes as a data: URL", () => {
    const doc = DocumentBuilder.create()
      .image({ data: new Uint8Array([1, 2, 3]), mime: "image/png" }, { widthPx: 100, heightPx: 50, align: "center" })
      .build();
    const img = doc.blocks[0]!;
    expect(img.kind).toBe("image");
    if (img.kind !== "image") return;
    expect(img.src).toBe("data:image/png;base64,AQID");
    expect(img).toMatchObject({ widthPx: 100, heightPx: 50, align: "center" });
  });

  it("requires explicit dimensions (warns and skips the image)", () => {
    // Warn-and-continue (not throw) — the builder's uniform invalid-input contract.
    const b = DocumentBuilder.create();
    const doc = b.image("data:image/png;base64,", { widthPx: 0, heightPx: 10 }).build();
    expect(b.warnings.some((w) => w.code === "image-size-invalid")).toBe(true);
    expect(doc.blocks.some((blk) => blk.kind === "image")).toBe(false);
  });
});

describe("lists", () => {
  it("registers default definitions and marks list membership with levels", () => {
    const doc = DocumentBuilder.create()
      .bulletList(["a", { text: "b", level: 1 }])
      .numberedList(["one"])
      .build();
    expect(doc.lists?.[DEFAULT_BULLET_LIST_ID]).toBeDefined();
    expect(doc.lists?.[DEFAULT_NUMBER_LIST_ID]).toBeDefined();
    expect(para(doc.blocks[0]).style.list).toEqual({ listId: DEFAULT_BULLET_LIST_ID, level: 0 });
    expect(para(doc.blocks[1]).style.list).toEqual({ listId: DEFAULT_BULLET_LIST_ID, level: 1 });
    expect(para(doc.blocks[2]).style.list).toEqual({ listId: DEFAULT_NUMBER_LIST_ID, level: 0 });
  });

  it("an unknown listId warns and registers a default definition under that id", () => {
    const b = DocumentBuilder.create();
    b.list(["x"], { listId: "tmpl-list", kind: "number" });
    const doc = b.build();
    expect(doc.lists?.["tmpl-list"]?.id).toBe("tmpl-list");
    expect(para(doc.blocks[0]).style.list?.listId).toBe("tmpl-list");
    expect(b.warnings.some((w) => w.code === "list-missing:tmpl-list")).toBe(true);
  });
});

describe("page breaks", () => {
  it("lands on the next paragraph, injects one before a table, and flushes at build", () => {
    const doc = DocumentBuilder.create()
      .paragraph("p1").pageBreak()
      .paragraph("p2").pageBreak()
      .table([["t"]])
      .pageBreak()
      .build();
    expect(para(doc.blocks[1]).style.pageBreakBefore).toBe(true);
    // break before the table → injected empty paragraph carries it
    expect(para(doc.blocks[2]).style.pageBreakBefore).toBe(true);
    expect(doc.blocks[3]!.kind).toBe("table");
    // trailing break materializes
    expect(para(doc.blocks[4]).style.pageBreakBefore).toBe(true);
  });
});

describe("headers, footers, page setup", () => {
  it("composes band stories including variants", () => {
    const doc = DocumentBuilder.create()
      .header((h) => h.paragraph("Acme Corp").align("right"))
      .footer((f) => f.paragraph("Page {page} of {pages}").align("center"))
      .header((h) => h.paragraph("First page only"), { variant: "first" })
      .build();
    expect(para(doc.section.header?.[0]).runs[0]!.text).toBe("Acme Corp");
    expect(para(doc.section.footer?.[0]).runs[0]!.text).toBe("Page {page} of {pages}");
    expect(para(doc.section.headerFirst?.[0]).runs[0]!.text).toBe("First page only");
    expect(doc.section.headerEven).toBeUndefined();
  });

  it("merges page setup: named size, landscape swap, partial margins", () => {
    const doc = DocumentBuilder.create()
      .pageSetup({ pageSize: "A4", orientation: "landscape", margins: { top: 48 }, pageNumberStart: 5 })
      .build();
    expect(doc.section.pageWidthPx).toBe(1123);
    expect(doc.section.pageHeightPx).toBe(794);
    expect(doc.section.marginPx).toMatchObject({ top: 48, left: 96 });
    expect(doc.section.pageNumberStart).toBe(5);
  });
});

describe("custom styles", () => {
  it("style() registers a named style usable by withStyle()", () => {
    const doc = DocumentBuilder.create()
      .style({ id: "Brand", name: "Brand", basedOn: "Normal", char: { color: "#123456", bold: true }, para: { align: "center" } })
      .paragraph("hello").withStyle("Brand")
      .build();
    const p = para(doc.blocks[0]);
    expect(p.style.namedStyle).toBe("Brand");
    expect(p.style.align).toBe("center");
    expect(p.runs[0]!.style.color).toBe("#123456");
  });
});

describe("build()", () => {
  it("returns isolated clones; the builder stays usable", () => {
    const b = DocumentBuilder.create().paragraph("x").end();
    const first = b.build();
    para(first.blocks[0]).runs[0]!.text = "mutated";
    const second = b.build();
    expect(para(second.blocks[0]).runs[0]!.text).toBe("x");
    b.paragraph("y");
    expect(b.build().blocks).toHaveLength(2);
    expect(first.blocks).toHaveLength(1);
  });

  it("an empty build still has a caret home", () => {
    const doc = DocumentBuilder.create().build();
    expect(doc.blocks).toHaveLength(1);
    expect(para(doc.blocks[0]).runs[0]!.text).toBe("");
  });
});
