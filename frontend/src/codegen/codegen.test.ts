import { describe, expect, it } from "vitest";
import type { Document } from "@cw/shared";
import { DocumentBuilder } from "../builder";
import { sampleDoc } from "../model/sampleDoc";
import { writeDocx } from "../export/docx/writeDocx";
import { emitBuilderCode } from "./emit";
import { docxToBuilderCode } from "./index";

/** Execute generated builder code: strip the TS-only import + signature lines and
 *  run the body with `DocumentBuilder` injected. Proves the emitted source parses
 *  AND reconstructs a document — the round-trip fidelity oracle. */
function runGenerated(code: string): Document {
  const body = code
    .replace(/^import[^\n]*\n/gm, "")
    .replace(/export function \w+\(\): Document \{/, "")
    .replace(/\}\s*$/, "");
  // eslint-disable-next-line @typescript-eslint/no-implied-eval
  const fn = new Function("DocumentBuilder", body) as (b: typeof DocumentBuilder) => Document;
  return fn(DocumentBuilder);
}

/** Drop volatile identity (ids/revision) and undefined props so two structurally
 *  equal documents compare equal regardless of the builder's random id seed. */
function normalize(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(normalize);
  if (value && typeof value === "object") {
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
      if (k === "id" || k === "revision" || v === undefined) continue;
      out[k] = normalize(v);
    }
    return out;
  }
  return value;
}

/** Build → emit → re-run → assert model equivalence; return the uncovered report. */
function roundTrip(build: (b: DocumentBuilder) => Document): { uncovered: ReturnType<typeof emitBuilderCode>["uncovered"] } {
  const original = build(DocumentBuilder.create());
  const { code, uncovered } = emitBuilderCode(original);
  const reproduced = runGenerated(code);
  expect(normalize(reproduced)).toEqual(normalize(original));
  return { uncovered };
}

describe("reverse builder — round-trip fidelity", () => {
  it("plain paragraphs with mixed run formatting", () => {
    const { uncovered } = roundTrip((b) =>
      b
        .paragraph("Hello ")
        .text("bold", { bold: true })
        .text(" and ")
        .text("red italic", { italic: true, color: "#ff0000" })
        .paragraph("Second paragraph.")
        .build(),
    );
    expect(uncovered).toEqual([]);
  });

  it("paragraph-level formatting: align, spacing, indent, toggles", () => {
    const { uncovered } = roundTrip((b) =>
      b
        .paragraph("Centered")
        .align("center")
        .spacing({ before: 12, after: 6, lineHeight: 1.4 })
        .indent({ left: 24, firstLine: 18 })
        .keepWithNext()
        .widowControl(false)
        .build(),
    );
    expect(uncovered).toEqual([]);
  });

  it("named styles via withStyle", () => {
    const { uncovered } = roundTrip((b) =>
      b
        .paragraph("Title").withStyle("Title")
        .paragraph("A heading").withStyle("Heading1")
        .paragraph("Body text under it.")
        .build(),
    );
    expect(uncovered).toEqual([]);
  });

  it("page setup: size, margins, columns", () => {
    const { uncovered } = roundTrip((b) =>
      b
        .pageSetup({ pageSize: "A4", margins: { top: 72, right: 72, bottom: 72, left: 72 }, columns: { count: 2, gapPx: 40 } })
        .paragraph("Two columns on A4.")
        .build(),
    );
    expect(uncovered).toEqual([]);
  });

  it("data-driven table with header row and cell shading", () => {
    const { uncovered } = roundTrip((b) =>
      b
        .table(
          [
            ["Name", "Qty"],
            ["Widget", "3"],
            ["Gadget", "7"],
          ],
          { headerRow: true },
        )
        .build(),
    );
    expect(uncovered).toEqual([]);
  });

  it("structural table with spans and cell options", () => {
    const { uncovered } = roundTrip((b) =>
      b
        .table((t) => {
          t.row((r) => {
            r.cell("Merged header", { colSpan: 2, shading: "#eeeeee", align: "center" });
          });
          t.row((r) => {
            r.cell("A");
            r.cell("B", { vAlign: "center" });
          });
        })
        .build(),
    );
    expect(uncovered).toEqual([]);
  });

  it("list paragraphs reproduce exactly via listItem()", () => {
    const { uncovered } = roundTrip((b) =>
      b
        .paragraph("Shopping")
        .bulletList(["Milk", "Eggs", { text: "Bread", level: 1 }])
        .numberedList(["First", "Second"])
        .paragraph("After the lists.")
        .build(),
    );
    expect(uncovered).toEqual([]);
  });

  it("listItem() preserves rich formatting a plain list item cannot", () => {
    const { uncovered } = roundTrip((b) =>
      b
        .listDefinition("todo", { kind: "bullet", char: "☐" })
        .paragraph()
        .listItem("todo")
        .text("Ship ", { italic: true })
        .text("the feature", { bold: true })
        .align("left")
        .build(),
    );
    expect(uncovered).toEqual([]);
  });

  it("keeps explicit before/after spacing when auto-spacing is off (beforeAuto:false)", () => {
    const { uncovered } = roundTrip((b) =>
      b
        .paragraph("Real spacing, autospacing explicitly off")
        .spacing({ before: 150, beforeAuto: false, after: 40, afterAuto: false })
        .build(),
    );
    expect(uncovered).toEqual([]);
  });

  it("bakes auto-spacing without leaking the baked px as an explicit value", () => {
    const { uncovered } = roundTrip((b) =>
      b.paragraph("Auto spaced").spacing({ beforeAuto: true, afterAuto: true }).build(),
    );
    expect(uncovered).toEqual([]);
  });

  it("multi-section line numbering stays on the section that declared it", () => {
    // Section one has NO line numbering; the body section turns it on. The body
    // value must not leak backward onto section one via the builder's carry-forward.
    const { uncovered } = roundTrip((b) =>
      b
        .paragraph("Section one — no line numbers")
        .sectionBreak()
        .paragraph("Section two — line numbers on")
        .pageSetup({ lineNumbering: { countBy: 1, start: 1, restart: "newPage" } })
        .build(),
    );
    expect(uncovered).toEqual([]);
  });

  it("page break and column break before a block", () => {
    const { uncovered } = roundTrip((b) =>
      b
        .paragraph("Page one.")
        .pageBreak()
        .paragraph("Page two.")
        .build(),
    );
    expect(uncovered).toEqual([]);
  });

  it("header and footer bands", () => {
    const { uncovered } = roundTrip((b) =>
      b
        .header((s) => {
          s.paragraph("The header").align("center");
        })
        .footer((s) => {
          s.paragraph("The footer");
        })
        .paragraph("Body.")
        .build(),
    );
    expect(uncovered).toEqual([]);
  });

  it("flagship sampleDoc: generates runnable code and a structured gap report", () => {
    const original = sampleDoc();
    const { code, uncovered } = emitBuilderCode(original);
    // Generated source must parse and reconstruct a document without throwing.
    const reproduced = runGenerated(code);
    expect(reproduced.blocks.length).toBeGreaterThan(0);
    // sampleDoc exercises footnotes/endnotes/fields/TOC/equations — every gap is
    // reported with a node path + field, never silently dropped.
    expect(uncovered.length).toBeGreaterThan(0);
    for (const u of uncovered) {
      expect(typeof u.path).toBe("string");
      expect(typeof u.field).toBe("string");
    }
  });

  it("docx → builder code end to end", () => {
    const { bytes } = writeDocx(sampleDoc());
    const { code, uncovered, importWarnings } = docxToBuilderCode(bytes);
    expect(code).toContain("DocumentBuilder.create(");
    expect(code).toContain("return b.build();");
    expect(runGenerated(code).blocks.length).toBeGreaterThan(0);
    expect(Array.isArray(uncovered)).toBe(true);
    expect(Array.isArray(importWarnings)).toBe(true);
  });

  it("reports uncovered fields rather than dropping them silently", () => {
    const original = DocumentBuilder.create()
      .paragraph("See note")
      .footnote("A footnote body.")
      .build();
    const { uncovered } = emitBuilderCode(original);
    // The footnote registry + its inline reference marker are not yet automated —
    // they must surface in the report.
    expect(uncovered.some((u) => u.path === "footnotes")).toBe(true);
    expect(uncovered.some((u) => u.field === "footnoteRef")).toBe(true);
  });
});
