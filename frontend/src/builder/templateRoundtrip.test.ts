// fromTemplate round-trip: a styled document built with the builder is exported
// to .docx (that's the template fixture — no binary committed), re-opened via
// fromTemplate, composed against, built, exported, and re-imported. Asserts the
// template's styles/section/lists carry over and the final structure survives
// a full export→import cycle. Pure Node.

import { describe, expect, it } from "vitest";
import type { Paragraph } from "@cw/shared";
import { runImport } from "../import/docx/pipeline";
import { runExport } from "../export/pipeline";
import { DocumentBuilder } from "./index";

const para = (b: unknown): Paragraph => b as Paragraph;

/** A "company template": custom style, custom page setup, branded header,
 *  and one body paragraph (which fromTemplate should discard by default). */
async function makeTemplate(): Promise<Uint8Array> {
  const doc = DocumentBuilder.create({ idSeed: "tpl", pageSize: "A4", margins: { top: 72, bottom: 72 } })
    .style({ id: "BrandHeading", name: "Brand Heading", basedOn: "Heading1", char: { color: "#aa0000" }, para: { spaceBeforePx: 0 } })
    .header((h) => h.paragraph("Acme Corp").align("right"))
    .footer((f) => f.paragraph("Page {page} of {pages}").align("center"))
    .paragraph("Template body placeholder — should be discarded").withStyle("BrandHeading")
    .build();
  return (await runExport(doc, "docx", {})).bytes;
}

describe("DocumentBuilder.fromTemplate", () => {
  it("carries the template's stylesheet, section, and bands; discards the body", async () => {
    const template = await makeTemplate();
    const b = await DocumentBuilder.fromTemplate(template, { idSeed: "gen" });

    // Body discarded, geometry + bands kept.
    expect(b.build().blocks.map(para).flatMap((p) => p.runs.map((r) => r.text)).join("")).toBe("");
    const built = b.build();
    expect(built.section.pageWidthPx).toBe(794); // A4 from the template
    expect(built.section.marginPx.top).toBe(72);
    expect(para(built.section.header?.[0]).runs[0]!.text).toBe("Acme Corp");
    expect(para(built.section.footer?.[0]).runs[0]!.text).toBe("Page {page} of {pages}");

    // The template's custom style resolves for new content.
    const doc = b
      .paragraph("Generated title").withStyle("BrandHeading")
      .paragraph("Body line")
      .table([["Item", "Qty"], ["Widget", "3"]], { headerRow: true })
      .bulletList(["alpha", "beta"])
      .build();
    const title = para(doc.blocks[0]);
    expect(title.style.namedStyle).toBe("BrandHeading");
    expect(title.runs[0]!.style.color).toBe("#aa0000");
    expect(title.runs[0]!.style.bold).toBe(true); // inherited via basedOn → Heading1

    // Builder ids live in their own namespace — disjoint from template-derived ids.
    expect(doc.blocks.every((blk) => blk.id.startsWith("gen-"))).toBe(true);
  });

  it("keepBody appends after the template's content", async () => {
    const template = await makeTemplate();
    const b = await DocumentBuilder.fromTemplate(template, { keepBody: true });
    const doc = b.paragraph("appended").build();
    expect(doc.blocks.length).toBeGreaterThanOrEqual(2);
    expect(para(doc.blocks[0]).runs.map((r) => r.text).join("")).toContain("Template body placeholder");
    expect(para(doc.blocks.at(-1)).runs[0]!.text).toBe("appended");
  });

  it("the built document survives export → re-import", async () => {
    const template = await makeTemplate();
    const b = await DocumentBuilder.fromTemplate(template);
    const doc = b
      .paragraph("Round trip").withStyle("BrandHeading")
      .paragraph("Plain text")
      .table([["a", "b"], ["c", "d"]])
      .numberedList(["one", "two"])
      .build();

    const docx = await runExport(doc, "docx", {});
    expect(docx.bytes[0]).toBe(0x50); // PK

    const reimported = runImport(docx.bytes).doc;
    expect(reimported.blocks.length).toBe(doc.blocks.length);
    const heading = para(reimported.blocks[0]);
    expect(heading.runs[0]!.text).toBe("Round trip");
    expect(heading.style.namedStyle).toBe("BrandHeading");
    expect(heading.runs[0]!.style.bold).toBe(true);
    expect(reimported.blocks[2]!.kind).toBe("table");
    expect(para(reimported.blocks[3]).style.list).toBeDefined();
    expect(para(reimported.section.header?.[0]).runs[0]!.text).toBe("Acme Corp");
  });
});
