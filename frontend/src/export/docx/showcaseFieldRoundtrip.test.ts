// The flagship showcase exercises every field type — inline in the body AND inside
// table cells — and one of those field paragraphs also carries a bookmark. That
// combination regressed once: the bookmark-splicing path in paragraphXml emitted
// runs one-by-one and dropped the inline-field wrapping, so the fields re-imported
// as plain text. This test pins that every field object survives a full round-trip
// (and that a second round-trip is idempotent — no drift).
import { beforeAll, describe, expect, it } from "vitest";
import type { Document, FieldDef } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { runExport } from "../pipeline";
import { installMeasureHost } from "../shared/measureHost";
import { sampleDoc } from "../../model/sampleDoc";

const fieldTypes = (doc: Document): string[] =>
  (Object.values(doc.fields ?? {}) as FieldDef[])
    .map((d) => (d.kind === "builtin" ? `builtin:${d.spec?.type}` : "custom"))
    .sort();

describe("showcase field objects round-trip", () => {
  beforeAll(async () => {
    await installMeasureHost();
  });

  it("preserves every field (body + table cells) across export→import, idempotently", async () => {
    const orig = sampleDoc();
    const want = fieldTypes(orig);
    // The showcase has 3×DATE, 4×PAGE, 1×IF (body inline, inside table cells, and
    // in the CJK/RTL section's bidi table cell + RTL field paragraph), plus a
    // REF/PAGEREF cross-reference pair pointing at the "sample" bookmark.
    expect(want).toEqual([
      "builtin:DATE", "builtin:DATE", "builtin:DATE", "builtin:IF",
      "builtin:PAGE", "builtin:PAGE", "builtin:PAGE", "builtin:PAGE",
      "builtin:PAGEREF", "builtin:REF",
    ]);

    const r1 = runImport((await runExport(orig, "docx")).bytes).doc;
    expect(fieldTypes(r1)).toEqual(want); // nothing dropped on round-trip

    const r2 = runImport((await runExport(r1, "docx")).bytes).doc;
    expect(fieldTypes(r2)).toEqual(want); // stable — no drift on repeated open/save
  });
});
