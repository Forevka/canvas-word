// Exports the showcase document (`sampleDoc()`, exactly what the offline example
// ships) so CI can validate it against the real OOXML schema with the Open XML SDK —
// the check our importer-only round-trip tests cannot perform, since our importer is
// lenient and order-independent while Word is strict (issue #180 / #193 / #195).
//
// In a normal run this is just a cheap smoke test. When OOXML_OUT is set (the CI
// `ooxml-validate` job), it also writes the .docx to that path for tools/OoxmlValidate.
import { describe, expect, it } from "vitest";
import { writeFileSync } from "node:fs";
import { writeDocx } from "./writeDocx";
import { sampleDoc } from "../../model/sampleDoc";

describe("showcase docx artifact", () => {
  it("exports non-empty bytes (and writes them to OOXML_OUT for external SDK validation)", () => {
    const { bytes } = writeDocx(sampleDoc());
    expect(bytes.length).toBeGreaterThan(1000);
    const out = process.env.OOXML_OUT;
    if (out) writeFileSync(out, bytes);
  });
});
