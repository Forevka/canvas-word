// Issue #214 (Part 1 of #206) — drawing shapes must survive a full .docx round-trip:
//   • ShapeBlock (rect / ellipse / line) with a solid fill + solid outline
//   • bare DrawingML wps:wsp (no VML / mc:AlternateContent)
//   • explicit no-fill / no-outline, and wp14 drawing identity
// Exercised both from a hand-written wps:wsp docx (the importer is the oracle) and
// through a model → writeDocx → re-import cycle.
import { describe, expect, it } from "vitest";
import type { Document, SectionProps, ShapeBlock } from "@cw/shared";
import { runImport } from "../../import/docx/pipeline";
import { simpleDocx } from "../../import/docx/fixture";
import { writeDocx } from "./writeDocx";
import { strFromU8, unzipSync } from "fflate";

const SECTION: SectionProps = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };

let n = 0;
const shape = (patch: Partial<ShapeBlock> = {}): ShapeBlock => ({
  kind: "shape", id: `s${n++}`, revision: 0, geometry: { preset: "rect" },
  widthPx: 220, heightPx: 90, align: "left", ...patch,
});
const docOf = (s: ShapeBlock): Document => ({ section: SECTION, blocks: [s] });
const roundTrip = (doc: Document): Document => runImport(writeDocx(doc).bytes).doc;
const exportedDocXml = (doc: Document): string => strFromU8(unzipSync(writeDocx(doc).bytes)["word/document.xml"]!);
const firstShape = (d: Document): ShapeBlock => d.blocks.find((b): b is ShapeBlock => b.kind === "shape")!;

describe("drawing shapes — DrawingML wps:wsp round-trip", () => {
  it("round-trips a filled rectangle (preset/size/align/fill/stroke)", () => {
    const s = shape({ geometry: { preset: "rect" }, align: "center", fill: { color: "#bcd6ef" }, stroke: { color: "#41719c", widthPt: 1.5 } });
    const out = firstShape(roundTrip(docOf(s)));
    expect(out.geometry.preset).toBe("rect");
    expect(out.widthPx).toBe(220);
    expect(out.heightPx).toBe(90);
    expect(out.align).toBe("center");
    expect(out.fill).toEqual({ color: "#bcd6ef" });
    expect(out.stroke).toEqual({ color: "#41719c", widthPt: 1.5 });
  });

  it("round-trips the ellipse and line presets", () => {
    expect(firstShape(roundTrip(docOf(shape({ geometry: { preset: "ellipse" } })))).geometry.preset).toBe("ellipse");
    expect(firstShape(roundTrip(docOf(shape({ geometry: { preset: "line" } })))).geometry.preset).toBe("line");
  });

  it("round-trips explicit no-fill and no-outline", () => {
    const s = shape({ fill: { none: true }, stroke: { none: true } });
    const out = firstShape(roundTrip(docOf(s)));
    expect(out.fill).toEqual({ none: true });
    expect(out.stroke).toEqual({ none: true });
  });

  it("round-trips the wp14 drawing identity (anchorId/editId)", () => {
    const s = shape({ drawingId: { anchorId: "12AB34CD", editId: "55667788" } });
    expect(firstShape(roundTrip(docOf(s))).drawingId).toEqual({ anchorId: "12AB34CD", editId: "55667788" });
  });

  it("emits bare DrawingML wps:wsp (no VML / AlternateContent) with prstGeom + fill + outline", () => {
    const xml = exportedDocXml(docOf(shape({ geometry: { preset: "ellipse" }, fill: { color: "#f4cccc" }, stroke: { color: "#cc4125", widthPt: 1 } })));
    expect(xml).toContain("<wps:wsp>");
    expect(xml).toContain('uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"');
    expect(xml).toContain('<a:prstGeom prst="ellipse">');
    expect(xml).toContain('<a:srgbClr val="f4cccc"/>');
    expect(xml).toContain('<a:srgbClr val="cc4125"/>');
    expect(xml).not.toContain("mc:AlternateContent");
    expect(xml).not.toContain("<v:");
    // wps namespace is declared on the document root.
    expect(xml).toContain('xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"');
  });

  it("imports a hand-written wps:wsp shape (the importer as oracle)", () => {
    const drawing = `<w:p><w:r><w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="2095500" cy="857250"/>
        <wp:docPr id="1" name="Shape"/>
        <a:graphic>
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp>
              <wps:cNvPr id="1" name="Shape"/>
              <wps:cNvSpPr/>
              <wps:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="2095500" cy="857250"/></a:xfrm>
                <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
                <a:solidFill><a:srgbClr val="bcd6ef"/></a:solidFill>
                <a:ln w="19050"><a:solidFill><a:srgbClr val="41719c"/></a:solidFill></a:ln>
              </wps:spPr>
              <wps:bodyPr/>
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing></w:r></w:p>`;
    const out = firstShape(runImport(simpleDocx(drawing)).doc);
    expect(out.geometry.preset).toBe("ellipse");
    expect(out.widthPx).toBe(220);
    expect(out.heightPx).toBe(90);
    expect(out.fill).toEqual({ color: "#bcd6ef" });
    expect(out.stroke).toEqual({ color: "#41719c", widthPt: 1.5 });
  });

  it("maps an unsupported preset to a rectangle (PR 1 preset set)", () => {
    const drawing = `<w:p><w:r><w:drawing>
      <wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Shape"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp><wps:cNvSpPr/><wps:spPr>
            <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="d9e2f0"/></a:solidFill>
          </wps:spPr><wps:bodyPr/></wps:wsp>
        </a:graphicData></a:graphic>
      </wp:inline>
    </w:drawing></w:r></w:p>`;
    const res = runImport(simpleDocx(drawing));
    const out = firstShape(res.doc);
    expect(out.geometry.preset).toBe("rect");
  });
});
