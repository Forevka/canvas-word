// Issue #214 (Part 1 of #206) — drawing shapes must survive a full .docx round-trip:
//   • ShapeBlock (rect / ellipse / line) with a solid fill + solid outline
//   • bare DrawingML wps:wsp (no VML / mc:AlternateContent)
//   • explicit no-fill / no-outline, and wp14 drawing identity
// Exercised both from a hand-written wps:wsp docx (the importer is the oracle) and
// through a model → writeDocx → re-import cycle.
import { describe, expect, it } from "vitest";
import type { Document, Paragraph, SectionProps, ShapeBlock, ShapeGroupChild, ShapePath } from "@cw/shared";
import { DEFAULT_CHAR_STYLE, DEFAULT_PARA_STYLE, applyOp } from "@cw/shared";
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
const txtPara = (text: string): Paragraph => ({
  kind: "paragraph", id: `p${n++}`, revision: 0,
  runs: [{ text, style: { ...DEFAULT_CHAR_STYLE } }], style: { ...DEFAULT_PARA_STYLE },
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

  it("round-trips every preset in the Part 2 set", () => {
    for (const preset of ["rect", "roundRect", "ellipse", "triangle", "diamond", "rightArrow", "leftArrow", "line"] as const) {
      expect(firstShape(roundTrip(docOf(shape({ geometry: { preset } })))).geometry.preset).toBe(preset);
    }
  });

  it("round-trips a dashed outline (a:prstDash)", () => {
    const s = shape({ stroke: { color: "#333333", widthPt: 1.5, dash: "dashDot" } });
    const out = firstShape(roundTrip(docOf(s)));
    expect(out.stroke).toEqual({ color: "#333333", widthPt: 1.5, dash: "dashDot" });
    // a solid stroke omits a:prstDash entirely.
    expect(exportedDocXml(docOf(s))).toContain('<a:prstDash val="dashDot"/>');
    expect(exportedDocXml(docOf(shape({ stroke: { color: "#333333", widthPt: 1 } })))).not.toContain("a:prstDash");
  });

  it("round-trips rotation (a:xfrm@rot degrees ↔ 60000ths)", () => {
    const out = firstShape(roundTrip(docOf(shape({ rotation: 20 }))));
    expect(out.rotation).toBe(20);
    expect(exportedDocXml(docOf(shape({ rotation: 20 })))).toContain('<a:xfrm rot="1200000">');
    // no rotation → the a:xfrm carries no rot attribute.
    expect(exportedDocXml(docOf(shape()))).toContain("<a:xfrm>");
  });

  it("round-trips a:avLst adjust guides on a parametric preset", () => {
    const s = shape({ geometry: { preset: "roundRect", adjust: { adj: 18000 } } });
    const out = firstShape(roundTrip(docOf(s)));
    expect(out.geometry).toEqual({ preset: "roundRect", adjust: { adj: 18000 } });
    expect(exportedDocXml(docOf(s))).toContain('<a:gd name="adj" fmla="val 18000"/>');
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

  it("imports a hand-written adjust + rotation + dash (the importer as oracle)", () => {
    const drawing = `<w:p><w:r><w:drawing>
      <wp:inline><wp:extent cx="1905000" cy="1143000"/><wp:docPr id="1" name="Shape"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp><wps:cNvSpPr/><wps:spPr>
            <a:xfrm rot="1200000"><a:off x="0" y="0"/><a:ext cx="1905000" cy="1143000"/></a:xfrm>
            <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 18000"/></a:avLst></a:prstGeom>
            <a:solidFill><a:srgbClr val="d9ead3"/></a:solidFill>
            <a:ln w="19050"><a:solidFill><a:srgbClr val="38761d"/></a:solidFill><a:prstDash val="dash"/></a:ln>
          </wps:spPr><wps:bodyPr/></wps:wsp>
        </a:graphicData></a:graphic>
      </wp:inline>
    </w:drawing></w:r></w:p>`;
    const out = firstShape(runImport(simpleDocx(drawing)).doc);
    expect(out.geometry).toEqual({ preset: "roundRect", adjust: { adj: 18000 } });
    expect(out.rotation).toBe(20);
    expect(out.stroke).toEqual({ color: "#38761d", widthPt: 1.5, dash: "dash" });
  });

  it("round-trips a text box body (wps:txbx paragraphs survive the cycle)", () => {
    const s = shape({ text: { blocks: [txtPara("Drawing text box"), txtPara("second line")] } });
    const out = firstShape(roundTrip(docOf(s)));
    expect(out.text?.blocks.length).toBe(2);
    expect(out.text?.blocks[0]?.runs.map((r) => r.text).join("")).toBe("Drawing text box");
    expect(out.text?.blocks[1]?.runs.map((r) => r.text).join("")).toBe("second line");
  });

  it("persists an EDITED text box body through the cycle (issue #219)", () => {
    // Edit the box the way the caret-in-box editing path does — through the model
    // ops (insert a char, then split into a second paragraph) — then round-trip.
    const s = shape({ text: { blocks: [txtPara("Hi")] } });
    const firstId = s.text!.blocks[0]!.id;
    let doc: Document = docOf(s);
    doc = applyOp(doc, { type: "insertText", at: { blockId: firstId, offset: 2 }, text: "!" }).doc;
    doc = applyOp(doc, { type: "splitParagraph", at: { blockId: firstId, offset: 3 }, newBlockId: "np" }).doc;
    doc = applyOp(doc, { type: "insertText", at: { blockId: "np", offset: 0 }, text: "line 2" }).doc;
    const out = firstShape(roundTrip(doc));
    expect(out.text?.blocks.map((b) => b.runs.map((r) => r.text).join(""))).toEqual(["Hi!", "line 2"]);
  });

  it("emits wps:txbx / w:txbxContent for a text box body", () => {
    const xml = exportedDocXml(docOf(shape({ text: { blocks: [txtPara("Boxed text")] } })));
    expect(xml).toContain("<wps:txbx>");
    expect(xml).toContain("<w:txbxContent>");
    expect(xml).toContain("Boxed text");
    // txbx must precede bodyPr per CT_WordprocessingShape.
    expect(xml.indexOf("<wps:txbx>")).toBeLessThan(xml.indexOf("<wps:bodyPr"));
  });

  it("omits wps:txbx when the shape has no text body", () => {
    const xml = exportedDocXml(docOf(shape()));
    expect(xml).not.toContain("wps:txbx");
  });

  it("imports a hand-written wps:txbx text box (the importer as oracle)", () => {
    const drawing = `<w:p><w:r><w:drawing>
      <wp:inline><wp:extent cx="2743200" cy="914400"/><wp:docPr id="1" name="Shape"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp><wps:cNvSpPr/><wps:spPr>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="fff2cc"/></a:solidFill>
          </wps:spPr>
          <wps:txbx><w:txbxContent>
            <w:p><w:r><w:t>Line one</w:t></w:r></w:p>
            <w:p><w:r><w:t>Line two</w:t></w:r></w:p>
          </w:txbxContent></wps:txbx>
          <wps:bodyPr/></wps:wsp>
        </a:graphicData></a:graphic>
      </wp:inline>
    </w:drawing></w:r></w:p>`;
    const out = firstShape(runImport(simpleDocx(drawing)).doc);
    expect(out.geometry.preset).toBe("rect");
    expect(out.text?.blocks.length).toBe(2);
    expect(out.text?.blocks[0]?.runs.map((r) => r.text).join("")).toBe("Line one");
    expect(out.text?.blocks[1]?.runs.map((r) => r.text).join("")).toBe("Line two");
  });

  // --- Freeform custom geometry (issue #220): a:custGeom path ------------------
  it("round-trips a freeform custom geometry (a:custGeom path with lines + a cubic)", () => {
    const path: ShapePath = {
      segments: [
        { type: "moveTo", x: 0.5, y: 0 },
        { type: "lineTo", x: 1, y: 1 },
        { type: "cubicBezierTo", x1: 0.75, y1: 0.5, x2: 0.25, y2: 0.5, x: 0, y: 1 },
        { type: "close" },
      ],
    };
    const s = shape({ geometry: { preset: "rect", custom: path } });
    const out = firstShape(roundTrip(docOf(s)));
    expect(out.geometry.custom).toEqual(path);
  });

  it("emits a:custGeom (not a:prstGeom) for a custom geometry, and can fill it", () => {
    const path: ShapePath = { segments: [{ type: "moveTo", x: 0, y: 0 }, { type: "lineTo", x: 1, y: 0 }, { type: "lineTo", x: 0.5, y: 1 }, { type: "close" }] };
    const xml = exportedDocXml(docOf(shape({ geometry: { preset: "rect", custom: path }, fill: { color: "#ffe599" } })));
    expect(xml).toContain("<a:custGeom>");
    expect(xml).toContain("<a:pathLst>");
    expect(xml).toContain('<a:path w="100000" h="100000">');
    expect(xml).toContain("<a:moveTo><a:pt");
    expect(xml).not.toContain("<a:cubicBezTo"); // this simple triangle path has no cubic
    expect(xml).not.toContain("<a:prstGeom");
    expect(xml).toContain('<a:srgbClr val="ffe599"/>');
  });

  it("imports a hand-written a:custGeom shape (the importer as oracle)", () => {
    const drawing = `<w:p><w:r><w:drawing>
      <wp:inline><wp:extent cx="1143000" cy="1143000"/><wp:docPr id="1" name="Shape"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp><wps:cNvSpPr/><wps:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="1143000" cy="1143000"/></a:xfrm>
            <a:custGeom><a:pathLst><a:path w="200" h="100">
              <a:moveTo><a:pt x="100" y="0"/></a:moveTo>
              <a:lnTo><a:pt x="200" y="100"/></a:lnTo>
              <a:lnTo><a:pt x="0" y="100"/></a:lnTo>
              <a:close/>
            </a:path></a:pathLst></a:custGeom>
            <a:solidFill><a:srgbClr val="ffe599"/></a:solidFill>
          </wps:spPr><wps:bodyPr/></wps:wsp>
        </a:graphicData></a:graphic>
      </wp:inline>
    </w:drawing></w:r></w:p>`;
    const out = firstShape(runImport(simpleDocx(drawing)).doc);
    expect(out.geometry.custom).toEqual({
      segments: [
        { type: "moveTo", x: 0.5, y: 0 },
        { type: "lineTo", x: 1, y: 1 },
        { type: "lineTo", x: 0, y: 1 },
        { type: "close" },
      ],
    });
  });

  it("warns (custom-path-simplified) when a custom path uses an unmodeled command (a:arcTo)", () => {
    const drawing = `<w:p><w:r><w:drawing>
      <wp:inline><wp:extent cx="1143000" cy="1143000"/><wp:docPr id="1" name="Shape"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp><wps:cNvSpPr/><wps:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="1143000" cy="1143000"/></a:xfrm>
            <a:custGeom><a:pathLst><a:path w="100" h="100">
              <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
              <a:lnTo><a:pt x="100" y="0"/></a:lnTo>
              <a:arcTo wR="50" hR="50" stAng="0" swAng="5400000"/>
              <a:close/>
            </a:path></a:pathLst></a:custGeom>
          </wps:spPr><wps:bodyPr/></wps:wsp>
        </a:graphicData></a:graphic>
      </wp:inline>
    </w:drawing></w:r></w:p>`;
    const result = runImport(simpleDocx(drawing));
    // The arc is dropped but the rest of the path still imports — and it warns.
    expect(result.warnings.map((w) => w.code)).toContain("custom-path-simplified");
    expect(firstShape(result.doc).geometry.custom).toEqual({
      segments: [
        { type: "moveTo", x: 0, y: 0 },
        { type: "lineTo", x: 1, y: 0 },
        { type: "close" },
      ],
    });
  });

  // --- Positioning parity (issue #217): wrap / float anchor / z-order ----------
  it("round-trips square text-wrap (wp:anchor + wp:wrapSquare)", () => {
    const s = shape({ wrap: "square", align: "left" });
    const xml = exportedDocXml(docOf(s));
    expect(xml).toContain("<wp:anchor");
    expect(xml).toContain('<wp:wrapSquare wrapText="bothSides"/>');
    const out = firstShape(roundTrip(docOf(s)));
    expect(out.wrap).toBe("square");
    expect(out.anchor).toBeUndefined(); // exclusive with anchor
  });

  it("round-trips a behind-text float anchor (wrapNone, behindDoc, offsets, z)", () => {
    const s = shape({
      anchor: { behind: true, offsetXPx: 40, offsetYPx: 120, relFromH: "margin", relFromV: "paragraph", z: -1 },
    });
    const xml = exportedDocXml(docOf(s));
    expect(xml).toContain("<wp:anchor");
    expect(xml).toContain("<wp:wrapNone/>");
    expect(xml).toContain('behindDoc="1"');
    const out = firstShape(roundTrip(docOf(s)));
    expect(out.wrap).toBeUndefined();
    expect(out.anchor).toMatchObject({ behind: true, offsetXPx: 40, offsetYPx: 120, relFromH: "margin", relFromV: "paragraph" });
  });

  it("round-trips an in-front-of-text anchor with a positive z-order", () => {
    const s = shape({
      anchor: { behind: false, offsetXPx: 10, offsetYPx: 5, relFromH: "margin", relFromV: "paragraph", z: 4 },
    });
    const out = firstShape(roundTrip(docOf(s)));
    expect(out.anchor?.behind).toBe(false);
    expect(out.anchor?.z).toBe(4);
  });

  // --- Grouped shapes (issue #221): wpg:wgp -----------------------------------
  const groupChild = (patch: Partial<ShapeBlock> & Pick<ShapeBlock, "widthPx" | "heightPx">): ShapeBlock => ({
    kind: "shape", id: `g${n++}`, revision: 0, geometry: { preset: "rect" }, align: "left", ...patch,
  });
  const groupShape = (children: ShapeGroupChild[], patch: Partial<ShapeBlock> = {}): ShapeBlock => ({
    kind: "shape", id: `grp${n++}`, revision: 0, geometry: { preset: "rect" },
    widthPx: 300, heightPx: 100, align: "center",
    group: { childOffsetXPx: 0, childOffsetYPx: 0, childExtentXPx: 300, childExtentYPx: 100, children },
    ...patch,
  });

  it("round-trips a flat group (members' preset/rect/fill/stroke survive wpg:wgp)", () => {
    const s = groupShape([
      { xPx: 0, yPx: 20, shape: groupChild({ geometry: { preset: "roundRect" }, widthPx: 120, heightPx: 56, fill: { color: "#d9ead3" }, stroke: { color: "#38761d", widthPt: 1.5 } }) },
      { xPx: 140, yPx: 40, shape: groupChild({ geometry: { preset: "rightArrow" }, widthPx: 60, heightPx: 20 }) },
      { xPx: 210, yPx: 20, shape: groupChild({ geometry: { preset: "ellipse" }, widthPx: 90, heightPx: 56, fill: { color: "#c9daf8" } }) },
    ]);
    const out = firstShape(roundTrip(docOf(s)));
    expect(out.group).toBeDefined();
    expect(out.widthPx).toBe(300);
    expect(out.heightPx).toBe(100);
    expect(out.group!.childExtentXPx).toBe(300);
    expect(out.group!.childExtentYPx).toBe(100);
    const kids = out.group!.children;
    expect(kids.length).toBe(3);
    expect(kids.map((c) => c.shape.geometry.preset)).toEqual(["roundRect", "rightArrow", "ellipse"]);
    expect(kids[0]).toMatchObject({ xPx: 0, yPx: 20 });
    expect(kids[0]!.shape).toMatchObject({ widthPx: 120, heightPx: 56, fill: { color: "#d9ead3" }, stroke: { color: "#38761d", widthPt: 1.5 } });
    expect(kids[1]).toMatchObject({ xPx: 140, yPx: 40 });
    expect(kids[2]!.shape.fill).toEqual({ color: "#c9daf8" });
  });

  it("emits bare DrawingML wpg:wgp with a:chOff/a:chExt and member wps:wsp (no VML)", () => {
    const s = groupShape([
      { xPx: 10, yPx: 10, shape: groupChild({ geometry: { preset: "rect" }, widthPx: 80, heightPx: 40 }) },
      { xPx: 120, yPx: 10, shape: groupChild({ geometry: { preset: "ellipse" }, widthPx: 80, heightPx: 40 }) },
    ]);
    const xml = exportedDocXml(docOf(s));
    expect(xml).toContain("<wpg:wgp>");
    expect(xml).toContain('uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"');
    expect(xml).toContain("<a:chOff");
    expect(xml).toContain("<a:chExt");
    expect(xml).toContain('xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"');
    // two members, each a wps:wsp inside the group.
    expect((xml.match(/<wps:wsp>/g) ?? []).length).toBe(2);
    expect(xml).not.toContain("mc:AlternateContent");
    expect(xml).not.toContain("<v:");
  });

  it("round-trips a group member's text box body", () => {
    const s = groupShape([
      { xPx: 0, yPx: 0, shape: groupChild({ geometry: { preset: "roundRect" }, widthPx: 130, heightPx: 56, text: { blocks: [txtPara("Boxed member")] } }) },
    ]);
    const out = firstShape(roundTrip(docOf(s)));
    expect(out.group!.children[0]!.shape.text?.blocks[0]?.runs.map((r) => r.text).join("")).toBe("Boxed member");
  });

  it("round-trips a nested group (a group inside a group)", () => {
    const inner = groupShape([
      { xPx: 0, yPx: 0, shape: groupChild({ geometry: { preset: "ellipse" }, widthPx: 40, heightPx: 40, fill: { color: "#ead1dc" } }) },
      { xPx: 50, yPx: 0, shape: groupChild({ geometry: { preset: "triangle" }, widthPx: 40, heightPx: 40 }) },
    ], { widthPx: 100, heightPx: 40 });
    const outer = groupShape([
      { xPx: 0, yPx: 0, shape: groupChild({ geometry: { preset: "rect" }, widthPx: 120, heightPx: 60, stroke: { color: "#333333", widthPt: 1 } }) },
      { xPx: 130, yPx: 10, shape: inner },
    ]);
    const out = firstShape(roundTrip(docOf(outer)));
    expect(out.group!.children.length).toBe(2);
    const nested = out.group!.children[1]!.shape;
    expect(nested.group).toBeDefined();
    expect(nested.group!.children.map((c) => c.shape.geometry.preset)).toEqual(["ellipse", "triangle"]);
    // export emits a nested wpg:grpSp for the inner group.
    expect(exportedDocXml(docOf(outer))).toContain("<wpg:grpSp>");
  });

  it("round-trips a rotated group (a:xfrm@rot on wpg:grpSpPr)", () => {
    const s = groupShape([
      { xPx: 0, yPx: 0, shape: groupChild({ geometry: { preset: "rect" }, widthPx: 120, heightPx: 60, fill: { color: "#d9ead3" } }) },
      { xPx: 140, yPx: 20, shape: groupChild({ geometry: { preset: "ellipse" }, widthPx: 100, heightPx: 60 }) },
    ], { rotation: 15 });
    // export puts the rotation on the group's a:xfrm (60000ths of a degree).
    expect(exportedDocXml(docOf(s))).toContain('<a:xfrm rot="900000">');
    const out = firstShape(roundTrip(docOf(s)));
    expect(out.rotation).toBe(15);
    expect(out.group!.children.length).toBe(2);
  });

  it("round-trips a rotated nested group", () => {
    const inner = groupShape([
      { xPx: 0, yPx: 0, shape: groupChild({ geometry: { preset: "triangle" }, widthPx: 40, heightPx: 40 }) },
    ], { widthPx: 60, heightPx: 40, rotation: 30 });
    const outer = groupShape([
      { xPx: 0, yPx: 0, shape: groupChild({ geometry: { preset: "rect" }, widthPx: 80, heightPx: 60 }) },
      { xPx: 100, yPx: 0, shape: inner },
    ]);
    const out = firstShape(roundTrip(docOf(outer)));
    expect(out.group!.children[1]!.shape.rotation).toBe(30);
  });

  it("skips a group member without an a:xfrm/a:ext (unsized) with a warning", () => {
    const drawing = `<w:p><w:r><w:drawing>
      <wp:inline><wp:extent cx="2743200" cy="914400"/><wp:docPr id="1" name="Group"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
          <wpg:wgp>
            <wpg:cNvGrpSpPr/>
            <wpg:grpSpPr><a:xfrm>
              <a:off x="0" y="0"/><a:ext cx="2743200" cy="914400"/>
              <a:chOff x="0" y="0"/><a:chExt cx="2743200" cy="914400"/>
            </a:xfrm></wpg:grpSpPr>
            <wps:wsp><wps:cNvSpPr/><wps:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="1143000" cy="533400"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </wps:spPr><wps:bodyPr/></wps:wsp>
            <wps:wsp><wps:cNvSpPr/><wps:spPr>
              <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
            </wps:spPr><wps:bodyPr/></wps:wsp>
          </wpg:wgp>
        </a:graphicData></a:graphic>
      </wp:inline>
    </w:drawing></w:r></w:p>`;
    const res = runImport(simpleDocx(drawing));
    const out = firstShape(res.doc);
    // the unsized second member is dropped; only the sized rect survives.
    expect(out.group!.children.length).toBe(1);
    expect(out.group!.children[0]!.shape.geometry.preset).toBe("rect");
    expect(res.warnings.some((w) => w.code === "shape-group-child-unsized")).toBe(true);
  });

  it("imports a hand-written wpg:wgp group (the importer as oracle)", () => {
    const drawing = `<w:p><w:r><w:drawing>
      <wp:inline><wp:extent cx="2743200" cy="914400"/><wp:docPr id="1" name="Group"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
          <wpg:wgp>
            <wpg:cNvGrpSpPr/>
            <wpg:grpSpPr>
              <a:xfrm>
                <a:off x="0" y="0"/><a:ext cx="2743200" cy="914400"/>
                <a:chOff x="0" y="0"/><a:chExt cx="2743200" cy="914400"/>
              </a:xfrm>
            </wpg:grpSpPr>
            <wps:wsp><wps:cNvSpPr/><wps:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="1143000" cy="533400"/></a:xfrm>
              <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
              <a:solidFill><a:srgbClr val="d9ead3"/></a:solidFill>
            </wps:spPr><wps:bodyPr/></wps:wsp>
            <wps:wsp><wps:cNvSpPr/><wps:spPr>
              <a:xfrm><a:off x="1600200" y="0"/><a:ext cx="1143000" cy="533400"/></a:xfrm>
              <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
              <a:solidFill><a:srgbClr val="c9daf8"/></a:solidFill>
            </wps:spPr><wps:bodyPr/></wps:wsp>
          </wpg:wgp>
        </a:graphicData></a:graphic>
      </wp:inline>
    </w:drawing></w:r></w:p>`;
    const out = firstShape(runImport(simpleDocx(drawing)).doc);
    expect(out.group).toBeDefined();
    expect(out.widthPx).toBe(288); // 2743200 EMU / 9525
    expect(out.group!.children.map((c) => c.shape.geometry.preset)).toEqual(["roundRect", "ellipse"]);
    expect(out.group!.children[0]!.xPx).toBe(0);
    expect(out.group!.children[1]!.xPx).toBe(168); // 1600200 / 9525
    expect(out.group!.children[0]!.shape.fill).toEqual({ color: "#d9ead3" });
  });

  it("maps a genuinely unsupported preset to a rectangle (+ warning), dropping its adjust", () => {
    const drawing = `<w:p><w:r><w:drawing>
      <wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Shape"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp><wps:cNvSpPr/><wps:spPr>
            <a:prstGeom prst="star5"><a:avLst><a:gd name="adj" fmla="val 19000"/></a:avLst></a:prstGeom>
            <a:solidFill><a:srgbClr val="d9e2f0"/></a:solidFill>
          </wps:spPr><wps:bodyPr/></wps:wsp>
        </a:graphicData></a:graphic>
      </wp:inline>
    </w:drawing></w:r></w:p>`;
    const res = runImport(simpleDocx(drawing));
    const out = firstShape(res.doc);
    expect(out.geometry).toEqual({ preset: "rect" });
  });
});
