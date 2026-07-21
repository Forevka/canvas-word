// Issue #218 (Part 5 of #206) — legacy VML shape import. Pre-DrawingML (Word
// 2003-era) documents draw shapes with `w:pict` → `v:rect`/`v:oval`/`v:line`/
// `v:shape` + `v:textbox`, not `w:drawing` → `wps:wsp`. Those used to be dropped;
// now they import into the SAME `kind:"shape"` model the DrawingML path produces.
//
// Import is read-only: a VML-sourced shape RE-EXPORTS as modern DrawingML `wps`
// (the model has no VML writer). These tests assert both the direct import (the
// importer as oracle, mirroring the wps:wsp round-trip suite) and the full
// VML → model → DrawingML → model cycle.
import { describe, expect, it } from "vitest";
import type { Document, ShapeBlock } from "@cw/shared";
import { runImport } from "./pipeline";
import { simpleDocx } from "./fixture";
import { writeDocx } from "../../export/docx/writeDocx";
import { strFromU8, unzipSync } from "fflate";

/** A w:pict body (namespaces are matched by qualified name, like Word output). */
const pictDocx = (vml: string): Uint8Array => simpleDocx(`<w:p><w:r><w:pict>${vml}</w:pict></w:r></w:p>`);
const importPict = (vml: string): ReturnType<typeof runImport> => runImport(pictDocx(vml));
const firstShape = (d: Document): ShapeBlock => d.blocks.find((b): b is ShapeBlock => b.kind === "shape")!;
const shapeOf = (vml: string): ShapeBlock => firstShape(importPict(vml).doc);
const codes = (r: ReturnType<typeof runImport>): string[] => r.warnings.map((w) => w.code);
const reExportXml = (d: Document): string => strFromU8(unzipSync(writeDocx(d).bytes)["word/document.xml"]!);

describe("legacy VML shape import (issue #218)", () => {
  it("imports a filled v:rect (preset / pt-size / fill / stroke)", () => {
    const s = shapeOf(
      `<v:rect style="width:150pt;height:75pt" fillcolor="#4472c4" strokecolor="#2f5496" strokeweight="1.5pt"/>`,
    );
    expect(s.geometry.preset).toBe("rect");
    expect(s.widthPx).toBe(200); // 150pt = 200px @96dpi
    expect(s.heightPx).toBe(100);
    expect(s.fill).toEqual({ color: "#4472c4" });
    expect(s.stroke).toEqual({ color: "#2f5496", widthPt: 1.5 });
  });

  it("maps v:oval → ellipse and v:roundrect → roundRect", () => {
    expect(shapeOf(`<v:oval style="width:80pt;height:80pt"/>`).geometry.preset).toBe("ellipse");
    expect(shapeOf(`<v:roundrect style="width:80pt;height:40pt"/>`).geometry.preset).toBe("roundRect");
  });

  it("imports a v:line, deriving its box from from/to endpoints", () => {
    const s = shapeOf(`<v:line from="0,0" to="150pt,0" strokecolor="#c00000" strokeweight="2pt"/>`);
    expect(s.geometry.preset).toBe("line");
    expect(s.widthPx).toBe(200); // |150pt - 0|
    expect(s.stroke).toEqual({ color: "#c00000", widthPt: 2 });
    // Degenerate axis (horizontal line) keeps a small selectable height, not 0.
    expect(s.heightPx).toBeGreaterThan(0);
  });

  it("synthesizes the VML default stroke for a bare line (black, 0.75pt)", () => {
    expect(shapeOf(`<v:line from="0,0" to="100pt,50pt"/>`).stroke).toEqual({ color: "#000000", widthPt: 0.75 });
  });

  it("honours filled=\"f\" / stroked=\"f\" as explicit none", () => {
    const s = shapeOf(`<v:rect style="width:60pt;height:60pt" filled="f" stroked="f"/>`);
    expect(s.fill).toEqual({ none: true });
    expect(s.stroke).toEqual({ none: true });
  });

  it("recognizes every VML boolean-off spelling (f/false/0/no/off)", () => {
    for (const off of ["f", "false", "0", "no", "off", "OFF", "False"]) {
      const s = shapeOf(`<v:rect style="width:20pt;height:20pt" filled="${off}" stroked="${off}"/>`);
      expect(s.fill).toEqual({ none: true });
      expect(s.stroke).toEqual({ none: true });
    }
    // A truthy value (or absence) does NOT clear the fill/stroke.
    expect(shapeOf(`<v:rect style="width:20pt;height:20pt" filled="t" fillcolor="#4472c4"/>`).fill).toEqual({ color: "#4472c4" });
  });

  it("synthesizes the VML default border for a bare shape (black 0.75pt)", () => {
    // VML shapes are stroked by default — a bare v:rect/v:oval still shows an
    // outline, so the importer materializes the default rather than dropping it.
    expect(shapeOf(`<v:rect style="width:60pt;height:60pt"/>`).stroke).toEqual({ color: "#000000", widthPt: 0.75 });
    expect(shapeOf(`<v:oval style="width:60pt;height:60pt"/>`).stroke).toEqual({ color: "#000000", widthPt: 0.75 });
  });

  it("warns and imports only the first shape of a VML group / multi-shape pict", () => {
    const r = importPict(
      `<v:group style="width:200pt;height:100pt" coordsize="21600,21600">` +
      `<v:rect style="width:80pt;height:80pt" fillcolor="#4472c4"/>` +
      `<v:oval style="width:80pt;height:80pt" fillcolor="#ed7d31"/>` +
      `</v:group>`,
    );
    expect(codes(r)).toContain("vml-group-flattened");
    const shapes = r.doc.blocks.filter((b) => b.kind === "shape");
    expect(shapes).toHaveLength(1); // only the first shape survives the flatten
    expect(firstShape(r.doc).geometry.preset).toBe("rect");
    expect(firstShape(r.doc).fill).toEqual({ color: "#4472c4" });
  });

  it("does not warn about a group for a lone shape (even a text box)", () => {
    const r = importPict(
      `<v:shape type="#_x0000_t202" style="width:120pt;height:60pt">` +
      `<v:textbox><w:txbxContent><w:p><w:r><w:t>solo</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>`,
    );
    expect(codes(r)).not.toContain("vml-group-flattened");
  });

  it("accepts a named fill colour and Word's '#hex [theme-idx]' form", () => {
    expect(shapeOf(`<v:rect style="width:20pt;height:20pt" fillcolor="red"/>`).fill).toEqual({ color: "#ff0000" });
    expect(shapeOf(`<v:rect style="width:20pt;height:20pt" fillcolor="#4472c4 [3204]"/>`).fill).toEqual({ color: "#4472c4" });
  });

  it("resolves a v:shape's preset from its shapetype @type _tNNN suffix", () => {
    // _x0000_t4 = MSO diamond; a bare v:shape references the type by id.
    expect(shapeOf(`<v:shape type="#_x0000_t4" style="width:60pt;height:60pt"/>`).geometry.preset).toBe("diamond");
    // Unknown / custom-geometry v:shape degrades to a rectangle box (still placed).
    expect(shapeOf(`<v:shape type="#_x0000_t0" style="width:60pt;height:60pt" o:spt="0"/>`).geometry.preset).toBe("rect");
  });

  it("imports a v:textbox body as the shape's text (a paragraph flow)", () => {
    const s = shapeOf(
      `<v:shapetype id="_x0000_t202" o:spt="202" coordsize="21600,21600"/>` +
      `<v:shape type="#_x0000_t202" style="width:200pt;height:100pt" fillcolor="#fff2cc" strokecolor="#bf9000">` +
      `<v:textbox><w:txbxContent>` +
      `<w:p><w:r><w:t>Legacy text box</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>second line</w:t></w:r></w:p>` +
      `</w:txbxContent></v:textbox></v:shape>`,
    );
    expect(s.geometry.preset).toBe("rect"); // spt 202 = text box → a box
    expect(s.fill).toEqual({ color: "#fff2cc" });
    expect(s.text?.blocks.length).toBe(2);
    expect(s.text?.blocks[0]?.runs.map((r) => r.text).join("")).toBe("Legacy text box");
    expect(s.text?.blocks[1]?.runs.map((r) => r.text).join("")).toBe("second line");
  });

  it("no longer warns for a recognized VML shape (warning relaxed)", () => {
    const r = importPict(`<v:rect style="width:60pt;height:60pt" fillcolor="#4472c4"/>`);
    expect(r.warnings.map((w) => w.code)).not.toContain("pict-skipped");
    expect(r.warnings.map((w) => w.code)).not.toContain("images-skipped");
    expect(firstShape(r.doc)).toBeDefined();
  });

  it("still warns for truly-unknown pict content (no picture, no shape)", () => {
    const r = importPict(`<w:binData w:name="foo">AA==</w:binData>`);
    expect(codes(r)).toContain("pict-skipped");
    expect(r.doc.blocks.some((b) => b.kind === "shape")).toBe(false);
  });

  it("keeps the legacy VML picture path working (v:imagedata is unchanged)", () => {
    // No image rel present, so it drops — but as a picture, not a shape: proves
    // the imagedata branch still runs first and doesn't fall through to a shape.
    const r = importPict(`<v:shape style="width:72pt;height:36pt"><v:imagedata r:id="rIdMissing"/></v:shape>`);
    expect(r.doc.blocks.some((b) => b.kind === "shape")).toBe(false);
  });

  it("full round-trips VML → model → DrawingML wps → model (read-only import)", () => {
    const vml =
      `<v:shape type="#_x0000_t202" style="width:180pt;height:90pt" fillcolor="#d9ead3" strokecolor="#38761d" strokeweight="1pt">` +
      `<v:textbox><w:txbxContent><w:p><w:r><w:t>Boxed</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>`;
    const imported = importPict(vml).doc;
    // Re-export emits modern DrawingML wps — NOT VML — per the intentional
    // VML→DrawingML normalization (documented in OOXML_COVERAGE.md).
    const xml = reExportXml(imported);
    expect(xml).toContain("<wps:wsp>");
    expect(xml).not.toContain("<v:");
    expect(xml).not.toContain("w:pict");
    // …and the re-imported document matches the first import (lossless model).
    const round = firstShape(runImport(writeDocx(imported).bytes).doc);
    const first = firstShape(imported);
    expect(round.geometry.preset).toBe(first.geometry.preset);
    expect(round.widthPx).toBe(first.widthPx);
    expect(round.fill).toEqual(first.fill);
    expect(round.stroke).toEqual(first.stroke);
    expect(round.text?.blocks[0]?.runs.map((r) => r.text).join("")).toBe("Boxed");
  });
});
