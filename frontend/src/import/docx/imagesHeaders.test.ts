// Milestone 3: embedded/linked images and header/footer block stories.

import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import type { ImageBlock, Paragraph, TableBlock } from "@cw/shared";
import { runImport } from "./pipeline";
import { writeDocx } from "../../export/docx/writeDocx";
import {
  CONTENT_TYPES_XML,
  documentXml,
  drawingXml,
  footerPartXml,
  headerPartXml,
  makeDocx,
  PNG_1PX,
  REL_TYPES,
  relsXml,
} from "./fixture";

const para = (b: unknown): Paragraph => {
  expect((b as Paragraph).kind).toBe("paragraph");
  return b as Paragraph;
};
const image = (b: unknown): ImageBlock => {
  expect((b as ImageBlock).kind).toBe("image");
  return b as ImageBlock;
};
const text = (p: Paragraph): string => p.runs.map((r) => r.text).join("");
const codes = (r: ReturnType<typeof runImport>): string[] => r.warnings.map((w) => w.code);

/** document.xml + rels + a PNG media part. */
function docxWithImage(body: string, extra: Record<string, string | Uint8Array> = {}): Uint8Array {
  return makeDocx({
    "[Content_Types].xml": CONTENT_TYPES_XML,
    "word/document.xml": documentXml(body),
    "word/_rels/document.xml.rels": relsXml([{ id: "rId10", type: REL_TYPES.image, target: "media/img.png" }]),
    "word/media/img.png": PNG_1PX,
    ...extra,
  });
}

describe("images — embedded", () => {
  it("surfaces an inline drawing as an ImageBlock, splitting the paragraph", () => {
    const r = runImport(
      docxWithImage(
        `<w:p><w:r><w:t>before</w:t>${drawingXml("rId10", 914400, 457200)}<w:t>after</w:t></w:r></w:p>`,
      ),
    );
    expect(r.doc.blocks).toHaveLength(3);
    expect(text(para(r.doc.blocks[0]))).toBe("before");
    const img = image(r.doc.blocks[1]);
    expect(img.src).toMatch(/^blob:/);
    expect(img.widthPx).toBe(96); // 1 inch
    expect(img.heightPx).toBe(48);
    expect(text(para(r.doc.blocks[2]))).toBe("after");
    expect(r.mediaUrls).toHaveLength(1);
    expect(r.mediaUrls[0]).toBe(img.src);
  });

  it("maps an image-only paragraph to a single ImageBlock (no stray empties)", () => {
    const r = runImport(docxWithImage(`<w:p><w:r>${drawingXml("rId10", 914400, 914400)}</w:r></w:p>`));
    expect(r.doc.blocks).toHaveLength(1);
    image(r.doc.blocks[0]);
  });

  it("inherits block alignment from the paragraph", () => {
    const r = runImport(
      docxWithImage(
        `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${drawingXml("rId10", 914400, 914400)}</w:r></w:p>`,
      ),
    );
    expect(image(r.doc.blocks[0]).align).toBe("center");
  });

  it("demotes overlap-wrapped anchored images into the text flow with a warning", () => {
    const r = runImport(docxWithImage(`<w:p><w:r>${drawingXml("rId10", 914400, 914400, true)}</w:r></w:p>`));
    const img = image(r.doc.blocks[0]);
    expect(img.wrap).toBeUndefined(); // block flow
    expect(codes(r)).toContain("images-anchored");
  });

  it("maps square-wrapped anchors to floating images (wrap: square)", () => {
    const anchor =
      `<w:drawing><wp:anchor><wp:extent cx="914400" cy="457200"/>` +
      `<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>` +
      `<wp:wrapSquare wrapText="bothSides"/>` +
      `<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>` +
      `</wp:anchor></w:drawing>`;
    const r = runImport(docxWithImage(`<w:p><w:r>${anchor}<w:t>wrapped text</w:t></w:r></w:p>`));
    const img = image(r.doc.blocks[0]);
    expect(img.wrap).toBe("square");
    expect(img.align).toBe("right");
    expect(codes(r)).not.toContain("images-anchored"); // faithful float, no warning
  });

  it("reuses one blob URL when the same media part is referenced twice", () => {
    const r = runImport(
      docxWithImage(
        `<w:p><w:r>${drawingXml("rId10", 914400, 914400)}${drawingXml("rId10", 457200, 457200)}</w:r></w:p>`,
      ),
    );
    expect(r.doc.blocks).toHaveLength(2);
    expect(image(r.doc.blocks[0]).src).toBe(image(r.doc.blocks[1]).src);
    expect(r.mediaUrls).toHaveLength(1);
  });

  it("unwraps mc:AlternateContent (how Word actually ships drawings)", () => {
    const body =
      `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wpg">${drawingXml("rId10", 914400, 457200)}</mc:Choice>` +
      `<mc:Fallback><w:pict><v:shape style="width:72pt;height:36pt"><v:imagedata r:id="rId10"/></v:shape></w:pict></mc:Fallback>` +
      `</mc:AlternateContent></w:r></w:p>`;
    const r = runImport(docxWithImage(body));
    expect(r.doc.blocks).toHaveLength(1); // Choice used, Fallback NOT duplicated
    const img = image(r.doc.blocks[0]);
    expect(img.widthPx).toBe(96);
    expect(r.mediaUrls).toHaveLength(1);
  });

  it("imports legacy VML pictures (w:pict) with pt-based sizing", () => {
    const body =
      `<w:p><w:r><w:pict><v:shape style="width:72pt;height:36pt">` +
      `<v:imagedata r:id="rId10"/></v:shape></w:pict></w:r></w:p>`;
    const r = runImport(docxWithImage(body));
    const img = image(r.doc.blocks[0]);
    expect(img.widthPx).toBe(96); // 72pt = 1in
    expect(img.heightPx).toBe(48);
  });

  it("imports images inside table cells (cells are full block stories)", () => {
    const body =
      `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>x</w:t>${drawingXml("rId10", 914400, 457200)}</w:r></w:p></w:tc></w:tr></w:tbl>`;
    const r = runImport(docxWithImage(body));
    const tbl = r.doc.blocks[0] as TableBlock;
    expect(tbl.kind).toBe("table");
    const cellBlocks = tbl.rows[0]!.cells[0]!.blocks;
    expect(text(para(cellBlocks[0]))).toBe("x");
    const img = image(cellBlocks[1]);
    expect(img.src).toMatch(/^blob:/);
    expect(img.widthPx).toBe(96);
    expect(r.mediaUrls).toContain(img.src);
  });

  it("skips unsupported formats (metafiles) with a warning", () => {
    const bytes = makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(`<w:p><w:r>${drawingXml("rId10", 914400, 914400)}</w:r></w:p>`),
      "word/_rels/document.xml.rels": relsXml([{ id: "rId10", type: REL_TYPES.image, target: "media/img.emf" }]),
      "word/media/img.emf": PNG_1PX,
    });
    const r = runImport(bytes);
    expect(r.doc.blocks.every((b) => b.kind !== "image")).toBe(true);
    expect(codes(r)).toContain("media-unsupported");
    expect(r.mediaUrls).toHaveLength(0);
  });

  it("warns when a drawing has no image reference", () => {
    const r = runImport(docxWithImage(`<w:p><w:r><w:drawing><wp:inline/></w:drawing><w:t>t</w:t></w:r></w:p>`));
    expect(text(para(r.doc.blocks[0]))).toBe("t");
    expect(codes(r)).toContain("images-skipped");
  });
});

describe("images — external (TargetMode=External)", () => {
  it("passes http(s) linked images through by URL (not a blob, not revocable)", () => {
    const bytes = makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(`<w:p><w:r>${drawingXml("rId10", 914400, 914400)}</w:r></w:p>`),
      "word/_rels/document.xml.rels": relsXml([
        { id: "rId10", type: REL_TYPES.image, target: "https://example.com/pic.png", external: true },
      ]),
    });
    const r = runImport(bytes);
    expect(image(r.doc.blocks[0]).src).toBe("https://example.com/pic.png");
    expect(r.mediaUrls).toHaveLength(0); // nothing to revoke
    expect(codes(r)).toContain("media-external");
  });

  it("skips linked images pointing at local files", () => {
    const bytes = makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(`<w:p><w:r>${drawingXml("rId10", 914400, 914400)}<w:t>t</w:t></w:r></w:p>`),
      "word/_rels/document.xml.rels": relsXml([
        { id: "rId10", type: REL_TYPES.image, target: "file:///C:/photos/pic.png", external: true },
      ]),
    });
    const r = runImport(bytes);
    expect(r.doc.blocks.every((b) => b.kind !== "image")).toBe(true);
    expect(codes(r)).toContain("media-external-unreachable");
  });
});

describe("headers and footers", () => {
  const SECT = (refs: string) =>
    `<w:sectPr>${refs}<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;

  function docxWithHeaderFooter(opts: {
    body?: string;
    headerXml?: string;
    footerXml?: string;
    headerRefType?: string;
    extraRefs?: string;
    extra?: Record<string, string | Uint8Array>;
  }): Uint8Array {
    const refs =
      (opts.headerXml !== undefined
        ? `<w:headerReference w:type="${opts.headerRefType ?? "default"}" r:id="rIdH"/>`
        : "") +
      (opts.footerXml !== undefined ? `<w:footerReference w:type="default" r:id="rIdF"/>` : "") +
      (opts.extraRefs ?? "");
    const rels = [
      ...(opts.headerXml !== undefined ? [{ id: "rIdH", type: REL_TYPES.header, target: "header1.xml" }] : []),
      ...(opts.footerXml !== undefined ? [{ id: "rIdF", type: REL_TYPES.footer, target: "footer1.xml" }] : []),
    ];
    return makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(`${opts.body ?? "<w:p><w:r><w:t>body</w:t></w:r></w:p>"}${SECT(refs)}`),
      "word/_rels/document.xml.rels": relsXml(rels),
      ...(opts.headerXml !== undefined ? { "word/header1.xml": opts.headerXml } : {}),
      ...(opts.footerXml !== undefined ? { "word/footer1.xml": opts.footerXml } : {}),
      ...(opts.extra ?? {}),
    });
  }

  it("imports and re-exports a bookmark inside a header", () => {
    const r = runImport(
      docxWithHeaderFooter({
        headerXml: headerPartXml(
          `<w:p><w:bookmarkStart w:id="7" w:name="hdrMark"/><w:r><w:t>top</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>`,
        ),
      }),
    );
    // imported: the bookmark resolves to the header band's paragraph
    const range = r.doc.bookmarks?.["hdrMark"];
    expect(range).toBeDefined();
    expect(range!.start.blockId).toBe(para(r.doc.section.header![0]).id);
    // re-exported: the marker lands in the HEADER part, not just the body
    const parts = unzipSync(writeDocx(r.doc).bytes);
    const headerPart = Object.keys(parts).find((k) => /word\/header\d+\.xml$/.test(k))!;
    expect(strFromU8(parts[headerPart]!)).toContain('w:name="hdrMark"');
  });

  it("imports the default header as a block story", () => {
    const r = runImport(
      docxWithHeaderFooter({
        headerXml: headerPartXml(`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Confidential</w:t></w:r></w:p>`),
      }),
    );
    expect(r.doc.section.header).toBeDefined();
    const h = para(r.doc.section.header![0]);
    expect(text(h)).toBe("Confidential");
    expect(h.runs[0]!.style.bold).toBe(true);
  });

  it("turns complex PAGE/NUMPAGES fields into live {page}/{pages} tokens", () => {
    const footer = footerPartXml(
      `<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>` +
        `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r>` + // stale cached result
        `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
        `<w:r><w:t xml:space="preserve"> of </w:t></w:r>` +
        `<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>99</w:t></w:r></w:fldSimple></w:p>`,
    );
    const r = runImport(docxWithHeaderFooter({ footerXml: footer }));
    expect(text(para(r.doc.section.footer![0]))).toBe("Page {page} of {pages}");
  });

  it("captures a BODY PAGE field as a {page} field object (fixes the old cached-text behavior)", () => {
    const body =
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const r = runImport(docxWithHeaderFooter({ body, headerXml: headerPartXml(`<w:p/>`) }));
    const p = para(r.doc.blocks[0]);
    expect(text(p)).toBe("{page}"); // a live token, not the stale cached "7"
    const fid = p.runs.find((rr) => rr.text === "{page}")?.style.fieldId;
    expect(fid).toBeDefined();
    expect(r.doc.fields?.[fid!]?.spec).toEqual({ type: "PAGE" });
  });

  it("keeps a TOC entry's PAGEREF cached number as text (must NOT become a token)", () => {
    const body =
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGEREF _Toc1 \\h </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const r = runImport(docxWithHeaderFooter({ body, headerXml: headerPartXml(`<w:p/>`) }));
    expect(text(para(r.doc.blocks[0]))).toBe("7"); // PAGEREF is not a tokenized field
  });

  it("resolves header images through the header part's own rels", () => {
    const r = runImport(
      docxWithHeaderFooter({
        headerXml: headerPartXml(`<w:p><w:r>${drawingXml("rId1", 914400, 457200)}</w:r></w:p>`),
        extra: {
          "word/_rels/header1.xml.rels": relsXml([{ id: "rId1", type: REL_TYPES.image, target: "media/logo.png" }]),
          "word/media/logo.png": PNG_1PX,
        },
      }),
    );
    const img = image(r.doc.section.header![0]);
    expect(img.src).toMatch(/^blob:/);
    expect(r.mediaUrls).toContain(img.src);
  });

  it("maps the default header and ignores a first variant without w:titlePg", () => {
    const r = runImport(
      docxWithHeaderFooter({
        headerXml: headerPartXml(`<w:p><w:r><w:t>default</w:t></w:r></w:p>`),
        extraRefs: `<w:headerReference w:type="first" r:id="rIdH"/>`,
      }),
    );
    expect(text(para(r.doc.section.header![0]))).toBe("default");
    expect(r.doc.section.headerFirst).toBeUndefined(); // no titlePg → first not used
    expect(codes(r)).not.toContain("header-variants");
  });

  it("warns when a referenced header part is missing", () => {
    const bytes = makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(
        `<w:p/>${SECT(`<w:headerReference w:type="default" r:id="rIdH"/>`)}`,
      ),
      "word/_rels/document.xml.rels": relsXml([{ id: "rIdH", type: REL_TYPES.header, target: "header1.xml" }]),
    });
    const r = runImport(bytes);
    expect(r.doc.section.header).toBeUndefined();
    expect(codes(r)).toContain("header-missing");
  });

  it("omits empty header stories", () => {
    const r = runImport(docxWithHeaderFooter({ headerXml: headerPartXml(``) }));
    expect(r.doc.section.header).toBeUndefined();
  });
});
