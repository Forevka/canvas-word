// Multi-section: header/footer variants (first/even), per-section geometry &
// columns & page-number restart, and column breaks.

import { describe, expect, it } from "vitest";
import type { Paragraph } from "@cw/shared";
import { runImport } from "./pipeline";
import {
  CONTENT_TYPES_XML,
  documentXml,
  footerPartXml,
  headerPartXml,
  makeDocx,
  REL_TYPES,
  relsXml,
  simpleDocx,
} from "./fixture";

const para = (b: unknown): Paragraph => {
  expect((b as Paragraph).kind).toBe("paragraph");
  return b as Paragraph;
};
const text = (p: Paragraph): string => p.runs.map((r) => r.text).join("");

describe("column breaks", () => {
  it("maps w:br type=column to columnBreakBefore on the follower", () => {
    const r = runImport(simpleDocx(`<w:p><w:r><w:t>col1</w:t><w:br w:type="column"/><w:t>col2</w:t></w:r></w:p>`));
    expect(r.doc.blocks).toHaveLength(2);
    expect(text(para(r.doc.blocks[1]))).toBe("col2");
    expect(para(r.doc.blocks[1]).style.columnBreakBefore).toBe(true);
    expect(para(r.doc.blocks[0]).style.columnBreakBefore).toBeUndefined();
  });
});

describe("per-section columns & page numbering", () => {
  const withSect = (sect: string) =>
    runImport(
      simpleDocx(
        `<w:p><w:pPr><w:sectPr>${sect}<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p>` +
          `<w:p><w:r><w:t>next</w:t></w:r></w:p>`,
      ),
    );

  it("applies a mid-doc column change via a sectionBreak patch", () => {
    const r = withSect(`<w:cols w:num="2" w:space="720"/>`);
    const sb = para(r.doc.blocks[0]).style.sectionBreak;
    expect(sb?.props.columns).toEqual({ count: 2, gapPx: 48 });
  });

  it("applies a page-number restart via a sectionBreak patch", () => {
    const r = withSect(`<w:pgNumType w:start="1"/>`);
    expect(para(r.doc.blocks[0]).style.sectionBreak?.props.pageNumberStart).toBe(1);
  });
});

// --- header/footer variants -------------------------------------------------

const REFS = {
  default: (kind: "header" | "footer", id: string) => `<w:${kind}Reference w:type="default" r:id="${id}"/>`,
  first: (kind: "header" | "footer", id: string) => `<w:${kind}Reference w:type="first" r:id="${id}"/>`,
  even: (kind: "header" | "footer", id: string) => `<w:${kind}Reference w:type="even" r:id="${id}"/>`,
};

/** Build a docx whose body sectPr references several header parts, with optional
 *  w:titlePg and a settings.xml that may enable even/odd. */
function docxVariants(opts: {
  refs: string;
  titlePg?: boolean;
  evenAndOdd?: boolean;
  parts: Record<string, string>;
  partRels: { id: string; target: string }[];
}): Uint8Array {
  const sect = `<w:sectPr>${opts.refs}${opts.titlePg ? "<w:titlePg/>" : ""}<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;
  const docRels = [
    { id: "rS", type: REL_TYPES.settings, target: "settings.xml" },
    ...opts.partRels.map((p) => ({ id: p.id, type: p.target.includes("header") ? REL_TYPES.header : REL_TYPES.footer, target: p.target })),
  ];
  return makeDocx({
    "[Content_Types].xml": CONTENT_TYPES_XML,
    "word/document.xml": documentXml(`<w:p><w:r><w:t>body</w:t></w:r></w:p>${sect}`),
    "word/_rels/document.xml.rels": relsXml(docRels),
    "word/settings.xml": `<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${opts.evenAndOdd ? "<w:evenAndOddHeadersAndFooters/>" : ""}</w:settings>`,
    ...opts.parts,
  });
}

describe("header/footer variants", () => {
  it("maps the first-page header only when w:titlePg is set", () => {
    const opts = {
      refs: REFS.default("header", "rH0") + REFS.first("header", "rH1"),
      parts: {
        "word/header0.xml": headerPartXml(`<w:p><w:r><w:t>main head</w:t></w:r></w:p>`),
        "word/header1.xml": headerPartXml(`<w:p><w:r><w:t>title head</w:t></w:r></w:p>`),
      },
      partRels: [
        { id: "rH0", target: "header0.xml" },
        { id: "rH1", target: "header1.xml" },
      ],
    };
    const withTitle = runImport(docxVariants({ ...opts, titlePg: true }));
    expect(text(para(withTitle.doc.section.header![0]))).toBe("main head");
    expect(text(para(withTitle.doc.section.headerFirst![0]))).toBe("title head");

    const noTitle = runImport(docxVariants(opts));
    expect(noTitle.doc.section.headerFirst).toBeUndefined();
  });

  it("maps the even-page footer only when settings enable even/odd", () => {
    const opts = {
      refs: REFS.default("footer", "rF0") + REFS.even("footer", "rF1"),
      parts: {
        "word/footer0.xml": footerPartXml(`<w:p><w:r><w:t>odd foot</w:t></w:r></w:p>`),
        "word/footer1.xml": footerPartXml(`<w:p><w:r><w:t>even foot</w:t></w:r></w:p>`),
      },
      partRels: [
        { id: "rF0", target: "footer0.xml" },
        { id: "rF1", target: "footer1.xml" },
      ],
    };
    const on = runImport(docxVariants({ ...opts, evenAndOdd: true }));
    expect(text(para(on.doc.section.footer![0]))).toBe("odd foot");
    expect(text(para(on.doc.section.footerEven![0]))).toBe("even foot");

    const off = runImport(docxVariants(opts));
    expect(off.doc.section.footerEven).toBeUndefined();
  });
});
