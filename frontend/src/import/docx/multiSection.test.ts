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

  it("carries column separator + per-column widths into the patch", () => {
    const r = withSect(
      `<w:cols w:num="2" w:sep="1" w:equalWidth="0"><w:col w:w="3000" w:space="240"/><w:col w:w="4000"/></w:cols>`,
    );
    const cols = para(r.doc.blocks[0]).style.sectionBreak?.props.columns;
    expect(cols?.count).toBe(2);
    expect(cols?.sep).toBe(true);
    expect(cols?.cols).toHaveLength(2);
    expect(cols?.cols?.[0]?.widthPx).toBeCloseTo(200, 0); // 3000 twips
  });

  it("treats a mid-doc page border as a distinct section and carries it", () => {
    const r = withSect(
      `<w:pgBorders w:offsetFrom="page"><w:top w:val="single" w:sz="12" w:space="24" w:color="FF0000"/></w:pgBorders>`,
    );
    const sb = para(r.doc.blocks[0]).style.sectionBreak;
    expect(sb?.type).toBe("nextPage");
    expect(sb?.props.pageBorders?.offsetFrom).toBe("page");
    expect(sb?.props.pageBorders?.top?.color.toLowerCase()).toBe("#ff0000");
    expect(sb?.props.pageBorders?.top?.widthPx).toBeCloseTo(2, 1); // 12 eighth-pt
  });
});

// --- non-inheriting own-properties don't bleed across a dropped break --------

describe("line numbering / page-number restart don't bleed backward", () => {
  // Line numbering and page-number restart are per-section OWN properties that
  // never inherit. A geometry-preserving Next Page break is normally dropped and
  // its content flowed into the FOLLOWING section — but if that section declares
  // one of these properties, dropping the boundary would let it bleed backward
  // onto the earlier content (regression: a round-tripped doc showed line numbers
  // on every page from page 1, not just inside its line-numbered section).
  const numberedFollows = (firstSect: string, secondSect: string) =>
    runImport(
      simpleDocx(
        `<w:p><w:r><w:t>intro</w:t></w:r></w:p>` +
          `<w:p><w:pPr><w:sectPr>${firstSect}</w:sectPr></w:pPr></w:p>` +
          `<w:p><w:r><w:t>numbered</w:t></w:r></w:p>` +
          `<w:p><w:pPr><w:sectPr>${secondSect}</w:sectPr></w:pPr></w:p>` +
          `<w:p><w:r><w:t>tail</w:t></w:r></w:p>` +
          `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`,
      ),
    );
  const breaksOf = (r: ReturnType<typeof runImport>): Paragraph[] =>
    r.doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph" && !!b.style.sectionBreak);

  it("keeps a geometry-preserving break before a line-numbered section", () => {
    const r = numberedFollows(
      `<w:type w:val="nextPage"/>`,
      `<w:lnNumType w:countBy="1" w:restart="newPage"/><w:type w:val="nextPage"/>`,
    );
    const breaks = breaksOf(r);
    // The first break is NOT distinct on its own (same geometry, plain Next Page),
    // but must survive so the intro keeps its own UNnumbered section.
    expect(breaks).toHaveLength(2);
    expect(breaks[0]!.style.sectionBreak!.props.lineNumbering).toBeUndefined();
    expect(breaks[1]!.style.sectionBreak!.props.lineNumbering).toEqual({ countBy: 1, restart: "newPage" });
  });

  it("keeps a geometry-preserving break before a page-number restart", () => {
    const r = numberedFollows(`<w:type w:val="nextPage"/>`, `<w:pgNumType w:start="1"/>`);
    const breaks = breaksOf(r);
    expect(breaks).toHaveLength(2);
    expect(breaks[0]!.style.sectionBreak!.props.pageNumberStart).toBeUndefined();
    expect(breaks[1]!.style.sectionBreak!.props.pageNumberStart).toBe(1);
  });

  it("keeps the final break when the body section itself is line-numbered", () => {
    // No later in-paragraph break follows, so the 'following section' is the body
    // sectPr — its line numbering must not bleed back onto the intro.
    const r = runImport(
      simpleDocx(
        `<w:p><w:r><w:t>intro</w:t></w:r></w:p>` +
          `<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr></w:p>` +
          `<w:p><w:r><w:t>numbered body</w:t></w:r></w:p>` +
          `<w:sectPr><w:lnNumType w:countBy="1"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`,
      ),
    );
    expect(r.doc.section.lineNumbering).toEqual({ countBy: 1 });
    const breaks = breaksOf(r);
    expect(breaks).toHaveLength(1);
    expect(breaks[0]!.style.sectionBreak!.props.lineNumbering).toBeUndefined();
  });

  it("still flows a geometry-preserving break when neither section is numbered", () => {
    // Guard: the fix must NOT make every plain Next Page break page-breaking —
    // footer-only breaks with matching own-properties continue to flow.
    const r = numberedFollows(`<w:type w:val="nextPage"/>`, `<w:type w:val="nextPage"/>`);
    // Only the body section remains; neither in-paragraph break is materialized.
    expect(breaksOf(r)).toHaveLength(0);
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
