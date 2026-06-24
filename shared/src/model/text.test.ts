import { describe, expect, it } from "vitest";
import type { Block, CharStyle, Document, Paragraph, Run, SectionProps, TableBlock } from "./document";
import {
  bandParagraphs,
  blockById,
  blockExists,
  blockIndexOf,
  bodyParagraphs,
  isHiddenParagraph,
  isInCell,
  locateParagraph,
  nextGrapheme,
  nextWordEnd,
  paragraphAt,
  paragraphsOf,
  prevGrapheme,
  prevWordStart,
  replaceParagraphAt,
  styleAtRuns,
  textOfBlock,
  textOfRuns,
} from "./text";

const section = (extra: Partial<SectionProps> = {}): SectionProps => ({
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
  ...extra,
});
const CHAR: CharStyle = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA = { align: "left" as const, lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };

const para = (id: string, text: string): Paragraph => ({ kind: "paragraph", id, revision: 0, runs: [{ text, style: CHAR }], style: PARA });
const paraRuns = (id: string, runs: Run[]): Paragraph => ({ kind: "paragraph", id, revision: 0, runs, style: PARA });
const docOf = (...blocks: Block[]): Document => ({ section: section(), blocks });

/** A 1x1 table holding the given paragraphs in its single cell. */
const tableWith = (id: string, paras: Paragraph[]): TableBlock => ({
  kind: "table",
  id,
  revision: 0,
  rows: [{ cells: [{ id: `${id}-c0`, blocks: paras }] }],
});

describe("isHiddenParagraph", () => {
  it("is true when every non-empty run is hidden", () => {
    const p = paraRuns("p", [{ text: "secret", style: { ...CHAR, hidden: true } }]);
    expect(isHiddenParagraph(p)).toBe(true);
  });

  it("is false for a truly empty paragraph (no text at all)", () => {
    const p = paraRuns("p", [{ text: "", style: { ...CHAR, hidden: true } }]);
    expect(isHiddenParagraph(p)).toBe(false);
  });

  it("is false when any non-empty run is visible", () => {
    const p = paraRuns("p", [
      { text: "vis", style: CHAR },
      { text: "hid", style: { ...CHAR, hidden: true } },
    ]);
    expect(isHiddenParagraph(p)).toBe(false);
  });

  it("ignores empty runs when judging hiddenness (empty visible run does not unhide)", () => {
    const p = paraRuns("p", [
      { text: "", style: CHAR },
      { text: "secret", style: { ...CHAR, hidden: true } },
    ]);
    expect(isHiddenParagraph(p)).toBe(true);
  });
});

describe("bodyParagraphs", () => {
  it("returns top-level paragraphs in document order", () => {
    const doc = docOf(para("a", "A"), para("b", "B"));
    expect(bodyParagraphs(doc).map((p) => p.id)).toEqual(["a", "b"]);
  });

  it("includes table-cell paragraphs one level deep, in order", () => {
    const doc = docOf(para("a", "A"), tableWith("t", [para("c0", "C0"), para("c1", "C1")]), para("b", "B"));
    expect(bodyParagraphs(doc).map((p) => p.id)).toEqual(["a", "c0", "c1", "b"]);
  });
});

describe("bandParagraphs", () => {
  it("returns header band paragraphs including cell paragraphs", () => {
    const doc: Document = {
      section: section({ header: [para("h0", "H0"), tableWith("ht", [para("hc", "HC")])] }),
      blocks: [para("body", "B")],
    };
    expect(bandParagraphs(doc, "header").map((p) => p.id)).toEqual(["h0", "hc"]);
  });

  it("returns empty array for an absent band", () => {
    expect(bandParagraphs(docOf(para("a", "A")), "footerEven")).toEqual([]);
  });
});

describe("paragraphsOf", () => {
  it("orders body (incl. cells), then bands, then footnotes", () => {
    const doc: Document = {
      section: section({ header: [para("h", "H")], footer: [para("f", "F")] }),
      blocks: [para("a", "A"), tableWith("t", [para("c", "C")])],
      footnotes: { n1: [para("fn", "FN")] },
    };
    expect(paragraphsOf(doc).map((p) => p.id)).toEqual(["a", "c", "h", "f", "fn"]);
  });
});

describe("blockById / blockIndexOf", () => {
  it("finds a paragraph by id and its index in paragraphsOf order", () => {
    const doc = docOf(para("a", "A"), tableWith("t", [para("c", "C")]), para("b", "B"));
    expect(blockById(doc, "c")?.id).toBe("c");
    expect(blockIndexOf(doc, "c")).toBe(1);
    expect(blockIndexOf(doc, "b")).toBe(2);
  });

  it("returns undefined / -1 for an unknown id", () => {
    const doc = docOf(para("a", "A"));
    expect(blockById(doc, "missing")).toBeUndefined();
    expect(blockIndexOf(doc, "missing")).toBe(-1);
  });
});

describe("blockExists", () => {
  it("finds top-level, table, cell, band, and footnote ids", () => {
    const doc: Document = {
      section: section({ header: [para("h", "H")] }),
      blocks: [para("a", "A"), tableWith("t", [para("c", "C")])],
      footnotes: { n1: [para("fn", "FN")] },
    };
    for (const id of ["a", "t", "c", "h", "fn"]) {
      expect(blockExists(doc, id)).toBe(true);
    }
  });

  it("does NOT treat a TableCell id as a block id (cells are not blocks)", () => {
    const doc = docOf(para("a", "A"), tableWith("t", [para("c", "C")]));
    expect(blockExists(doc, "t-c0")).toBe(false);
  });

  it("is false for an absent id", () => {
    expect(blockExists(docOf(para("a", "A")), "nope")).toBe(false);
  });
});

describe("locateParagraph", () => {
  it("locates a top-level body paragraph", () => {
    const doc = docOf(para("a", "A"), para("b", "B"));
    expect(locateParagraph(doc, "b")).toEqual({ kind: "top", bi: 1 });
  });

  it("locates a table-cell paragraph with full coordinates", () => {
    const doc = docOf(para("a", "A"), tableWith("t", [para("c", "C")]));
    expect(locateParagraph(doc, "c")).toEqual({ kind: "cell", where: "body", bi: 1, ri: 0, ci: 0, pi: 0 });
  });

  it("locates a band paragraph", () => {
    const doc: Document = { section: section({ footer: [para("x", "x"), para("f", "F")] }), blocks: [para("a", "A")] };
    expect(locateParagraph(doc, "f")).toEqual({ kind: "band", band: "footer", bi: 1 });
  });

  it("locates a footnote paragraph", () => {
    const doc: Document = { section: section(), blocks: [para("a", "A")], footnotes: { n1: [para("p0", "0"), para("p1", "1")] } };
    expect(locateParagraph(doc, "p1")).toEqual({ kind: "footnote", noteId: "n1", pi: 1 });
  });

  it("returns null for an unknown id", () => {
    expect(locateParagraph(docOf(para("a", "A")), "missing")).toBeNull();
  });
});

describe("paragraphAt", () => {
  it("round-trips with locateParagraph for body, cell, band, footnote", () => {
    const doc: Document = {
      section: section({ header: [para("h", "H")] }),
      blocks: [para("a", "A"), tableWith("t", [para("c", "C")])],
      footnotes: { n1: [para("fn", "FN")] },
    };
    for (const id of ["a", "c", "h", "fn"]) {
      const loc = locateParagraph(doc, id)!;
      expect(paragraphAt(doc, loc).id).toBe(id);
    }
  });
});

describe("replaceParagraphAt", () => {
  it("replaces a top-level paragraph without mutating the input", () => {
    const doc = docOf(para("a", "A"), para("b", "B"));
    const loc = locateParagraph(doc, "b")!;
    const next = replaceParagraphAt(doc, loc, para("b", "B2"));
    expect(textOfBlock(next, "b")).toBe("B2");
    // input untouched
    expect(textOfBlock(doc, "b")).toBe("B");
    expect(doc.blocks).not.toBe(next.blocks);
  });

  it("replaces a cell paragraph, bumps the table revision, and leaves the input intact", () => {
    const doc = docOf(para("a", "A"), tableWith("t", [para("c", "C")]));
    const loc = locateParagraph(doc, "c")!;
    const next = replaceParagraphAt(doc, loc, para("c", "C2"));
    expect(textOfBlock(next, "c")).toBe("C2");
    const origTable = doc.blocks[1] as TableBlock;
    const newTable = next.blocks[1] as TableBlock;
    expect(newTable.revision).toBe(origTable.revision + 1);
    // input untouched
    expect(textOfBlock(doc, "c")).toBe("C");
    expect(origTable.revision).toBe(0);
  });

  it("replaces a band paragraph immutably", () => {
    const doc: Document = { section: section({ footer: [para("f", "F")] }), blocks: [para("a", "A")] };
    const loc = locateParagraph(doc, "f")!;
    const next = replaceParagraphAt(doc, loc, para("f", "F2"));
    expect(textOfBlock(next, "f")).toBe("F2");
    expect(textOfBlock(doc, "f")).toBe("F");
    expect(doc.section.footer).not.toBe(next.section.footer);
  });

  it("replaces a footnote paragraph immutably", () => {
    const doc: Document = { section: section(), blocks: [para("a", "A")], footnotes: { n1: [para("fn", "FN")] } };
    const loc = locateParagraph(doc, "fn")!;
    const next = replaceParagraphAt(doc, loc, para("fn", "FN2"));
    expect(textOfBlock(next, "fn")).toBe("FN2");
    expect(textOfBlock(doc, "fn")).toBe("FN");
    expect(doc.footnotes).not.toBe(next.footnotes);
  });
});

describe("isInCell", () => {
  it("is true for a table-cell paragraph, false for a body paragraph", () => {
    const doc = docOf(para("a", "A"), tableWith("t", [para("c", "C")]));
    expect(isInCell(doc, "c")).toBe(true);
    expect(isInCell(doc, "a")).toBe(false);
  });

  it("is false for an unknown id", () => {
    expect(isInCell(docOf(para("a", "A")), "missing")).toBe(false);
  });
});

describe("textOfRuns / textOfBlock", () => {
  it("concatenates run text in order", () => {
    expect(textOfRuns([{ text: "foo", style: CHAR }, { text: "bar", style: CHAR }])).toBe("foobar");
    expect(textOfRuns([])).toBe("");
  });

  it("textOfBlock returns the block text, or empty string for an unknown id", () => {
    const doc = docOf(paraRuns("a", [{ text: "Hello, ", style: CHAR }, { text: "world", style: CHAR }]));
    expect(textOfBlock(doc, "a")).toBe("Hello, world");
    expect(textOfBlock(doc, "missing")).toBe("");
  });
});

describe("prevGrapheme / nextGrapheme", () => {
  it("steps one ASCII character", () => {
    expect(prevGrapheme("abc", 2)).toBe(1);
    expect(nextGrapheme("abc", 1)).toBe(2);
  });

  it("clamps at the boundaries", () => {
    expect(prevGrapheme("abc", 0)).toBe(0);
    expect(nextGrapheme("abc", 3)).toBe(3);
    expect(nextGrapheme("abc", 99)).toBe(3);
  });

  it("treats a surrogate-pair emoji as one grapheme (2 UTF-16 units)", () => {
    const s = "a\u{1F600}b"; // a + grin + b, grin is offsets 1..3
    expect(nextGrapheme(s, 1)).toBe(3);
    expect(prevGrapheme(s, 3)).toBe(1);
  });

  it("treats a ZWJ family emoji as a single grapheme cluster", () => {
    const family = "\u{1F468}‍\u{1F469}‍\u{1F467}"; // man+ZWJ+woman+ZWJ+girl
    const s = `${family}!`;
    // a single next-step from 0 jumps the whole cluster
    expect(nextGrapheme(s, 0)).toBe(family.length);
    expect(prevGrapheme(s, family.length)).toBe(0);
  });
});

describe("prevWordStart / nextWordEnd", () => {
  it("jumps to word boundaries within a sentence", () => {
    const s = "foo bar baz"; // foo[0..3] bar[4..7] baz[8..11]
    // from inside "bar", previous word start is "bar" at 4
    expect(prevWordStart(s, 6)).toBe(4);
    // from the start of "bar", next word end is end of "bar" at 7
    expect(nextWordEnd(s, 4)).toBe(7);
  });

  it("skips whitespace/punctuation to the next word-like segment", () => {
    const s = "foo, bar";
    // offset 3 is at the comma; next word-like end is end of "bar" (8)
    expect(nextWordEnd(s, 3)).toBe(8);
  });

  it("clamps at the ends", () => {
    const s = "foo bar";
    expect(prevWordStart(s, 0)).toBe(0);
    expect(nextWordEnd(s, 7)).toBe(7); // already at end, no further word-like end
  });
});

describe("styleAtRuns", () => {
  const bold: CharStyle = { ...CHAR, bold: true };
  const italic: CharStyle = { ...CHAR, italic: true };

  it("returns the style of the run preceding the caret", () => {
    const runs: Run[] = [{ text: "foo", style: bold }, { text: "bar", style: italic }];
    // offset 1 falls inside the first run
    expect(styleAtRuns(runs, 1)).toBe(bold);
    // offset 3 = end of first run, still inherits the first run's style
    expect(styleAtRuns(runs, 3)).toBe(bold);
    // offset 4 falls inside the second run
    expect(styleAtRuns(runs, 4)).toBe(italic);
  });

  it("at offset 0 falls back to the first run's style", () => {
    const runs: Run[] = [{ text: "foo", style: bold }, { text: "bar", style: italic }];
    expect(styleAtRuns(runs, 0)).toBe(bold);
  });

  it("returns undefined for no runs", () => {
    expect(styleAtRuns([], 0)).toBeUndefined();
  });

  it("skips empty runs when finding the preceding style", () => {
    const runs: Run[] = [{ text: "foo", style: bold }, { text: "", style: italic }, { text: "bar", style: bold }];
    // offset 3 = boundary; empty run is skipped, first run's style precedes
    expect(styleAtRuns(runs, 3)).toBe(bold);
  });
});
