import { describe, expect, it } from "vitest";
import type { Document, Paragraph } from "@cw/shared";
import { applyOp } from "@cw/shared";
import { runImport } from "../import/docx/pipeline";
import { CONTENT_TYPES_XML, documentXml, drawingXml, makeDocx, PNG_1PX, REL_TYPES, relsXml, simpleDocx } from "../import/docx/fixture";
import { parseOoxmlFragment } from "../import/docx/fragment";
import { fieldAtBlock, insertTocCmd, replaceFieldResultCmd, updateTocFieldCmd } from "./commands";
import type { EditorState } from "./state";

const fld = (t: string): string => `<w:r><w:fldChar w:fldCharType="${t}"/></w:r>`;
const instr = (s: string): string => `<w:r><w:instrText xml:space="preserve">${s}</w:instrText></w:r>`;
const txt = (s: string): string => `<w:r><w:t>${s}</w:t></w:r>`;

const apply = (state: EditorState, cmd: ReturnType<typeof replaceFieldResultCmd>): Document => {
  const trn = cmd(state);
  if (!trn) throw new Error("no transaction");
  let doc = state.doc;
  for (const op of trn.ops) doc = applyOp(doc, op).doc;
  return doc;
};

const customFieldDoc = (): Document =>
  runImport(
    simpleDocx(
      `<w:p>${txt("Intro")}</w:p>` +
        `<w:p>${fld("begin")}${instr(' MYCHART "x" ')}${fld("separate")}${txt("[old]")}${fld("end")}</w:p>` +
        `<w:p>${txt("Outro")}</w:p>`,
    ),
  ).doc;

describe("parseOoxmlFragment", () => {
  it("parses a fragment of paragraphs + a table into blocks", () => {
    const blocks = parseOoxmlFragment(
      `<w:p>${txt("A")}</w:p>` +
        `<w:tbl><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>` +
        `<w:tr><w:tc><w:p>${txt("cell")}</w:p></w:tc></w:tr></w:tbl>` +
        `<w:p>${txt("B")}</w:p>`,
    );
    expect(blocks.map((b) => b.kind)).toEqual(["paragraph", "table", "paragraph"]);
  });
});

describe("fieldAtBlock", () => {
  it("classifies custom-field blocks and ignores plain ones", () => {
    const doc = customFieldDoc();
    const fieldPara = doc.blocks.find((b): b is Paragraph => b.kind === "paragraph" && !!b.fieldId)!;
    const at = fieldAtBlock(doc, fieldPara.id);
    expect(at).toEqual({ kind: "custom", def: doc.fields![fieldPara.fieldId!] });

    const plain = doc.blocks.find((b): b is Paragraph => b.kind === "paragraph" && !b.fieldId)!;
    expect(fieldAtBlock(doc, plain.id)).toBeNull();
  });
});

describe("replaceFieldResultCmd", () => {
  it("swaps the field result, re-stamps fieldId, and keeps the field definition", () => {
    const doc = customFieldDoc();
    const fieldId = Object.keys(doc.fields!)[0]!;
    const fresh = parseOoxmlFragment(`<w:p>${txt("New A")}</w:p><w:p>${txt("New B")}</w:p>`);

    const out = apply({ doc, selection: null }, replaceFieldResultCmd(fieldId, fresh));

    // Definition preserved.
    expect(out.fields![fieldId]!.instruction).toBe(' MYCHART "x" ');
    // The result region is now the two new paragraphs, both tagged, in place.
    const tagged = out.blocks.filter((b): b is Paragraph => b.kind === "paragraph" && b.fieldId === fieldId);
    expect(tagged.map((p) => p.runs.map((r) => r.text).join(""))).toEqual(["New A", "New B"]);
    // Surrounding content untouched and order preserved.
    const allText = out.blocks.map((b) => (b.kind === "paragraph" ? b.runs.map((r) => r.text).join("") : "")).filter(Boolean);
    expect(allText).toEqual(["Intro", "New A", "New B", "Outro"]);
  });

  it("accepts a full .docx result (with an image) — blocks splice in, image tagged", () => {
    // A full document.docx is how a resolver passes images/rich content: imported
    // via the normal pipeline (here runImport, the core importDocx wraps) → blocks.
    const imageDocx = makeDocx({
      "[Content_Types].xml": CONTENT_TYPES_XML,
      "word/document.xml": documentXml(`<w:p><w:r>${drawingXml("rId1", 914400, 914400)}</w:r></w:p>`),
      "word/_rels/document.xml.rels": relsXml([{ id: "rId1", type: REL_TYPES.image, target: "media/image1.png" }]),
      "word/media/image1.png": PNG_1PX,
    });
    const fresh = runImport(imageDocx, undefined, { collectMediaBytes: true }).doc.blocks;
    expect(fresh.some((b) => b.kind === "image")).toBe(true);

    const doc = customFieldDoc();
    const fieldId = Object.keys(doc.fields!)[0]!;
    const out = apply({ doc, selection: null }, replaceFieldResultCmd(fieldId, fresh));
    const img = out.blocks.find((b) => b.kind === "image");
    expect(img).toBeDefined();
    expect(img!.fieldId).toBe(fieldId); // image carries field membership
    expect(out.fields![fieldId]!.name).toBe("MYCHART"); // definition intact
  });
});

describe("updateTocFieldCmd (Word F9) preserves the imported TOC look", () => {
  // A real imported TOC: flattened entries (styled distinctively) + headings with
  // _Toc bookmarks. Regenerating must NOT inject a "Table of Contents" title and
  // must keep the entries' own styling — not the editor's Georgia defaults.
  const entry = (anchor: string, label: string, page: string): string =>
    `<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Verdana"/><w:sz w:val="30"/></w:rPr></w:pPr>` +
    `<w:hyperlink w:anchor="${anchor}"><w:r><w:rPr><w:rFonts w:ascii="Verdana"/><w:sz w:val="30"/></w:rPr><w:t>${label}</w:t></w:r>` +
    `<w:r><w:tab/></w:r><w:r><w:t>${page}</w:t></w:r></w:hyperlink></w:p>`;
  const head = (anchor: string, label: string): string => {
    const id = anchor.replace(/\D/g, "");
    return `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="${id}" w:name="${anchor}"/>` +
      `<w:r><w:t>${label}</w:t></w:r><w:bookmarkEnd w:id="${id}"/></w:p>`;
  };

  const apply2 = (state: EditorState, cmd: ReturnType<typeof updateTocFieldCmd>): Document => {
    const trn = cmd(state);
    if (!trn) throw new Error("no transaction");
    let doc = state.doc;
    for (const op of trn.ops) doc = applyOp(doc, op).doc;
    return doc;
  };

  it("adds no title and keeps the entry font (15px = sz 30)", () => {
    const body =
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>` +
      entry("_Toc1", "Chapter One", "3") +
      entry("_Toc2", "Chapter Two", "5") +
      head("_Toc1", "Chapter One") +
      head("_Toc2", "Chapter Two");
    const doc = runImport(simpleDocx(body)).doc;
    expect(doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph" && !!b.style.tocEntry).length).toBe(2);

    const out = apply2({ doc, selection: null }, updateTocFieldCmd());
    const ps = out.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
    // No injected "Table of Contents" title.
    expect(ps.some((p) => p.style.namedStyle === "tocTitle" || p.runs.map((r) => r.text).join("") === "Table of Contents")).toBe(false);
    // Entries preserved with their imported font (Verdana, 15px), not Georgia.
    const tocEntries = ps.filter((p) => p.style.tocEntry);
    expect(tocEntries.length).toBe(2);
    expect(tocEntries[0]!.runs[0]!.style.fontFamily).toContain("Verdana");
    expect(tocEntries[0]!.runs[0]!.style.fontSizePx).toBe(20); // sz 30 = 15pt = 20px (preserved, not the 14px default)
    expect(tocEntries[0]!.runs[0]!.style.link).toBeUndefined(); // stale #anchor stripped
  });

  it("the ribbon insert/update (insertTocCmd, no opts) also preserves the look", () => {
    // Regression: the ribbon "Insert / update table of contents" button calls
    // insertTocCmd() with NO options. Regenerating an existing imported TOC must
    // preserve its title presence + entry styling, not clobber it with the editor
    // defaults — the same fix the context-menu "Update Field (TOC)" already had.
    const body =
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>` +
      entry("_Toc1", "Chapter One", "3") +
      entry("_Toc2", "Chapter Two", "5") +
      head("_Toc1", "Chapter One") +
      head("_Toc2", "Chapter Two");
    const doc = runImport(simpleDocx(body)).doc;

    const out = apply2({ doc, selection: null }, insertTocCmd());
    const ps = out.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
    expect(ps.some((p) => p.style.namedStyle === "tocTitle" || p.runs.map((r) => r.text).join("") === "Table of Contents")).toBe(false);
    const tocEntries = ps.filter((p) => p.style.tocEntry);
    expect(tocEntries.length).toBe(2);
    expect(tocEntries[0]!.runs[0]!.style.fontFamily).toContain("Verdana");
    expect(tocEntries[0]!.runs[0]!.style.fontSizePx).toBe(20); // preserved, not the 14px default
  });
});
