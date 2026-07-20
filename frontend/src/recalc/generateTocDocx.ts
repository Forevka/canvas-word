// Headless TOC field-RESULT generation, patch-in-place (drift-free). A caller
// emits a Word `TOC` field with its instruction and an empty/placeholder result;
// we generate the entries Word would render on F9 — styled, with live PAGEREF
// page numbers and hyperlinks — and splice them into the field's result region in
// the ORIGINAL document.xml, re-zipping. Every other byte is preserved, so there
// is no whole-document re-export and no pagination drift.
//
// Page numbers account for the TOC's own height: we splice the entries in, re-import
// + lay out the result, read each heading's page, then splice once more with the
// real numbers (entry height is independent of the digit, so this converges).

import { strFromU8, strToU8, unzipSync, type Unzipped, zipSync } from "fflate";
import {
  detectTocHeadings,
  inRange,
  parseTocInstruction,
  textOfRuns,
  tocEntryStyle,
  type Document,
  type TocOptions,
  type TocSwitches,
} from "@cw/shared";
import { findMainDocumentPart } from "../import/docx/contentTypes";
import { runImport } from "../import/docx/pipeline";
import { createLayoutEngine } from "../layout/engine";
import { pageOfBlockMap } from "./recalcToc";
import { paraCoreXml, runPropsXml } from "../export/docx/styleProps";
import { orderChildren, PPR_ORDER } from "../export/docx/schemaOrder";
import { el, escapeText, textEl } from "../export/docx/xmlWrite";

// --- OOXML serialization for the generated entries (full rPr/pPr shared with the
// main exporter via styleProps so the two can't drift) -------------------------

const fld = (t: "begin" | "separate" | "end"): string => el("w:r", undefined, el("w:fldChar", { "w:fldCharType": t }));
const instrRun = (i: string): string => el("w:r", undefined, el("w:instrText", { "xml:space": "preserve" }, escapeText(i)));

function pagerefXml(anchor: string, num: string, rPr: string): string {
  return fld("begin") + instrRun(` PAGEREF ${anchor} \\h `) + fld("separate") + el("w:r", undefined, rPr + (num ? textEl(num) : "")) + fld("end");
}

interface EntryPlan {
  label: string;
  level: number;
  anchor?: string | undefined; // heading bookmark (when a page number / hyperlink is wanted)
  numText: string; // "" when the page number is hidden (\n) or unresolved
}

function entryXml(
  e: EntryPlan,
  opts: TocOptions,
  contentWidthPx: number,
  ctx: { instr: string; first: boolean; last: boolean; hyperlink: boolean; separator?: string },
): string {
  const { char, para } = tocEntryStyle(opts, e.level, contentWidthPx);
  const rPr = runPropsXml(char);
  let body = el("w:pPr", undefined, orderChildren(paraCoreXml(para), PPR_ORDER).join(""));
  if (ctx.first) body += fld("begin") + instrRun(ctx.instr) + fld("separate");
  const sep = ctx.separator && ctx.separator !== "\t" ? el("w:r", undefined, rPr + textEl(ctx.separator)) : el("w:r", undefined, rPr + el("w:tab"));
  const num = e.numText
    ? (e.anchor ? pagerefXml(e.anchor, e.numText, rPr) : el("w:r", undefined, rPr + textEl(e.numText)))
    : "";
  let entry = el("w:r", undefined, rPr + textEl(e.label)) + (num ? sep + num : "");
  if (ctx.hyperlink && e.anchor) entry = el("w:hyperlink", { "w:anchor": e.anchor, "w:history": "1" }, entry);
  body += entry;
  if (ctx.last) body += fld("end");
  return el("w:p", undefined, body);
}

// --- locating the TOC field result region + heading paragraphs in raw XML ------

const FIELD_TOKEN = /<w:fldChar\b[^>]*?w:fldCharType="(begin|separate|end)"[^>]*?\/?>|<w:instrText\b[^>]*?>([\s\S]*?)<\/w:instrText>/g;
const P_OPEN = /<w:p(?:\s[^>]*)?>/g;

interface TocRegion { pStart: number; pEnd: number; instr: string }

/** Find the first `TOC` field and the paragraph span [begin's <w:p> … end's </w:p>]. */
function findTocRegion(xml: string): TocRegion | null {
  const stack: { instr: string; beginIdx: number }[] = [];
  let beginIdx = -1, endIdx = -1, instr = "";
  FIELD_TOKEN.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = FIELD_TOKEN.exec(xml)) !== null) {
    if (m[1] === "begin") stack.push({ instr: "", beginIdx: m.index });
    else if (m[1] === "end") {
      const f = stack.pop();
      if (f && beginIdx < 0 && /\bTOC\b/.test(f.instr)) {
        beginIdx = f.beginIdx;
        endIdx = m.index + m[0].length;
        instr = f.instr;
      }
    } else if (m[2] !== undefined) {
      const top = stack[stack.length - 1];
      if (top) top.instr += m[2];
    }
  }
  if (beginIdx < 0) return null;
  let pStart = -1;
  P_OPEN.lastIndex = 0;
  let pm: RegExpExecArray | null;
  while ((pm = P_OPEN.exec(xml)) !== null) {
    if (pm.index < beginIdx) pStart = pm.index;
    else break;
  }
  if (pStart < 0) return null;
  const close = xml.indexOf("</w:p>", endIdx);
  if (close < 0) return null;
  return { pStart, pEnd: close + "</w:p>".length, instr };
}

interface Edit { start: number; end: number; text: string }
const applyEdits = (xml: string, edits: Edit[]): string => {
  const sorted = [...edits].sort((a, b) => b.start - a.start);
  let out = xml;
  for (const e of sorted) out = out.slice(0, e.start) + e.text + out.slice(e.end);
  return out;
};

/** Inject a point bookmark into a heading paragraph (after its w:pPr). */
function bookmarkInjections(
  xml: string,
  headingStyleIds: Set<string>,
  wants: Map<number, { name: string; id: number }>, // heading order-index → bookmark
): Edit[] {
  if (wants.size === 0) return [];
  const edits: Edit[] = [];
  // Heading paragraphs in document order: a <w:p> whose pPr references a heading style id.
  const pRe = /<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g;
  let order = -1;
  let m: RegExpExecArray | null;
  while ((m = pRe.exec(xml)) !== null) {
    const inner = m[1]!;
    const styleId = inner.match(/<w:pStyle\b[^>]*\bw:val="([^"]+)"/)?.[1];
    if (!styleId || !headingStyleIds.has(styleId)) continue;
    order++;
    const want = wants.get(order);
    if (!want) continue;
    const pprEnd = inner.indexOf("</w:pPr>");
    const innerStart = m.index + m[0].indexOf(inner);
    const startMark = el("w:bookmarkStart", { "w:id": want.id, "w:name": want.name });
    const endMark = el("w:bookmarkEnd", { "w:id": want.id });
    if (pprEnd >= 0) {
      const at = innerStart + pprEnd + "</w:pPr>".length;
      edits.push({ start: at, end: at, text: startMark });
    } else {
      edits.push({ start: innerStart, end: innerStart, text: startMark });
    }
    const endAt = innerStart + inner.length;
    edits.push({ start: endAt, end: endAt, text: endMark });
  }
  return edits;
}

export interface GenerateTocResult {
  bytes: Uint8Array;
  /** Entries written into the field result. */
  generated: number;
  /** Headings detected in the document. */
  headings: number;
  /** Heading bookmarks synthesized + injected (those that had none). */
  bookmarksSynthesized: number;
}

/** Generate the TOC field result in `docxBytes` and return the patched .docx. The
 *  caller MUST have awaited installMeasureHost() (layout measures text). */
export function generateTocInDocx(docxBytes: Uint8Array, opts: TocOptions = {}): GenerateTocResult {
  const parts: Unzipped = unzipSync(docxBytes);
  const ctXml = parts["[Content_Types].xml"];
  const mainPart = (ctXml ? findMainDocumentPart(strFromU8(ctXml)) : undefined) ?? "word/document.xml";
  const docPart = parts[mainPart];
  if (!docPart) return { bytes: docxBytes, generated: 0, headings: 0, bookmarksSynthesized: 0 };
  const xml0 = strFromU8(docPart);

  const region = findTocRegion(xml0);
  if (!region) return { bytes: docxBytes, generated: 0, headings: 0, bookmarksSynthesized: 0 };
  const switches: TocSwitches = parseTocInstruction(region.instr);

  const doc: Document = runImport(docxBytes).doc;
  const maxLevel = switches.outlineRange?.to ?? opts.maxLevel ?? 3;
  const headings = detectTocHeadings(doc, switches, maxLevel);
  const contentWidthPx = doc.section.pageWidthPx - doc.section.marginPx.left - doc.section.marginPx.right;
  const hyperlink = opts.hyperlink ?? switches.hyperlinks;
  const showNumbers = opts.includePageNumbers ?? true;

  // Each heading's bookmark: reuse an existing _Toc* (or any) bookmark, else synth.
  const nameByBlock = new Map<string, string>();
  for (const [name, range] of Object.entries(doc.bookmarks ?? {})) {
    const cur = nameByBlock.get(range.start.blockId);
    if (!cur || (!cur.startsWith("_Toc") && name.startsWith("_Toc"))) nameByBlock.set(range.start.blockId, name);
  }
  const headingStyleIds = new Set<string>();
  for (const h of headings) if (h.block.style.namedStyle) headingStyleIds.add(h.block.style.namedStyle);

  const synth = new Map<number, { name: string; id: number }>(); // heading order-index → injected bookmark
  let bmId = 900000;
  let synthN = 0;
  const plans: EntryPlan[] = headings.map((h, i) => {
    const label = textOfRuns(h.block.runs).replace(/\v/g, " ").trim();
    const wantNumber = showNumbers && !inRange(switches.hidePageNumberRange, h.level);
    let anchor: string | undefined;
    if (wantNumber || hyperlink) {
      anchor = nameByBlock.get(h.block.id);
      if (!anchor) {
        anchor = `_CwToc${synthN++}`;
        synth.set(i, { name: anchor, id: bmId++ });
      }
    }
    return { label, level: h.level, anchor, numText: wantNumber ? "1" : "" };
  }).filter((p) => p.label.length > 0);

  if (plans.length === 0) return { bytes: docxBytes, generated: 0, headings: headings.length, bookmarksSynthesized: 0 };

  const bmEdits = bookmarkInjections(xml0, headingStyleIds, synth);

  const renderField = (entries: EntryPlan[]): string =>
    entries
      .map((e, i) => entryXml(e, opts, contentWidthPx, {
        instr: region.instr, first: i === 0, last: i === entries.length - 1, hyperlink,
        ...(switches.separator !== undefined ? { separator: switches.separator } : {}),
      }))
      .join("");

  const splice = (entries: EntryPlan[]): string =>
    applyEdits(xml0, [{ start: region.pStart, end: region.pEnd, text: renderField(entries) }, ...bmEdits]);

  // Pass 1: splice with placeholder numbers so the TOC's height is in the layout.
  const xml1 = splice(plans);
  const parts1: Unzipped = { ...parts, [mainPart]: strToU8(xml1) };
  const bytes1 = zipSync(parts1);

  // Lay the result out to get each heading's real page (via its bookmark).
  const doc2 = runImport(bytes1).doc;
  const pageOfBlock = pageOfBlockMap(createLayoutEngine().layout(doc2));
  const blockByBookmark = new Map<string, string>();
  for (const [name, range] of Object.entries(doc2.bookmarks ?? {})) blockByBookmark.set(name, range.start.blockId);
  for (const p of plans) {
    if (!p.numText || !p.anchor) continue;
    const blk = blockByBookmark.get(p.anchor);
    const page = blk !== undefined ? pageOfBlock.get(blk) : undefined;
    p.numText = page !== undefined ? String(page) : "";
  }

  // Pass 2: final splice with real numbers, into the original XML.
  parts[mainPart] = strToU8(splice(plans));
  return {
    bytes: zipSync(parts),
    generated: plans.length,
    headings: headings.length,
    bookmarksSynthesized: synth.size,
  };
}
