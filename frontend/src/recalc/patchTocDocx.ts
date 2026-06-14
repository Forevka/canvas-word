// Patch a .docx's TOC page numbers IN PLACE — the faithful path for fixing an
// externally-authored document (e.g. a C# render pipeline that emits a real Word
// TOC field but can't compute page numbers). We import only to RUN THE LAYOUT and
// learn each heading's page; we then surgically rewrite the cached PAGEREF result
// digits in the ORIGINAL document.xml and re-zip. Everything else — the live TOC
// field, PAGEREF/HYPERLINK fields, footer PAGE/NUMPAGES, styles, and any OOXML our
// model doesn't represent — survives byte-for-byte. Because we never regenerate the
// file through the export writer, there is no import↔export pagination drift.
//
// Limitation (documented, intentional): only arabic (\d+) cached page numbers are
// recalculated. A roman/alpha cached value is left untouched and counted in
// `skipped` rather than being overwritten with a wrong arabic number.

import { strFromU8, strToU8, unzipSync, zipSync, type Unzipped } from "fflate";
import type { Document } from "@cw/shared";
import { findMainDocumentPart } from "../import/docx/contentTypes";
import { createLayoutEngine } from "../layout/engine";
import type { LayoutTree } from "../layout/layoutTree";
import { pageOfBlockMap } from "./recalcToc";

/** bookmark name → displayed page number (as a string) of the block it starts at.
 *  The PAGEREF instruction in the XML references the same bookmark name, so this is
 *  how the layout's per-block page maps onto the document's fields. */
export function pageByBookmark(doc: Document, tree: LayoutTree): Record<string, string> {
  const pageOfBlock = pageOfBlockMap(tree);
  const out: Record<string, string> = {};
  for (const [name, range] of Object.entries(doc.bookmarks ?? {})) {
    const page = pageOfBlock.get(range.start.blockId);
    if (page !== undefined) out[name] = String(page);
  }
  return out;
}

export interface PatchResult {
  bytes: Uint8Array;
  /** PAGEREF cached numbers rewritten to a new value. */
  changed: number;
  /** PAGEREF results left untouched because the cached value was non-arabic. */
  skipped: number;
}

const PAGEREF_RE = /\bPAGEREF\s+("[^"]+"|\S+)/;

// fldChar (begin|separate|end) | instrText(content) | w:t(open)(content)(close)
const TOKEN_RE =
  /<w:fldChar\b[^>]*?w:fldCharType="(begin|separate|end)"[^>]*?\/?>|<w:instrText\b[^>]*?>([\s\S]*?)<\/w:instrText>|(<w:t\b[^>]*?>)([\s\S]*?)(<\/w:t>)/g;

interface Frame {
  instr: string;
  anchor?: string; // set when the field is a PAGEREF
  sawSeparate: boolean;
  done: boolean; // first result run handled
  replaced: boolean; // we rewrote the number → blank any split remainder
}

/** Rewrite cached PAGEREF result digits in one document-part XML string. */
function patchXml(xml: string, pages: Record<string, string>): { xml: string; changed: number; skipped: number } {
  const stack: Frame[] = [];
  let out = "";
  let cursor = 0;
  let changed = 0;
  let skipped = 0;
  TOKEN_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = TOKEN_RE.exec(xml)) !== null) {
    const top = stack[stack.length - 1];
    let replacement: string | null = null; // non-null → replace this match's slice

    if (m[1] !== undefined) {
      // w:fldChar
      if (m[1] === "begin") stack.push({ instr: "", sawSeparate: false, done: false, replaced: false });
      else if (m[1] === "separate") {
        if (top) top.sawSeparate = true;
      } else stack.pop(); // "end"
    } else if (m[2] !== undefined) {
      // w:instrText
      if (top) {
        top.instr += m[2];
        const pm = top.instr.match(PAGEREF_RE);
        if (pm) top.anchor = pm[1]!.replace(/^"|"$/g, "");
      }
    } else {
      // w:t — group 3 open, 4 content, 5 close
      if (top?.anchor && top.sawSeparate) {
        if (!top.done) {
          top.done = true;
          const num = pages[top.anchor];
          if (num !== undefined && /^\d+$/.test(m[4]!)) {
            if (m[4] !== num) {
              replacement = m[3]! + num + m[5]!;
              changed++;
              top.replaced = true;
            }
          } else if (num !== undefined && !/^\d+$/.test(m[4]!.trim()) && m[4]!.trim().length > 0) {
            skipped++; // non-arabic cached value — leave it (unsupported)
          }
        } else if (top.replaced && m[4] !== "") {
          // Split number ("1","2"): the full value went into the first run; blank the rest.
          replacement = m[3]! + m[5]!;
        }
      }
    }

    if (replacement !== null) {
      out += xml.slice(cursor, m.index) + replacement;
      cursor = m.index + m[0].length;
    }
  }
  out += xml.slice(cursor);
  return { xml: out, changed, skipped };
}

/** Patch the cached TOC/cross-reference page numbers in a .docx, preserving every
 *  field and every other part. `pages` maps bookmark name → new page number. */
export function patchTocPageNumbers(docxBytes: Uint8Array, pages: Record<string, string>): PatchResult {
  const parts: Unzipped = unzipSync(docxBytes);
  const ctXml = parts["[Content_Types].xml"];
  const mainPart =
    (ctXml ? findMainDocumentPart(strFromU8(ctXml)) : undefined) ?? "word/document.xml";
  const docPart = parts[mainPart];
  if (!docPart) return { bytes: docxBytes, changed: 0, skipped: 0 };

  const { xml, changed, skipped } = patchXml(strFromU8(docPart), pages);
  if (changed === 0) return { bytes: docxBytes, changed: 0, skipped };

  parts[mainPart] = strToU8(xml);
  return { bytes: zipSync(parts), changed, skipped };
}

/** Convenience: lay the doc out and patch the original bytes from it. The caller
 *  MUST have awaited installMeasureHost() first (layout measures text). */
export function patchTocFromLayout(docxBytes: Uint8Array, doc: Document): PatchResult {
  const tree = createLayoutEngine().layout(doc);
  return patchTocPageNumbers(docxBytes, pageByBookmark(doc, tree));
}
