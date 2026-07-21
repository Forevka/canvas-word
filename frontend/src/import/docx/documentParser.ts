// document.xml / header / footer body walk → IR. Pure: XML text in, IR + warnings out.
//
// Decodes direct formatting only. Style references (w:pStyle, w:rStyle) are
// recorded in the IR unresolved — styles.ts resolves them into effective
// properties during mapping. Everything the parser understands goes into the
// IR; mapToModel decides what the model can hold.

import type { CharStyle, FieldDef, FieldSpec, ShapePath, ShapePathSegment } from "@cw/shared";
import { isCustomFieldInstruction, parseFieldInstruction, parseFieldSpec } from "@cw/shared";
import { ImportError, WarningSink } from "./types";
import type {
  BandRefs,
  BookmarkMarker,
  IRBlock,
  IRCellMargin,
  IRDocument,
  IRInline,
  IRParagraph,
  IRRunProps,
  IRSdtProps,
  IRSection,
  IRShapeChild,
  IRShapeGroup,
  IRTable,
  IRTableCell,
  IRTableRow,
} from "./types";
import { decodeBorders, decodeShdFill } from "./borders";
import { decodeLineNumbering, decodeParaProps, decodeRunProps } from "./props";
import { attr, children, el, els, findDeep, numAttr, onOff, parseXml, rootEl, textOf, val, type XmlNode } from "./xml";
import { halfPointsToPx } from "./units";
import { ommlToMathml } from "../../mathml/fromOmml";
import { EQUATION_DISPLAY_PX } from "../../layout/math/equationLayout";

interface ParseCtx {
  warnings: WarningSink;
  /** Header/footer mode: PAGE / NUMPAGES fields become live {page}/{pages}
   *  tokens (the layout substitutes them per page). In the body, fields keep
   *  their cached result text — a TOC's "page 5" should stay "5". */
  fieldTokens: boolean;
  /** Content-control registry filled as w:sdt elements are parsed; shared
   *  across body and band parts so ids never collide. */
  sdts: Record<string, IRSdtProps>;
  nextSdt: { n: number };
  /** Ids of the block-level w:sdt ancestors currently open (outer→inner), stamped
   *  onto each block as `sdtPath`. */
  blockSdtStack: string[];
  /** Ids of the inline w:sdt ancestors currently open (outer→inner), snapshotted
   *  onto each run created by `parseRun` as `sdtPath`. */
  inlineSdtStack: string[];
  /** Bookmark names seen between paragraphs (block-level w:bookmarkStart),
   *  attached to the NEXT paragraph. */
  pendingBookmarks: string[];
  /** Bookmark names collected inside the paragraph currently being walked. */
  currentBookmarks: string[] | null;
  /** Block-level bookmark markers, attached at offset 0 of the NEXT paragraph. */
  pendingMarkers: BookmarkMarker[];
  /** Bookmark markers (start/end + offset) inside the paragraph being walked. */
  currentMarkers: BookmarkMarker[] | null;
  /** Capture non-built-in complex fields into a model field registry + stamp the
   *  result blocks' fieldId. True for the body only — bands/footnotes keep their
   *  cached result flattened (custom fields live in the body). */
  trackFields: boolean;
  /** Block-scoped complex-field tracker — survives across paragraphs so a field
   *  whose result spans multiple blocks is captured whole. */
  fieldTrack: FieldTrack;
}

/** Tracks custom (host-resolvable, non-built-in) complex fields at BLOCK scope —
 *  the per-paragraph FieldState below resets each paragraph, but a field's result
 *  can span many blocks (begin in block 1, end in block N). */
interface FieldTrack {
  /** Global complex-field nesting depth (begin++/end--), spanning paragraphs. */
  depth: number;
  /** Instruction text of the current top-level (depth-1) field, accumulated. */
  topInstr: string;
  /** fieldId of the current top-level field IF it's a custom field, else null. */
  openId: string | null;
  /** fieldId to stamp on the block currently being built — the field open when the
   *  block started, or one that opened during it (kept set through the block that
   *  closes the field). */
  markBlock: string | null;
  next: { n: number };
  registry: Record<string, FieldDef>;
}

const newFieldTrack = (): FieldTrack => ({ depth: 0, topInstr: "", openId: null, markBlock: null, next: { n: 0 }, registry: {} });

/** Open a custom field if the just-completed top-level instruction is a non-built-in
 *  one (idempotent — only the first call per field, e.g. at `separate`, opens it). */
function decideCustomField(ctx: ParseCtx): void {
  const ft = ctx.fieldTrack;
  if (ft.openId !== null || !isCustomFieldInstruction(ft.topInstr)) return;
  const id = `field${ft.next.n++}`;
  ft.registry[id] = { id, instruction: ft.topInstr, name: parseFieldInstruction(ft.topInstr).name, kind: "custom" };
  ft.openId = id;
  ft.markBlock = id;
}

/** UTF-16 offset accumulated so far in a paragraph's inline list (runs count
 *  their text; a soft break occupies one offset; images contribute nothing). */
function inlineOffset(out: IRInline[]): number {
  let n = 0;
  for (const inl of out) {
    if (inl.kind === "run") n += inl.text.length;
    else if (inl.kind === "break") n += 1;
  }
  return n;
}

export function parseDocumentXml(xmlText: string, partName: string, warnings: WarningSink): IRDocument {
  const doc = rootEl(parseXml(xmlText, partName), "w:document");
  const body = doc && el(doc, "w:body");
  if (!body) {
    throw new ImportError("MALFORMED_XML", `${partName} has no w:document/w:body root.`);
  }
  const sdts: Record<string, IRSdtProps> = {};
  const ctx: ParseCtx = { warnings, fieldTokens: false, sdts, nextSdt: { n: 0 }, blockSdtStack: [], inlineSdtStack: [], pendingBookmarks: [], currentBookmarks: null, pendingMarkers: [], currentMarkers: null, trackFields: true, fieldTrack: newFieldTrack() };

  const blocks: IRBlock[] = [];
  walkBlocks(children(body), blocks, ctx);

  const sectPr = el(body, "w:sectPr");
  const section = sectPr ? parseSection(sectPr, warnings) : null;

  // w:background is document-global (sibling of w:body) — the page fill color.
  const bgEl = doc ? el(doc, "w:background") : undefined;
  const bgColor = bgEl ? attr(bgEl, "w:color") : undefined;
  const pageColorHex = bgColor && bgColor !== "auto" ? bgColor : undefined;

  const fields = ctx.fieldTrack.registry;
  const out: IRDocument =
    Object.keys(fields).length > 0 ? { blocks, section, sdts, fields } : { blocks, section, sdts };
  if (pageColorHex) out.pageColorHex = pageColorHex;
  return out;
}

/** header1.xml / footer1.xml — same block content under a w:hdr / w:ftr root.
 *  Pass the document's sdt registry so band controls join the same id space. */
export function parseHeaderFooterXml(
  xmlText: string,
  partName: string,
  warnings: WarningSink,
  sdts: Record<string, IRSdtProps> = {},
): IRBlock[] {
  const nodes = parseXml(xmlText, partName);
  const root = rootEl(nodes, "w:hdr") ?? rootEl(nodes, "w:ftr");
  if (!root) {
    throw new ImportError("MALFORMED_XML", `${partName} has no w:hdr/w:ftr root.`);
  }
  const ctx: ParseCtx = { warnings, fieldTokens: true, sdts, nextSdt: { n: Object.keys(sdts).length }, blockSdtStack: [], inlineSdtStack: [], pendingBookmarks: [], currentBookmarks: null, pendingMarkers: [], currentMarkers: null, trackFields: false, fieldTrack: newFieldTrack() };
  const blocks: IRBlock[] = [];
  walkBlocks(children(root), blocks, ctx);
  return blocks;
}

const BOOKMARK_IGNORE = new Set(["_GoBack", "_Toc_Placeholder"]);
function bookmarkName(node: XmlNode): string | null {
  const name = attr(node, "w:name");
  return name && !BOOKMARK_IGNORE.has(name) ? name : null;
}

// ---------------------------------------------------------------------------
// Block-level walk

/** Alignment of a display equation from `m:oMathPara/m:oMathParaPr/m:jc`. */
function mathParaJc(node: XmlNode): "left" | "center" | "right" | undefined {
  const pr = el(node, "m:oMathParaPr");
  const jc = pr ? el(pr, "m:jc") : undefined;
  const v = jc ? attr(jc, "m:val") : undefined;
  if (v === "left" || v === "right") return v;
  if (v === "center" || v === "centerGroup") return "center";
  return undefined;
}

/** Uniform scale of a display equation from its paragraph's run font size
 *  (`w:p/w:pPr/w:rPr/w:sz`, half-points) relative to the base display size — the
 *  inverse of the export in documentXml.equationParagraphXml. Returns undefined
 *  when absent or ≈1 (the default size), so unresized equations carry no field. */
function mathParaScale(pNode: XmlNode): number | undefined {
  const pPr = el(pNode, "w:pPr");
  const rPr = pPr ? el(pPr, "w:rPr") : undefined;
  const sz = rPr ? el(rPr, "w:sz") : undefined;
  const hp = sz ? Number(attr(sz, "w:val")) : NaN;
  if (!Number.isFinite(hp) || hp <= 0) return undefined;
  const scale = halfPointsToPx(hp) / EQUATION_DISPLAY_PX;
  return Math.abs(scale - 1) < 0.02 ? undefined : scale;
}

function walkBlocks(nodes: XmlNode[], out: IRBlock[], ctx: ParseCtx): void {
  for (const node of nodes) {
    // A custom field open here marks the block being built; one that opens/closes
    // mid-block is captured during the parse (decideCustomField sets markBlock).
    if (ctx.trackFields) ctx.fieldTrack.markBlock = ctx.fieldTrack.openId;
    switch (node.tagName) {
      case "w:p": {
        // A display equation lives inside its own paragraph (<w:p><m:oMathPara/></w:p>)
        // — how Word stores it and how we export it (issue #193). Read that paragraph
        // straight to a display EquationBlock, carrying the same field/control metadata
        // as any other block, rather than a paragraph with a dropped math child.
        const mathPara = el(node, "m:oMathPara");
        if (mathPara) {
          const jc = mathParaJc(mathPara);
          const m: IRBlock = { kind: "math", root: ommlToMathml(mathPara), display: true };
          if (jc) m.align = jc;
          const scale = mathParaScale(node);
          if (scale !== undefined) m.scale = scale;
          if (ctx.trackFields && ctx.fieldTrack.markBlock) m.fieldId = ctx.fieldTrack.markBlock;
          if (ctx.blockSdtStack.length) m.sdtPath = [...ctx.blockSdtStack];
          out.push(m);
          break;
        }
        const p = parseParagraph(node, ctx);
        if (ctx.trackFields && ctx.fieldTrack.markBlock) p.fieldId = ctx.fieldTrack.markBlock;
        if (ctx.blockSdtStack.length) p.sdtPath = [...ctx.blockSdtStack];
        out.push(p);
        break;
      }
      case "w:tbl": {
        const t = parseTable(node, ctx);
        if (ctx.trackFields && ctx.fieldTrack.markBlock) t.fieldId = ctx.fieldTrack.markBlock;
        if (ctx.blockSdtStack.length) t.sdtPath = [...ctx.blockSdtStack];
        out.push(t);
        break;
      }
      case "w:sdt": {
        const content = el(node, "w:sdtContent");
        if (!content) break;
        const props = parseSdtPr(el(node, "w:sdtPr"));
        if (!props) {
          // No usable w:sdtPr — a transparent container; unwrap.
          ctx.warnings.add("sdt-unwrapped", "Content controls were unwrapped to their content.");
          walkBlocks(children(content), out, ctx);
          break;
        }
        // Block-level control: push its id so contained blocks carry it as part of
        // their sdtPath ancestry (supports nesting). Inner blocks are stamped at
        // their own w:p/w:tbl emit sites above.
        const sdtId = `sdt${ctx.nextSdt.n++}`;
        ctx.sdts[sdtId] = props;
        ctx.blockSdtStack.push(sdtId);
        walkBlocks(children(content), out, ctx);
        ctx.blockSdtStack.pop();
        break;
      }
      case "w:bookmarkStart": {
        // Block-level bookmark — attach to the start (offset 0) of the next paragraph.
        const name = bookmarkName(node);
        const idAttr = attr(node, "w:id");
        if (name) {
          ctx.pendingBookmarks.push(name);
          if (idAttr) ctx.pendingMarkers.push({ id: idAttr, name, kind: "start", offset: 0, inlineIndex: 0 });
        }
        break;
      }
      case "w:bookmarkEnd": {
        const idAttr = attr(node, "w:id");
        if (idAttr) ctx.pendingMarkers.push({ id: idAttr, kind: "end", offset: 0, inlineIndex: 0 });
        break;
      }
      case "m:oMathPara": {
        // Display (block) equation — convert OMML → MathML AST, carrying alignment
        // (m:oMathParaPr/m:jc) and the same field/control block metadata as w:p/w:tbl.
        const jc = mathParaJc(node);
        const m: IRBlock = { kind: "math", root: ommlToMathml(node), display: true };
        if (jc) m.align = jc;
        if (ctx.trackFields && ctx.fieldTrack.markBlock) m.fieldId = ctx.fieldTrack.markBlock;
        if (ctx.blockSdtStack.length) m.sdtPath = [...ctx.blockSdtStack];
        out.push(m);
        break;
      }
      case "w:sectPr":
        break; // handled by the caller
      default:
        break; // proofing marks, etc.
    }
  }
}

// ---------------------------------------------------------------------------
// Paragraphs and inline content

function parseParagraph(p: XmlNode, ctx: ParseCtx): IRParagraph {
  const pPr = el(p, "w:pPr");
  const props = pPr ? decodeParaProps(pPr, ctx.warnings) : {};
  const inlines: IRInline[] = [];
  // Bookmarks: block-level ones that preceded this paragraph, plus any
  // w:bookmarkStart found inside it (collected during walkInlines).
  const bookmarks = ctx.pendingBookmarks;
  ctx.pendingBookmarks = [];
  ctx.currentBookmarks = bookmarks;
  // Block-level markers (pending from before this paragraph) sit at its start;
  // inline markers accumulate during walkInlines.
  const markers = ctx.pendingMarkers;
  ctx.pendingMarkers = [];
  ctx.currentMarkers = markers;
  // Complex fields (w:fldChar begin → instr → separate → result → end) span
  // multiple runs; the state lives at paragraph scope.
  const field: FieldState = { depth: 0, instr: "", suppressResult: false };
  walkInlines(children(p), inlines, ctx, field);
  ctx.currentBookmarks = null;
  ctx.currentMarkers = null;
  const para: IRParagraph = { kind: "paragraph", props, inlines };
  if (bookmarks.length > 0) para.bookmarks = bookmarks;
  if (markers.length > 0) para.bookmarkMarkers = markers;
  if (field.tocInstr !== undefined) para.tocField = field.tocInstr;
  // w14:paraId/textId — Word's persistent paragraph ids, preserved verbatim.
  const paraId = attr(p, "w14:paraId");
  if (paraId) para.paraId = paraId;
  const textId = attr(p, "w14:textId");
  if (textId) para.textId = textId;
  return para;
}

interface FieldState {
  depth: number;
  instr: string;
  /** True between "separate" and "end" when the field became a token — the
   *  cached result runs must not ALSO be emitted. */
  suppressResult: boolean;
  /** Set between "separate" and "end" for a PAGEREF field: the bookmark it points
   *  at. Its cached result runs (the page number) are tagged with this in-document
   *  anchor so a TOC whose entries are plain text + a PAGEREF (no surrounding
   *  hyperlink) still maps each entry to its heading for "recalculate TOC". */
  pagerefAnchor?: string | undefined;
  /** The verbatim instrText of a `TOC` field in this paragraph (first one wins) —
   *  surfaced on IRParagraph.tocField so a headless render can anchor a built TOC. */
  tocInstr?: string | undefined;
  /** Set between "separate" and "end" for a built-in inline field whose cached
   *  result we keep (DATE/TIME/IF): the field id stamped on those result runs so
   *  they import as an editable field object. */
  resultFieldId?: string | undefined;
}

/** Built-in inline fields the editor models as field objects (PAGE/NUMPAGES emit a
 *  {page}/{pages} token; DATE/TIME/IF keep their materialized cached result). */
const BUILTIN_INLINE = new Set(["PAGE", "NUMPAGES", "DATE", "TIME", "IF"]);
/** Placeholder style for an IF spec's true/false runs parsed from the instruction
 *  (their text round-trips; the displayed result keeps its own resolved style). */
const FIELD_SPEC_STYLE: CharStyle = { fontFamily: "", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000000" };

/** The {page}/{pages} token (with format suffix) the layout substitutes per page. */
function pageTokenFor(spec: FieldSpec): string {
  if (spec.type === "NUMPAGES") return "{pages}";
  if (spec.type === "PAGE") return `{page${spec.numFmt && spec.numFmt !== "arabic" ? `:${spec.numFmt}` : ""}}`;
  return "";
}

/** " PAGE \\* MERGEFORMAT " → "{page}". Layout substitutes per page. */
function fieldToken(instr: string): string | undefined {
  if (/\bNUMPAGES\b/.test(instr)) return "{pages}"; // before PAGE — substring!
  if (/\bPAGE\b/.test(instr)) return "{page}";
  return undefined;
}

/** The in-document bookmark a TOC entry points at, from its field instruction:
 *  ` PAGEREF _Toc12345 \h ` or ` HYPERLINK \l "_Toc12345" ` → "_Toc12345". Word
 *  writes TOC entries as a HYPERLINK field wrapping the label + a nested PAGEREF
 *  for the number; either carries the same target. */
function anchorFromInstr(instr: string): string | undefined {
  const m = instr.match(/\bPAGEREF\s+("[^"]+"|\S+)/) ?? instr.match(/\bHYPERLINK\s+\\l\s+("[^"]+"|\S+)/);
  return m ? m[1]!.replace(/^"|"$/g, "") : undefined;
}

function walkInlines(nodes: XmlNode[], out: IRInline[], ctx: ParseCtx, field: FieldState): void {
  for (const node of nodes) {
    switch (node.tagName) {
      case "w:r":
        parseRun(node, out, ctx, field);
        break;
      case "w:hyperlink": {
        // Tag the contained runs with the link target — r:id is an external
        // rel (URL, resolved in mapToModel), w:anchor an in-document bookmark.
        const relId = attr(node, "r:id");
        const anchor = attr(node, "w:anchor");
        const start = out.length;
        walkInlines(children(node), out, ctx, field);
        if (relId || anchor) {
          for (let i = start; i < out.length; i++) {
            const inline = out[i]!;
            // Runs and inline equations both carry link metadata (math under
            // w:hyperlink is valid OOXML); other inline kinds don't.
            const props = inline.kind === "run" ? inline.props : inline.kind === "mathInline" ? (inline.props ??= {}) : undefined;
            if (!props) continue;
            if (relId) props.linkRelId = relId;
            else if (anchor) props.linkAnchor = anchor;
          }
        }
        break;
      }
      case "w:sdt": {
        const content = el(node, "w:sdtContent");
        if (!content) break;
        const props = parseSdtPr(el(node, "w:sdtPr"));
        if (!props) {
          ctx.warnings.add("sdt-unwrapped", "Content controls were unwrapped to their content.");
          walkInlines(children(content), out, ctx, field);
          break;
        }
        // Inline control: push its id so runs created inside (via parseRun) carry
        // it on their sdtPath ancestry — recursion handles inline-in-inline nesting.
        const sdtId = `sdt${ctx.nextSdt.n++}`;
        ctx.sdts[sdtId] = props;
        ctx.inlineSdtStack.push(sdtId);
        walkInlines(children(content), out, ctx, field);
        ctx.inlineSdtStack.pop();
        break;
      }
      case "w:fldSimple": {
        const instr = attr(node, "w:instr") ?? "";
        const token = ctx.fieldTokens ? fieldToken(instr) : undefined;
        if (token) {
          out.push({ kind: "run", text: token, props: firstRunProps(node) });
        } else {
          walkInlines(children(node), out, ctx, field); // cached result runs
        }
        break;
      }
      case "w:smartTag":
      case "w:ins": // accepted tracked insertion
        walkInlines(children(node), out, ctx, field);
        break;
      case "w:del": // tracked deletion — content is already "deleted"
        break;
      case "w:bookmarkStart": {
        const name = bookmarkName(node);
        const idAttr = attr(node, "w:id");
        if (name && ctx.currentBookmarks) ctx.currentBookmarks.push(name);
        if (name && idAttr && ctx.currentMarkers) {
          ctx.currentMarkers.push({ id: idAttr, name, kind: "start", offset: inlineOffset(out), inlineIndex: out.length });
        }
        break;
      }
      case "w:bookmarkEnd": {
        const idAttr = attr(node, "w:id");
        if (idAttr && ctx.currentMarkers) ctx.currentMarkers.push({ id: idAttr, kind: "end", offset: inlineOffset(out), inlineIndex: out.length });
        break;
      }
      case "m:oMath": {
        // Inline equation → a single-U+FFFC run carrying the MathML AST.
        out.push({ kind: "mathInline", root: ommlToMathml(node) });
        break;
      }
      default:
        break; // w:pPr, w:proofErr, comment ranges, …
    }
  }
}

/** Props of the first run inside a container — used to style synthesized
 *  field-token runs like the field's visible result. */
function firstRunProps(container: XmlNode): IRRunProps {
  const run = findDeep(container, "w:r");
  const rPr = run && el(run, "w:rPr");
  return rPr ? decodeRunProps(rPr) : {};
}

/** A single w:r can hold many content children (w:t, w:tab, w:br, …).
 *  Text accumulates into one IR run; breaks flush and emit a break marker. */
function parseRun(r: XmlNode, out: IRInline[], ctx: ParseCtx, field: FieldState): void {
  const rPr = el(r, "w:rPr");
  const props = rPr ? decodeRunProps(rPr) : {};
  // Snapshot the open inline-control ancestry onto every run this w:r produces.
  const startIdx = out.length;
  const sdtPath = ctx.inlineSdtStack.length ? [...ctx.inlineSdtStack] : undefined;
  let text = "";
  const flush = (): void => {
    if (text.length > 0) {
      // Tag a PAGEREF's cached result (the page number) with its in-document
      // anchor so TOC entries lacking a surrounding hyperlink still map to their
      // heading. A real hyperlink wrapper overrides this afterwards (same target).
      if (field.pagerefAnchor && props.linkAnchor === undefined && props.linkRelId === undefined) {
        props.linkAnchor = field.pagerefAnchor;
      }
      out.push({ kind: "run", text, props, ...(field.resultFieldId ? { fieldId: field.resultFieldId } : {}) });
    }
    text = "";
  };
  const walkContent = (nodes: XmlNode[]): void => {
    for (const node of nodes) {
      switch (node.tagName) {
        case "w:t":
          if (!field.suppressResult) text += textOf(node);
          break;
        case "w:tab":
          if (!field.suppressResult) text += "\t"; // policy (fixed spaces) applied in mapToModel
          break;
        case "w:br": {
          flush();
          const brType = attr(node, "w:type");
          if (brType === "page") out.push({ kind: "break", page: true });
          else if (brType === "column") out.push({ kind: "break", column: true });
          else out.push({ kind: "break" });
          break;
        }
        case "w:cr":
          flush();
          out.push({ kind: "break" });
          break;
        case "w:noBreakHyphen":
          text += "‑";
          break;
        case "w:sym": {
          // A symbol-font glyph (font + hex code point). Flush pending text, then
          // emit a standalone run carrying the symbol marker; its text is the
          // decoded glyph so layout/paint render it in the symbol font.
          const font = attr(node, "w:font");
          const charHex = attr(node, "w:char");
          if (font && charHex) {
            flush();
            out.push({
              kind: "run",
              text: symbolGlyph(charHex),
              props: { ...props, symbol: { font, char: charHex.toUpperCase() } },
              ...(field.resultFieldId ? { fieldId: field.resultFieldId } : {}),
            });
          }
          break;
        }
        case "w:drawing": {
          flush();
          const image = parseDrawing(node, ctx);
          if (image) out.push(image);
          break;
        }
        case "w:pict": {
          flush();
          const image = parseVmlPict(node, ctx);
          if (image) out.push(image);
          break;
        }
        case "w:object":
          ctx.warnings.add("objects-skipped", "Embedded OLE objects are not imported.");
          break;
        case "w:instrText":
          if (field.depth > 0) {
            field.instr += textOf(node);
            // A TOC entry's HYPERLINK/PAGEREF target — remember it so the cached
            // result runs (label + page number) get tagged as in-document links,
            // even when nested (Word nests PAGEREF inside the entry's HYPERLINK).
            const anchor = anchorFromInstr(field.instr);
            if (anchor) field.pagerefAnchor = anchor;
            // The paragraph's own TOC field instruction (first wins) — drives the
            // headless TOC build's anchor + the re-emitted instruction on export.
            const t = textOf(node);
            if (field.tocInstr === undefined && /\bTOC\b/.test(t)) field.tocInstr = t;
          }
          // Block-scoped custom-field tracker: accumulate the TOP-level field's
          // instruction (nested fields' instrText belongs to them, not the field
          // whose result region we're capturing).
          if (ctx.trackFields && ctx.fieldTrack.depth === 1) ctx.fieldTrack.topInstr += textOf(node);
          break;
        case "w:fldChar":
          handleFldChar(node, props, out, ctx, field);
          break;
        case "mc:AlternateContent": {
          // Word wraps newer markup (drawings, groups) in AlternateContent:
          // mc:Choice carries the real content, mc:Fallback a legacy stand-in.
          const choice = el(node, "mc:Choice") ?? el(node, "mc:Fallback");
          if (choice) walkContent(children(choice));
          break;
        }
        case "w:footnoteReference": {
          // The marker run: text becomes the note number in mapToModel.
          const fnId = attr(node, "w:id");
          if (fnId !== undefined) {
            flush();
            out.push({ kind: "run", text: "", props: { ...props, footnoteId: fnId } });
          }
          break;
        }
        case "w:endnoteReference": {
          // The marker run: text becomes the note number in mapToModel.
          const enId = attr(node, "w:id");
          if (enId !== undefined) {
            flush();
            out.push({ kind: "run", text: "", props: { ...props, endnoteId: enId } });
          }
          break;
        }
        case "w:footnoteRef":
        case "w:endnoteRef":
          break; // the auto-number placeholder inside a note BODY — engine paints it
        default:
          break; // w:rPr, w:lastRenderedPageBreak, …
      }
    }
  };
  walkContent(children(r));
  flush();
  if (sdtPath) {
    for (let i = startIdx; i < out.length; i++) {
      const inl = out[i]!;
      if (inl.kind === "run") inl.sdtPath = sdtPath;
    }
  }
}

function handleFldChar(
  node: XmlNode,
  runProps: IRRunProps,
  out: IRInline[],
  ctx: ParseCtx,
  field: FieldState,
): void {
  const ft = ctx.fieldTrack;
  switch (attr(node, "w:fldCharType")) {
    case "begin":
      field.depth++;
      if (field.depth === 1) {
        field.instr = "";
        field.suppressResult = false;
        field.pagerefAnchor = undefined;
        field.resultFieldId = undefined;
      }
      if (ctx.trackFields) {
        ft.depth++;
        if (ft.depth === 1) {
          ft.topInstr = "";
          ft.openId = null;
        }
      }
      break;
    case "separate": {
      // Custom field's instruction is complete by `separate` — open it now so its
      // result blocks (which follow) get tagged with the field id.
      if (ctx.trackFields && ft.depth === 1) decideCustomField(ctx);
      if (field.depth !== 1) break;
      // Built-in inline field → a field object (body only; bands keep the token
      // path below for back-compat). PAGE/NUMPAGES emit a tagged {page}/{pages}
      // token; DATE/TIME/IF keep their cached result runs, tagged via resultFieldId.
      if (ctx.trackFields) {
        const parsed = parseFieldInstruction(field.instr);
        if (BUILTIN_INLINE.has(parsed.name)) {
          const spec = parseFieldSpec(parsed, FIELD_SPEC_STYLE);
          if (spec) {
            const fid = `field${ft.next.n++}`;
            ft.registry[fid] = { id: fid, instruction: field.instr, name: parsed.name, kind: "builtin", spec };
            if (spec.type === "PAGE" || spec.type === "NUMPAGES") {
              out.push({ kind: "run", text: pageTokenFor(spec), props: runProps, fieldId: fid });
              field.suppressResult = true;
            } else {
              field.resultFieldId = fid;
            }
            break;
          }
        }
      }
      const token = ctx.fieldTokens ? fieldToken(field.instr) : undefined;
      if (token) {
        out.push({ kind: "run", text: token, props: runProps });
        field.suppressResult = true; // the cached result would duplicate the token
      }
      break;
    }
    case "end":
      if (ctx.trackFields && ft.depth > 0) {
        if (ft.depth === 1) {
          decideCustomField(ctx); // custom field with no `separate` (no cached result)
          ft.openId = null; // closed — but markBlock keeps this block tagged
          ft.topInstr = "";
        }
        ft.depth--;
      }
      if (field.depth > 0) field.depth--;
      if (field.depth === 0) {
        // Field without a "separate" (no cached result): emit the token now.
        if (!field.suppressResult) {
          const token = ctx.fieldTokens ? fieldToken(field.instr) : undefined;
          if (token) out.push({ kind: "run", text: token, props: runProps });
        }
        field.instr = "";
        field.suppressResult = false;
        field.pagerefAnchor = undefined;
        field.resultFieldId = undefined;
      }
      break;
  }
}

// ---------------------------------------------------------------------------
// Images

/** Decode a w:sym/@w:char hex code point to its glyph. Word stores symbol-font
 *  glyphs in the Private-Use range (e.g. "F0E0"); we render that code point in the
 *  symbol font. Falls back to a replacement char for an unparseable code. */
function symbolGlyph(charHex: string): string {
  const cp = parseInt(charHex, 16);
  // Guard the full valid range — String.fromCodePoint throws RangeError above
  // 0x10FFFF, so a malformed w:char must fall back instead of aborting the import.
  return Number.isFinite(cp) && cp > 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : "�";
}

/** wp:anchor wrap + float extraction shared by the image and shape drawing paths.
 *  Square/tight/through wrap → the model's "square" float; wrapNone (behind/in-front
 *  of text) → an absolutely-positioned anchor; only topAndBottom/none-without-wrapNone
 *  fall back to block flow (`unsupported`). */
interface ParsedAnchor {
  wrap: "square" | "block";
  align?: "left" | "right" | "center";
  float?: NonNullable<Extract<IRInline, { kind: "image" }>["anchorFloat"]>;
  /** True for an anchor whose wrap mode the model can't express (caller may warn). */
  unsupported: boolean;
}
function parseAnchorProps(anchor: XmlNode): ParsedAnchor {
  const square = el(anchor, "wp:wrapSquare") ?? el(anchor, "wp:wrapTight") ?? el(anchor, "wp:wrapThrough");
  const wrapNone = el(anchor, "wp:wrapNone");
  const posH = el(anchor, "wp:positionH");
  const posV = el(anchor, "wp:positionV");
  const alignText = posH && el(posH, "wp:align");
  const alignVal = alignText && textOf(alignText);
  const result: ParsedAnchor = { wrap: square ? "square" : "block", unsupported: !wrapNone && !square };
  if (alignVal === "left" || alignVal === "right" || alignVal === "center") result.align = alignVal;
  if (wrapNone) {
    const dec = findDeep(anchor, "adec:decorative");
    const z = numAttr(anchor, "relativeHeight");
    result.float = {
      behind: attr(anchor, "behindDoc") === "1",
      offsetXEmu: posOffsetEmu(posH),
      offsetYEmu: posOffsetEmu(posV),
      relFromH: relFromH(posH),
      relFromV: relFromV(posV),
      decorative: !!dec && attr(dec, "val") !== "0",
      ...(z !== undefined ? { z } : {}),
    };
  }
  return result;
}

/** DrawingML: w:drawing → wp:inline|wp:anchor → … → a:blip r:embed. The exact
 *  nesting varies by producer, so the blip is found by deep search. */
function parseDrawing(drawing: XmlNode, ctx: ParseCtx): IRInline | undefined {
  const anchor = el(drawing, "wp:anchor");
  const container = el(drawing, "wp:inline") ?? anchor;
  if (!container) return undefined;
  // A grouped-shapes container (…/wordprocessingGroup → wpg:wgp) — a wps:wsp lives
  // INSIDE it, so this must be checked BEFORE the plain-shape branch below or the
  // group's first child would be mistaken for a standalone shape.
  const wgp = findDeep(container, "wpg:wgp");
  if (wgp) {
    const group = parseShapeGroupDrawing(wgp, container, anchor, ctx);
    if (group) return group;
    ctx.warnings.add("shape-skipped", "A grouped drawing without any supported member shapes was skipped.");
    return undefined;
  }
  // A DrawingML preset shape (…/wordprocessingShape → wps:wsp) has no a:blip. Parse
  // it before the image path so the images-skipped warning stays narrowed to
  // genuinely unknown drawings.
  const wsp = findDeep(container, "wps:wsp");
  if (wsp) {
    const shape = parseShapeWsp(wsp, container, anchor, ctx);
    if (shape) return shape;
    // Structurally a wps shape but with no geometry we can place (neither a
    // a:prstGeom preset nor a a:custGeom path) — warn instead of dropping it
    // silently, mirroring the image branch's images-skipped warning.
    ctx.warnings.add("shape-skipped", "A drawing shape without a supported geometry (no preset or custom path) was skipped.");
    return undefined;
  }
  const blip = findDeep(container, "a:blip");
  // r:embed = image bytes packaged inside the docx; r:link = a "Link to File"
  // image whose bytes live outside the package (an http(s) URL or local path,
  // via a TargetMode="External" relationship). Accept either — the media store
  // resolves an external rel to its URL.
  const relId = blip && (attr(blip, "r:embed") ?? attr(blip, "r:link"));
  if (!relId) {
    ctx.warnings.add("images-skipped", "A drawing without an embedded image reference was skipped.");
    return undefined;
  }
  const extent = el(container, "wp:extent");
  const image: IRInline = { kind: "image", relId, anchored: !!anchor };
  // wp14:anchorId/editId — Word's persistent drawing ids, on the inline|anchor container.
  const anchorId = attr(container, "wp14:anchorId");
  if (anchorId) image.anchorId = anchorId;
  const editId = attr(container, "wp14:editId");
  if (editId) image.editId = editId;
  const cx = numAttr(extent, "cx");
  if (cx !== undefined) image.widthEmu = cx;
  const cy = numAttr(extent, "cy");
  if (cy !== undefined) image.heightEmu = cy;
  // a:srcRect — crop insets in 1/1000 of a percent (so 10% = 10000). Normalize to
  // a 0..1 fraction per edge; skip a degenerate rect (would crop the whole image).
  const srcRect = findDeep(container, "a:srcRect");
  if (srcRect) {
    const inset = (name: string): number => {
      const v = numAttr(srcRect, name);
      return v !== undefined && Number.isFinite(v) ? v / 100000 : 0;
    };
    const crop = { left: inset("l"), top: inset("t"), right: inset("r"), bottom: inset("b") };
    if ((crop.left || crop.top || crop.right || crop.bottom) && crop.left + crop.right < 1 && crop.top + crop.bottom < 1) {
      image.crop = crop;
    }
  }
  if (anchor) {
    const ap = parseAnchorProps(anchor);
    image.anchorWrap = ap.wrap;
    if (ap.align) image.anchorAlign = ap.align;
    if (ap.float) image.anchorFloat = ap.float;
    else if (ap.unsupported) {
      ctx.warnings.add(
        "images-anchored",
        "Some floating images (overlapping or top-and-bottom wrap) were placed in the text flow.",
      );
    }
  }
  return image;
}

/** a:custGeom → a:pathLst/a:path into a normalized (0–1) ShapePath. Point coords are
 *  divided by the a:path @w/@h design space so the model path is box-independent
 *  (mirrors the exporter's fixed-unit emission). The model only carries move/line/
 *  cubic/close, so unmodeled commands (a:arcTo / a:quadBezTo), malformed segments,
 *  and any a:path beyond the first are dropped — but never silently: a
 *  `custom-path-simplified` warning fires per the importer's warn-don't-drop contract.
 *  Returns undefined when there is no usable path. */
function parseCustGeom(custGeom: XmlNode, ctx: ParseCtx): ShapePath | undefined {
  const pathLst = el(custGeom, "a:pathLst");
  const paths = pathLst ? els(pathLst, "a:path") : [];
  const path = paths[0];
  if (!path) return undefined;
  // Extra sub-paths (a:pathLst with >1 a:path) collapse to the first — flag the loss.
  let simplified = paths.length > 1;
  const w = numAttr(path, "w") ?? 0;
  const h = numAttr(path, "h") ?? 0;
  if (w <= 0 || h <= 0) return undefined; // no design space to normalize against
  const ptOf = (node: XmlNode): { x: number; y: number } => ({
    x: (numAttr(node, "x") ?? 0) / w,
    y: (numAttr(node, "y") ?? 0) / h,
  });
  const segments: ShapePathSegment[] = [];
  for (const cmd of children(path)) {
    switch (cmd.tagName) {
      case "a:moveTo": {
        const p = el(cmd, "a:pt");
        if (p) segments.push({ type: "moveTo", ...ptOf(p) });
        else simplified = true;
        break;
      }
      case "a:lnTo": {
        const p = el(cmd, "a:pt");
        if (p) segments.push({ type: "lineTo", ...ptOf(p) });
        else simplified = true;
        break;
      }
      case "a:cubicBezTo": {
        const pts = els(cmd, "a:pt");
        if (pts.length === 3) {
          const c1 = ptOf(pts[0]!), c2 = ptOf(pts[1]!), end = ptOf(pts[2]!);
          segments.push({ type: "cubicBezierTo", x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y, x: end.x, y: end.y });
        } else {
          simplified = true; // a cubic needs exactly 3 control/end points
        }
        break;
      }
      case "a:close":
        segments.push({ type: "close" });
        break;
      default:
        // a:arcTo / a:quadBezTo or any other command the model can't represent.
        simplified = true;
    }
  }
  if (simplified) {
    ctx.warnings.add(
      "custom-path-simplified",
      "A freeform shape's custom geometry used path commands we don't model (arcs, quadratic Béziers, malformed segments, or multiple sub-paths) — they were simplified out.",
    );
  }
  return segments.length > 0 ? { segments } : undefined;
}

/** The geometry/style props shared by a leaf shape and a group member: a:prstGeom
 *  (preset + a:avLst adjust) or a:custGeom (freeform path), a:xfrm@rot rotation,
 *  spPr's direct fill, the a:ln outline (+ a:prstDash), and the wps:txbx text body.
 *  Size and anchor/position are the caller's concern (they differ for a top-level
 *  shape vs a group child). Returns undefined when there's no geometry at all. */
interface WspCore {
  preset: string;
  adjust?: Record<string, number>;
  custom?: ShapePath;
  rotationDeg?: number;
  fill?: { color: string } | { none: true };
  stroke?: { color: string; widthPt: number; dash?: string } | { none: true };
  text?: IRBlock[];
}
function parseWspCore(wsp: XmlNode, ctx: ParseCtx): WspCore | undefined {
  const spPr = el(wsp, "wps:spPr");
  const prstGeom = spPr && el(spPr, "a:prstGeom");
  const prst = prstGeom && attr(prstGeom, "prst");
  const custGeom = spPr && el(spPr, "a:custGeom");
  if (!prst && !custGeom) return undefined; // no geometry → not a shape we can place
  // A custom geometry has no preset; keep a "rect" fallback so the box still maps.
  const core: WspCore = { preset: prst ?? "rect" };
  // a:custGeom → a:pathLst/a:path: the freeform path, normalized to 0–1 fractions of
  // the box (dividing point coords by the a:path @w/@h design space). Mirrors #220.
  if (custGeom) {
    const custom = parseCustGeom(custGeom, ctx);
    if (custom) core.custom = custom;
  }
  // a:avLst → a:gd @name/@fmla="val N": the parametric adjust handles, kept raw.
  const avLst = prstGeom && el(prstGeom, "a:avLst");
  if (avLst) {
    const adjust: Record<string, number> = {};
    for (const gd of els(avLst, "a:gd")) {
      const name = attr(gd, "name");
      const fmla = attr(gd, "fmla");
      const m = fmla?.match(/^val\s+(-?\d+)$/);
      if (name && m) adjust[name] = Number(m[1]);
    }
    if (Object.keys(adjust).length > 0) core.adjust = adjust;
  }
  // a:xfrm@rot — clockwise rotation in 60000ths of a degree → degrees.
  const rot = numAttr(spPr && el(spPr, "a:xfrm"), "rot");
  if (rot !== undefined && rot !== 0) core.rotationDeg = rot / 60000;
  const srgbVal = (parent: XmlNode | undefined): string | undefined => {
    const solid = parent && el(parent, "a:solidFill");
    const clr = solid && el(solid, "a:srgbClr");
    return clr ? attr(clr, "val") : undefined;
  };
  if (spPr) {
    // Fill: spPr's DIRECT solidFill/noFill (never descend into a:ln's fill).
    const fillVal = srgbVal(spPr);
    if (fillVal) core.fill = { color: `#${fillVal}` };
    else if (el(spPr, "a:noFill")) core.fill = { none: true };
    // Outline: a:ln → noFill (no outline) or solidFill (color) with @w (EMU) width
    // and an optional a:prstDash@val dash preset.
    const ln = el(spPr, "a:ln");
    if (ln) {
      if (el(ln, "a:noFill")) {
        core.stroke = { none: true };
      } else {
        const lnVal = srgbVal(ln);
        if (lnVal) {
          const w = numAttr(ln, "w");
          const stroke: { color: string; widthPt: number; dash?: string } = { color: `#${lnVal}`, widthPt: w !== undefined ? w / 12700 : 1 };
          const prstDash = el(ln, "a:prstDash");
          const dash = prstDash && attr(prstDash, "val");
          if (dash) stroke.dash = dash;
          core.stroke = stroke;
        }
      }
    }
  }
  // Text box body: wps:txbx → w:txbxContent holds a block flow (paragraphs).
  // Parse it as a fresh block container (like a table cell) so any surrounding
  // field/control tracking never leaks into the nested paragraphs.
  const txbxContent = el(wsp, "wps:txbx") && el(el(wsp, "wps:txbx")!, "w:txbxContent");
  if (txbxContent) {
    const savedBlockStack = ctx.blockSdtStack;
    const savedTrackFields = ctx.trackFields;
    ctx.blockSdtStack = [];
    ctx.trackFields = false;
    const blocks: IRBlock[] = [];
    walkBlocks(children(txbxContent), blocks, ctx);
    ctx.blockSdtStack = savedBlockStack;
    ctx.trackFields = savedTrackFields;
    if (blocks.length > 0) core.text = blocks;
  }
  return core;
}

/** DrawingML preset shape: wps:wsp → wps:spPr (a:prstGeom, a:solidFill/a:noFill,
 *  a:ln). Size comes from the container's wp:extent (like an image). The spPr's
 *  DIRECT fill is the shape fill; a:ln carries the outline (its own nested fill). */
function parseShapeWsp(
  wsp: XmlNode,
  container: XmlNode,
  anchor: XmlNode | undefined,
  ctx: ParseCtx,
): Extract<IRInline, { kind: "shape" }> | undefined {
  const core = parseWspCore(wsp, ctx);
  if (!core) return undefined;
  const shape: Extract<IRInline, { kind: "shape" }> = { kind: "shape", ...core };
  const extent = el(container, "wp:extent");
  const cx = numAttr(extent, "cx");
  if (cx !== undefined) shape.widthEmu = cx;
  const cy = numAttr(extent, "cy");
  if (cy !== undefined) shape.heightEmu = cy;
  if (anchor) applyShapeAnchor(shape, anchor, ctx);
  applyDrawingIds(shape, container);
  return shape;
}

/** wps:wsp / wpg:grpSp @a:xfrm a:off/a:ext — a group member's local rect in the
 *  child coordinate space (EMU). Absent → 0. */
function xfrmRect(spPr: XmlNode | undefined): { xEmu: number; yEmu: number; cxEmu: number; cyEmu: number } {
  const xfrm = spPr && el(spPr, "a:xfrm");
  const off = xfrm && el(xfrm, "a:off");
  const ext = xfrm && el(xfrm, "a:ext");
  return {
    xEmu: numAttr(off, "x") ?? 0,
    yEmu: numAttr(off, "y") ?? 0,
    cxEmu: numAttr(ext, "cx") ?? 0,
    cyEmu: numAttr(ext, "cy") ?? 0,
  };
}

/** wpg:grpSpPr / a:xfrm — a group's own extent + child coordinate space
 *  (a:chOff / a:chExt), all in EMU. */
function parseGroupXfrm(wgp: XmlNode): { off: { x: number; y: number }; ext: { cx: number; cy: number }; chOff: { x: number; y: number }; chExt: { cx: number; cy: number }; rotDeg?: number } {
  const grpSpPr = el(wgp, "wpg:grpSpPr");
  const xfrm = grpSpPr && el(grpSpPr, "a:xfrm");
  const off = xfrm && el(xfrm, "a:off");
  const ext = xfrm && el(xfrm, "a:ext");
  const chOff = xfrm && el(xfrm, "a:chOff");
  const chExt = xfrm && el(xfrm, "a:chExt");
  const rot = numAttr(xfrm, "rot");
  return {
    off: { x: numAttr(off, "x") ?? 0, y: numAttr(off, "y") ?? 0 },
    ext: { cx: numAttr(ext, "cx") ?? 0, cy: numAttr(ext, "cy") ?? 0 },
    chOff: { x: numAttr(chOff, "x") ?? 0, y: numAttr(chOff, "y") ?? 0 },
    chExt: { cx: numAttr(chExt, "cx") ?? 0, cy: numAttr(chExt, "cy") ?? 0 },
    ...(rot !== undefined && rot !== 0 ? { rotDeg: rot / 60000 } : {}),
  };
}

/** A grouped-shapes container (wpg:wgp / wpg:grpSp): its child coordinate space and
 *  its member drawings (leaf wps:wsp shapes and nested wpg:grpSp groups). Returns
 *  undefined when no member has a supported geometry. */
function parseShapeGroupData(wgp: XmlNode, ctx: ParseCtx): IRShapeGroup | undefined {
  const xf = parseGroupXfrm(wgp);
  const kids: IRShapeChild[] = [];
  for (const node of children(wgp)) {
    if (node.tagName === "wps:wsp") {
      const core = parseWspCore(node, ctx);
      if (!core) continue;
      const rect = xfrmRect(el(node, "wps:spPr"));
      kids.push({ xEmu: rect.xEmu, yEmu: rect.yEmu, widthEmu: rect.cxEmu, heightEmu: rect.cyEmu, ...core });
    } else if (node.tagName === "wpg:grpSp") {
      const nested = parseShapeGroupData(node, ctx);
      if (!nested) continue;
      // A nested group's own a:xfrm carries its rect (a:off/a:ext) AND its rotation
      // (a:xfrm@rot) — read both so a rotated nested group round-trips.
      const gx = parseGroupXfrm(node);
      // A nested group is a member shape whose geometry is the group; carry a rect +
      // a placeholder preset so mapShape treats it as a (nested) group container.
      kids.push({
        xEmu: gx.off.x, yEmu: gx.off.y, widthEmu: gx.ext.cx, heightEmu: gx.ext.cy,
        preset: "rect", group: nested,
        ...(gx.rotDeg !== undefined ? { rotationDeg: gx.rotDeg } : {}),
      });
    }
  }
  if (kids.length === 0) return undefined;
  return {
    childOffXEmu: xf.chOff.x,
    childOffYEmu: xf.chOff.y,
    childExtXEmu: xf.chExt.cx,
    childExtYEmu: xf.chExt.cy,
    children: kids,
  };
}

/** Top-level grouped shapes: wp:inline|wp:anchor → …/wordprocessingGroup → wpg:wgp.
 *  Becomes a block-level ShapeBlock whose `group` carries the members. Size comes
 *  from the container's wp:extent (like a shape/image); anchor + wp14 ids mirror the
 *  leaf-shape path. */
function parseShapeGroupDrawing(
  wgp: XmlNode,
  container: XmlNode,
  anchor: XmlNode | undefined,
  ctx: ParseCtx,
): Extract<IRInline, { kind: "shape" }> | undefined {
  const group = parseShapeGroupData(wgp, ctx);
  if (!group) return undefined;
  // The container is a rect box; its preset is a placeholder (never painted — the
  // paint layer draws the children when `group` is present).
  const shape: Extract<IRInline, { kind: "shape" }> = { kind: "shape", preset: "rect", group };
  const extent = el(container, "wp:extent");
  const cx = numAttr(extent, "cx");
  if (cx !== undefined) shape.widthEmu = cx;
  const cy = numAttr(extent, "cy");
  if (cy !== undefined) shape.heightEmu = cy;
  // The group's a:xfrm@rot (wpg:grpSpPr) — a rotated group must round-trip its
  // rotation onto the container (mapped to ShapeBlock.rotation, re-emitted on export).
  const rotDeg = parseGroupXfrm(wgp).rotDeg;
  if (rotDeg !== undefined) shape.rotationDeg = rotDeg;
  if (anchor) applyShapeAnchor(shape, anchor, ctx);
  applyDrawingIds(shape, container);
  return shape;
}

/** wp:anchor → the shape's wrap/float props (shared by leaf shapes and groups). */
function applyShapeAnchor(shape: Extract<IRInline, { kind: "shape" }>, anchor: XmlNode, ctx: ParseCtx): void {
  shape.anchored = true;
  const ap = parseAnchorProps(anchor);
  shape.anchorWrap = ap.wrap;
  if (ap.align) shape.anchorAlign = ap.align;
  if (ap.float) shape.anchorFloat = ap.float;
  else if (ap.unsupported) {
    ctx.warnings.add(
      "shapes-anchored",
      "Some floating shapes (overlapping or top-and-bottom wrap) were placed in the text flow.",
    );
  }
}

/** wp14:anchorId / wp14:editId on the wp:inline|wp:anchor container — preserved verbatim. */
function applyDrawingIds(shape: Extract<IRInline, { kind: "shape" }>, container: XmlNode): void {
  const anchorId = attr(container, "wp14:anchorId");
  if (anchorId) shape.anchorId = anchorId;
  const editId = attr(container, "wp14:editId");
  if (editId) shape.editId = editId;
}

/** wp:positionH|V → wp:posOffset (EMU, signed). Absent/non-numeric → 0.
 *  Word wraps wp14 *percentage* positioning in `mc:AlternateContent`, which pushes
 *  the absolute `wp:posOffset` down into `mc:Fallback` — no longer a direct child.
 *  We don't model the percentage (the mc:Choice), so recover the fallback absolute
 *  by descending; within a `wp:positionH|V` the only nested posOffset IS the
 *  fallback, so the deep search is unambiguous. Without this the offset read as 0
 *  and a percent-positioned float snapped to the top-left of its reference. */
function posOffsetEmu(pos: XmlNode | undefined): number {
  if (!pos) return 0;
  const off = el(pos, "wp:posOffset") ?? findDeep(pos, "wp:posOffset");
  const n = off ? Number(textOf(off).trim()) : NaN;
  return Number.isFinite(n) ? n : 0;
}

const REL_FROM_H = new Set(["page", "margin", "column", "leftMargin", "rightMargin", "character"]);
const REL_FROM_V = new Set(["page", "margin", "paragraph", "line", "topMargin", "bottomMargin"]);

/** wp:positionH/@relativeFrom — OOXML default is "column". Unknown → "column". */
function relFromH(pos: XmlNode | undefined): NonNullable<Extract<IRInline, { kind: "image" }>["anchorFloat"]>["relFromH"] {
  const v = pos && attr(pos, "relativeFrom");
  return v && REL_FROM_H.has(v) ? (v as never) : "column";
}

/** wp:positionV/@relativeFrom — OOXML default is "paragraph". Unknown → "paragraph". */
function relFromV(pos: XmlNode | undefined): NonNullable<Extract<IRInline, { kind: "image" }>["anchorFloat"]>["relFromV"] {
  const v = pos && attr(pos, "relativeFrom");
  return v && REL_FROM_V.has(v) ? (v as never) : "paragraph";
}

/** Legacy VML (w:pict), still produced for some content and the only shape markup
 *  in pre-DrawingML (Word 2003-era) documents. Two flavours are recognized:
 *   • a picture — v:shape → v:imagedata r:id (checked first; unchanged behaviour);
 *   • a drawn shape / text box — v:rect|v:roundrect|v:oval|v:line|v:shape, mapped
 *     read-only into the SAME `kind:"shape"` IR the DrawingML wps path emits
 *     (export always re-emits modern DrawingML — see parseVmlShape). */
function parseVmlPict(pict: XmlNode, ctx: ParseCtx): IRInline | undefined {
  const imagedata = findDeep(pict, "v:imagedata");
  const relId = imagedata && attr(imagedata, "r:id");
  if (relId) {
    const image: IRInline = { kind: "image", relId, anchored: false };
    const shape = findDeep(pict, "v:shape");
    const w = vmlLengthToEmu(vmlStyle(shape)["width"]);
    if (w !== undefined) image.widthEmu = w;
    const h = vmlLengthToEmu(vmlStyle(shape)["height"]);
    if (h !== undefined) image.heightEmu = h;
    return image;
  }
  const shape = parseVmlShape(pict, ctx);
  if (shape) return shape;
  // Neither a picture nor a recognized drawn shape — the pict carries content we
  // don't model (WordArt, a group we can't place, an OLE preview, …). Warn once
  // rather than dropping it silently.
  ctx.warnings.add("pict-skipped", "A legacy VML drawing that isn't a picture or a recognized shape was skipped.");
  return undefined;
}

// ---------------------------------------------------------------------------
// Legacy VML drawn shapes (import-only; export always emits DrawingML wps)

/** Parse a VML style="k:v;…" declaration list into a lookup (keys lower-cased). */
function vmlStyle(node: XmlNode | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  for (const decl of (node ? attr(node, "style") ?? "" : "").split(";")) {
    const i = decl.indexOf(":");
    if (i < 0) continue;
    const k = decl.slice(0, i).trim().toLowerCase();
    if (k) out[k] = decl.slice(i + 1).trim();
  }
  return out;
}

const EMU_PER = { pt: 12700, px: 9525, in: 914400, cm: 360000, mm: 36000, pc: 152400 } as const;

/** A VML/CSS length ("72pt", "1in", "96px", ".5cm", bare number) → EMU. Bare
 *  numbers default to points — the unit Word emits for w:pict style width/height. */
function vmlLengthToEmu(raw: string | undefined): number | undefined {
  const m = raw?.trim().match(/^(-?[\d.]+)(pt|px|in|cm|mm|pc)?$/);
  if (!m) return undefined;
  const n = Number(m[1]);
  if (!Number.isFinite(n)) return undefined;
  return Math.round(n * (m[2] ? EMU_PER[m[2] as keyof typeof EMU_PER] : EMU_PER.pt));
}

/** A handful of CSS/VML named colours (legacy docs mostly use hex; cover the
 *  common words so a `fillcolor="red"` still maps). */
const VML_NAMED_COLORS: Record<string, string> = {
  black: "000000", white: "ffffff", red: "ff0000", green: "008000", lime: "00ff00",
  blue: "0000ff", yellow: "ffff00", gray: "808080", grey: "808080", silver: "c0c0c0",
};

/** A VML colour attribute → "#rrggbb". Accepts "#RRGGBB", "#RGB", a named colour,
 *  and Word's "#rrggbb [theme-index]" form (takes the leading token). */
function vmlColor(raw: string | undefined): string | undefined {
  const tok = raw?.trim().split(/\s+/)[0];
  if (!tok) return undefined;
  const named = VML_NAMED_COLORS[tok.toLowerCase()];
  let hex = tok.startsWith("#") ? tok.slice(1) : named;
  if (!hex) return undefined;
  if (hex.length === 3) hex = hex.split("").map((c) => c + c).join("");
  return /^[0-9a-fA-F]{6}$/.test(hex) ? `#${hex.toLowerCase()}` : undefined;
}

/** A VML boolean attribute (filled / stroked / …). VML accepts several spellings:
 *  t/true/1/yes/on ⇒ true; f/false/0/no/off ⇒ false; anything else (incl. absent)
 *  ⇒ undefined (leave the caller's default). */
function vmlBool(raw: string | undefined): boolean | undefined {
  const v = raw?.trim().toLowerCase();
  if (v === "t" || v === "true" || v === "1" || v === "yes" || v === "on") return true;
  if (v === "f" || v === "false" || v === "0" || v === "no" || v === "off") return false;
  return undefined;
}

/** MSO shape-type numbers (v:shape @o:spt, or the `_tNNN` suffix of a shapetype
 *  @type reference) → the DrawingML preset the model draws. Unlisted → rect. */
const VML_SPT_PRESET: Record<string, string> = {
  "1": "rect", "2": "roundRect", "3": "ellipse", "4": "diamond",
  "5": "triangle", "6": "triangle", "13": "rightArrow", "20": "line",
  "202": "rect", // text box
};

/** The DrawingML preset for a concrete VML element tag (rect/oval/line/…). */
function vmlElementPreset(tagName: string): string | undefined {
  switch (tagName) {
    case "v:rect": return "rect";
    case "v:roundrect": return "roundRect";
    case "v:oval": return "ellipse";
    case "v:line": return "line";
    default: return undefined;
  }
}

/** The VML tags that denote a *drawn* shape instance (a v:shapetype is only a
 *  template, and an image is handled by the v:imagedata branch — neither counts). */
const VML_SHAPE_TAGS = new Set(["v:rect", "v:roundrect", "v:oval", "v:line", "v:shape"]);

/** Count the drawn VML shapes anywhere under a node. >1 means a group (v:group)
 *  or several shapes share one w:pict — we only import the first, so warn. */
function countVmlShapes(node: XmlNode): number {
  let n = 0;
  for (const c of children(node)) {
    if (VML_SHAPE_TAGS.has(c.tagName)) n++;
    n += countVmlShapes(c);
  }
  return n;
}

/** "x,y" VML coordinate pair (from/to on v:line) → [emuX, emuY]. */
function vmlPoint(raw: string | undefined): [number, number] | undefined {
  const parts = raw?.split(",");
  if (!parts || parts.length !== 2) return undefined;
  const x = vmlLengthToEmu(parts[0]);
  const y = vmlLengthToEmu(parts[1]);
  return x === undefined || y === undefined ? undefined : [x, y];
}

/** Parse the first recognized VML drawn shape / text box in a w:pict into the
 *  shared `kind:"shape"` IR (geometry + fill + stroke + size + text body). Legacy
 *  float positioning (w10:wrap / absolute style) is NOT modelled — the shape lands
 *  in the text flow, matching the read-only scope. */
function parseVmlShape(pict: XmlNode, ctx: ParseCtx): Extract<IRInline, { kind: "shape" }> | undefined {
  // Prefer a concrete primitive; fall back to a generic v:shape (text boxes,
  // typed shapes). A v:shapetype is only a template — never a drawn instance.
  const node =
    findDeep(pict, "v:rect") ?? findDeep(pict, "v:roundrect") ?? findDeep(pict, "v:oval") ??
    findDeep(pict, "v:line") ?? findDeep(pict, "v:shape");
  if (!node) return undefined;
  // A group (or several shapes in one w:pict) flattens to just the first — the
  // model has no VML group container. Surface the loss instead of hiding it.
  if (countVmlShapes(pict) > 1) {
    ctx.warnings.add("vml-group-flattened", "A VML shape group (or multiple VML shapes in one drawing) was found; only the first shape was imported.");
  }

  let preset = vmlElementPreset(node.tagName);
  if (preset === undefined) {
    // v:shape — resolve the MSO shape type from @o:spt, else the `_tNNN` suffix
    // of its @type shapetype reference (Word encodes the spt in the id).
    const spt = attr(node, "o:spt") ?? attr(node, "type")?.match(/_t(\d+)/)?.[1];
    preset = (spt !== undefined && VML_SPT_PRESET[spt]) || "rect";
  }
  const shape: Extract<IRInline, { kind: "shape" }> = { kind: "shape", preset };

  const style = vmlStyle(node);
  let widthEmu = vmlLengthToEmu(style["width"]);
  let heightEmu = vmlLengthToEmu(style["height"]);
  if (node.tagName === "v:line") {
    // A line's box comes from its from/to endpoints; keep a small floor on the
    // degenerate axis of a horizontal/vertical line so the box stays selectable.
    const from = vmlPoint(attr(node, "from")) ?? [0, 0];
    const to = vmlPoint(attr(node, "to")) ?? [0, 0];
    widthEmu ??= Math.max(Math.abs(to[0] - from[0]), EMU_PER.pt);
    heightEmu ??= Math.max(Math.abs(to[1] - from[1]), EMU_PER.pt);
  }
  if (widthEmu !== undefined) shape.widthEmu = widthEmu;
  if (heightEmu !== undefined) shape.heightEmu = heightEmu;

  // Fill: filled off (f/false/0/no/off) ⇒ no fill; else fillcolor (default fill
  // left to the model — VML's default fill is white, which the model draws too).
  const fillColor = vmlColor(attr(node, "fillcolor"));
  if (vmlBool(attr(node, "filled")) === false) shape.fill = { none: true };
  else if (fillColor) shape.fill = { color: fillColor };

  // Stroke: stroked off (f/false/0/no/off) ⇒ no outline; otherwise VML shapes are
  // stroked by default — synthesize the default black 0.75pt border when the
  // markup omits both colour and weight (a bare shape still shows an outline).
  const strokeColor = vmlColor(attr(node, "strokecolor"));
  const strokeWeightEmu = vmlLengthToEmu(attr(node, "strokeweight"));
  if (vmlBool(attr(node, "stroked")) === false) {
    shape.stroke = { none: true };
  } else {
    shape.stroke = { color: strokeColor ?? "#000000", widthPt: strokeWeightEmu !== undefined ? strokeWeightEmu / EMU_PER.pt : 0.75 };
  }

  // Text box body: v:textbox → w:txbxContent (a paragraph flow). Parse it as a
  // fresh block container (like a table cell) so surrounding field/control
  // tracking never leaks into the nested paragraphs — mirrors parseShapeWsp.
  const txbxContent = findDeep(node, "w:txbxContent") ?? findDeep(pict, "w:txbxContent");
  if (txbxContent) {
    const savedBlockStack = ctx.blockSdtStack;
    const savedTrackFields = ctx.trackFields;
    ctx.blockSdtStack = [];
    ctx.trackFields = false;
    const blocks: IRBlock[] = [];
    walkBlocks(children(txbxContent), blocks, ctx);
    ctx.blockSdtStack = savedBlockStack;
    ctx.trackFields = savedTrackFields;
    if (blocks.length > 0) shape.text = blocks;
  }
  return shape;
}

// ---------------------------------------------------------------------------
// Tables

/** w:tcMar / w:tblCellMar → per-side twips. Each side is a <w:top w:w=".."/> etc.
 *  Returns undefined when no side is present so the cascade can fall through. */
function decodeCellMargin(node: XmlNode | undefined): IRCellMargin | undefined {
  if (!node) return undefined;
  const side = (name: string): number | undefined => numAttr(el(node, name), "w:w");
  const m: IRCellMargin = {};
  const top = side("w:top");
  if (top !== undefined) m.top = top;
  const right = side("w:right");
  if (right !== undefined) m.right = right;
  const bottom = side("w:bottom");
  if (bottom !== undefined) m.bottom = bottom;
  const left = side("w:left");
  if (left !== undefined) m.left = left;
  return Object.keys(m).length > 0 ? m : undefined;
}

/** w:trPr → row properties. w:trHeight carries a height (twips) + an hRule
 *  ("auto"/"atLeast"/"exact"); an "auto" rule (or absent) is a pure hint we drop,
 *  keeping only enforceable atLeast/exact heights. w:cantSplit and w:tblHeader are
 *  on/off toggles. Returns undefined when no enforceable property is present. */
function parseRowProps(trPr: XmlNode | undefined): IRTableRow["props"] | undefined {
  if (!trPr) return undefined;
  const props: NonNullable<IRTableRow["props"]> = {};
  const trH = el(trPr, "w:trHeight");
  if (trH) {
    const h = numAttr(trH, "w:val");
    const rule = attr(trH, "w:hRule");
    // Only "atLeast"/"exact" pin the height; "auto" (the default) leaves it to content.
    if (h !== undefined && h > 0 && (rule === "atLeast" || rule === "exact")) {
      props.heightTwips = h;
      props.heightRule = rule;
    }
  }
  if (onOff(el(trPr, "w:cantSplit"))) props.cantSplit = true;
  if (onOff(el(trPr, "w:tblHeader"))) props.tblHeader = true;
  return Object.keys(props).length > 0 ? props : undefined;
}

function parseTable(tbl: XmlNode, ctx: ParseCtx): IRTable {
  const rows: IRTableRow[] = [];
  for (const tr of els(tbl, "w:tr")) {
    const cells: IRTableCell[] = [];
    for (const tc of els(tr, "w:tc")) {
      cells.push(parseCell(tc, ctx));
    }
    const row: IRTableRow = { cells };
    const props = parseRowProps(el(tr, "w:trPr"));
    if (props) row.props = props;
    rows.push(row);
  }
  const table: IRTable = { kind: "table", rows };
  const tblGrid = el(tbl, "w:tblGrid");
  if (tblGrid) {
    const widths = els(tblGrid, "w:gridCol")
      .map((col) => numAttr(col, "w:w"))
      .filter((w): w is number => w !== undefined && w > 0);
    if (widths.length > 0) table.colWidthsTwips = widths;
  }
  const tblPr = el(tbl, "w:tblPr");
  if (tblPr) {
    const styleId = val(tblPr, "w:tblStyle");
    if (styleId) table.styleId = styleId;
    const look = el(tblPr, "w:tblLook");
    if (look) {
      const on = (a: string): boolean => { const v = attr(look, a); return v === "1" || v === "true"; };
      table.look = {
        firstRow: on("w:firstRow"),
        lastRow: on("w:lastRow"),
        firstCol: on("w:firstColumn"),
        lastCol: on("w:lastColumn"),
        bandRows: !on("w:noHBand"),
        bandCols: !on("w:noVBand"),
      };
    }
    if (el(tblPr, "w:tblBorders")) table.bordersSpecified = true;
    const borders = decodeBorders(el(tblPr, "w:tblBorders"));
    if (borders) table.borders = borders;
    // Tri-state (issue #150): keep an explicit clear (null) so a table-level
    // "No Color" can override the table style's fill through the cascade.
    const shd = decodeShdFill(el(tblPr, "w:shd"));
    if (shd !== undefined) table.shd = shd;
    const cellMar = decodeCellMargin(el(tblPr, "w:tblCellMar"));
    if (cellMar) table.cellMarginTwips = cellMar;
    // Column-sizing strategy (w:tblLayout + w:tblW). Conservative mapping: a table
    // is autofit ONLY when it explicitly declares so (w:tblLayout type="autofit",
    // or a percentage w:tblW). An absent layout with auto/dxa width keeps the
    // historical fixed-proportional behavior, so existing imports never shift.
    const layoutEl = el(tblPr, "w:tblLayout");
    const layoutType = layoutEl ? attr(layoutEl, "w:type") : undefined;
    const tblW = el(tblPr, "w:tblW");
    const tblWType = tblW && attr(tblW, "w:type");
    const tblWVal = tblW ? numAttr(tblW, "w:w") : undefined;
    const isPct = tblWType === "pct" && (tblWVal ?? 0) > 0;
    if (layoutType === "autofit") table.widthMode = isPct ? "autofitWindow" : "autofitContents";
    else if (layoutType !== "fixed" && isPct) table.widthMode = "autofitWindow";
    // A fixed-layout table can still carry an explicit preferred TOTAL width
    // (w:tblW). Absolute (dxa) widths were previously dropped; capture them — plus a
    // percentage when the layout is explicitly fixed — so they round-trip.
    if (!table.widthMode) {
      if (tblWType === "dxa" && tblWVal !== undefined && tblWVal > 0) table.preferredWidthTwips = tblWVal;
      else if (isPct && layoutType === "fixed") table.preferredWidthPct = (tblWVal ?? 0) / 50;
    }
    // Table-level alignment (w:jc on tblPr — distinct from a paragraph's w:jc).
    const tblJc = el(tblPr, "w:jc");
    const tblJcVal = tblJc ? attr(tblJc, "w:val") : undefined;
    if (tblJcVal === "center") table.align = "center";
    else if (tblJcVal === "right" || tblJcVal === "end") table.align = "right";
    // w:tblInd — table indent from the leading edge. Only dxa (twips) is meaningful
    // for placement; a 0 or pct/auto indent round-trips as absent.
    const tblInd = el(tblPr, "w:tblInd");
    if (tblInd) {
      const indType = attr(tblInd, "w:type");
      const indVal = numAttr(tblInd, "w:w");
      if (indVal !== undefined && indVal !== 0 && (indType === "dxa" || indType === undefined)) {
        table.indentTwips = indVal;
      }
    }
    // w:bidiVisual — RTL visual column order (a CT_OnOff toggle).
    if (onOff(el(tblPr, "w:bidiVisual"))) table.bidiVisual = true;
    // w:tblOverlap — floating-table overlap behavior. "overlap" is Word's default,
    // so only "never" carries information; we still round-trip an explicit "overlap".
    const overlapVal = val(tblPr, "w:tblOverlap");
    if (overlapVal === "never") table.overlap = "never";
    else if (overlapVal === "overlap") table.overlap = "overlap";
    // w:tblCaption / w:tblDescription — accessibility title + alt text.
    const caption = val(tblPr, "w:tblCaption");
    if (caption) table.caption = caption;
    const description = val(tblPr, "w:tblDescription");
    if (description) table.description = description;
  }
  return table;
}

/** Cell content is a full block story (paragraphs, images, nested tables) —
 *  same walk as the body. */
function parseCell(tc: XmlNode, ctx: ParseCtx): IRTableCell {
  const tcPr = el(tc, "w:tcPr");
  const gridSpan = (tcPr && numAttr(el(tcPr, "w:gridSpan"), "w:val")) ?? 1;
  const vMerge = tcPr && el(tcPr, "w:vMerge");
  const vMergeContinue = !!vMerge && (attr(vMerge, "w:val") ?? "continue") === "continue";

  const blocks: IRBlock[] = [];
  // Cell content is a fresh block container: any block-level control wrapping the
  // TABLE is carried on the table block itself, so cell blocks must NOT re-inherit
  // it (that would re-wrap every cell on export). Block controls opened INSIDE the
  // cell still nest normally on this reset stack.
  const savedBlockStack = ctx.blockSdtStack;
  ctx.blockSdtStack = [];
  walkBlocks(children(tc), blocks, ctx);
  ctx.blockSdtStack = savedBlockStack;
  const cell: IRTableCell = { blocks, gridSpan, vMergeContinue };
  if (tcPr) {
    if (el(tcPr, "w:tcBorders")) cell.bordersSpecified = true;
    const borders = decodeBorders(el(tcPr, "w:tcBorders"));
    if (borders) cell.borders = borders;
    // Tri-state (issue #150): keep an explicit clear (null) so a cell "No Color"
    // overrides the table / table-style fill instead of falling through to it.
    const shd = decodeShdFill(el(tcPr, "w:shd"));
    if (shd !== undefined) cell.shd = shd;
    const cellMar = decodeCellMargin(el(tcPr, "w:tcMar"));
    if (cellMar) cell.marginTwips = cellMar;
    // w:tcW preferred cell width — abs (dxa twips) or pct (fiftieths of a percent).
    const tcW = el(tcPr, "w:tcW");
    const tcWVal = tcW ? numAttr(tcW, "w:w") : undefined;
    if (tcW && tcWVal !== undefined && tcWVal > 0) {
      const tcWType = attr(tcW, "w:type");
      if (tcWType === "pct") cell.preferredWidth = { type: "pct", frac: tcWVal / 5000 };
      else if (tcWType === "dxa" || tcWType === undefined) cell.preferredWidth = { type: "abs", twips: tcWVal };
    }
    // w:vAlign — vertical content alignment. "top" is the default, so we only
    // carry center/bottom (an explicit "top" round-trips as absent).
    const vAlignEl = el(tcPr, "w:vAlign");
    const vAlign = vAlignEl && attr(vAlignEl, "w:val");
    if (vAlign === "center" || vAlign === "bottom") cell.vAlign = vAlign;
    // w:textDirection — text flow direction. "lrTb" is the default (absent).
    const textDir = val(tcPr, "w:textDirection");
    if (textDir === "tbRl" || textDir === "btLr" || textDir === "lrTbV" || textDir === "tbRlV" || textDir === "tbLrV") {
      cell.textDirection = textDir;
    }
    // w:noWrap / w:tcFitText / w:hideMark — CT_OnOff cell toggles.
    if (onOff(el(tcPr, "w:noWrap"))) cell.noWrap = true;
    if (onOff(el(tcPr, "w:tcFitText"))) cell.fitText = true;
    if (onOff(el(tcPr, "w:hideMark"))) cell.hideMark = true;
  }
  return cell;
}

// ---------------------------------------------------------------------------
// Content controls (w:sdt)

/** w:sdtPr → IRSdtProps. Returns null when there are no properties at all
 *  (some producers emit bare w:sdt wrappers — those just unwrap). */
function parseSdtPr(pr: XmlNode | undefined): IRSdtProps | null {
  if (!pr) return null;
  const props: IRSdtProps = { type: "richText" };
  const alias = el(pr, "w:alias");
  const aliasVal = alias && attr(alias, "w:val");
  if (aliasVal) props.alias = aliasVal;
  const tag = el(pr, "w:tag");
  const tagVal = tag && attr(tag, "w:val");
  if (tagVal) props.tag = tagVal;
  if (el(pr, "w:text")) props.type = "plainText";
  const list = el(pr, "w:dropDownList") ?? el(pr, "w:comboBox");
  if (list) {
    props.type = list.tagName === "w:dropDownList" ? "dropDown" : "comboBox";
    props.listItems = els(list, "w:listItem").map((li) => ({
      display: attr(li, "w:displayText") ?? attr(li, "w:value") ?? "",
      value: attr(li, "w:value") ?? attr(li, "w:displayText") ?? "",
    }));
  }
  const date = el(pr, "w:date");
  if (date) {
    props.type = "date";
    const fmt = el(date, "w:dateFormat");
    const fmtVal = fmt && attr(fmt, "w:val");
    if (fmtVal) props.dateFormat = fmtVal;
  }
  const checkbox = el(pr, "w14:checkbox");
  if (checkbox) {
    props.type = "checkbox";
    const checked = el(checkbox, "w14:checked");
    const v = checked && attr(checked, "w14:val");
    props.checked = v === "1" || v === "true";
    const stateSym = (name: string): { font: string; val: string } | undefined => {
      const s = el(checkbox, name);
      const font = s && attr(s, "w14:font");
      const val = s && attr(s, "w14:val");
      return font && val ? { font, val: val.toUpperCase() } : undefined;
    };
    const checkedSym = stateSym("w14:checkedState");
    if (checkedSym) props.checkedSymbol = checkedSym;
    const uncheckedSym = stateSym("w14:uncheckedState");
    if (uncheckedSym) props.uncheckedSymbol = uncheckedSym;
  }
  if (el(pr, "w:showingPlcHdr")) props.placeholder = true;
  const lock = el(pr, "w:lock");
  const lockVal = lock && attr(lock, "w:val");
  if (lockVal === "contentLocked" || lockVal === "sdtContentLocked") props.lockContent = true;
  if (lockVal === "sdtLocked" || lockVal === "sdtContentLocked") props.lockControl = true;
  return props;
}

// ---------------------------------------------------------------------------
// Section properties

function parseSection(sectPr: XmlNode, warnings: WarningSink): IRSection {
  const section: IRSection = {};
  const pgSz = el(sectPr, "w:pgSz");
  const w = numAttr(pgSz, "w:w");
  if (w !== undefined) section.pageWidthTwips = w;
  const h = numAttr(pgSz, "w:h");
  if (h !== undefined) section.pageHeightTwips = h;
  const pgMar = el(sectPr, "w:pgMar");
  if (pgMar) {
    const top = numAttr(pgMar, "w:top");
    const right = numAttr(pgMar, "w:right");
    const bottom = numAttr(pgMar, "w:bottom");
    const left = numAttr(pgMar, "w:left");
    if (top !== undefined && right !== undefined && bottom !== undefined && left !== undefined) {
      section.marginTwips = { top, right, bottom, left };
    }
    const headerDist = numAttr(pgMar, "w:header");
    if (headerDist !== undefined) section.headerDistTwips = headerDist;
    const footerDist = numAttr(pgMar, "w:footer");
    if (footerDist !== undefined) section.footerDistTwips = footerDist;
  }
  const headerRefs = bandRefs(els(sectPr, "w:headerReference"));
  if (headerRefs) section.headerRefs = headerRefs;
  const footerRefs = bandRefs(els(sectPr, "w:footerReference"));
  if (footerRefs) section.footerRefs = footerRefs;
  if (el(sectPr, "w:titlePg")) section.titlePg = true;

  const cols = el(sectPr, "w:cols");
  if (cols) {
    const colEls = els(cols, "w:col");
    // count: explicit @w:num, else inferred from the w:col children.
    const count = numAttr(cols, "w:num") ?? (colEls.length > 1 ? colEls.length : 1);
    if (count > 1) {
      section.columns = { count };
      const space = numAttr(cols, "w:space");
      if (space !== undefined) section.columns.spaceTwips = space;
      if (attr(cols, "w:sep") === "1" || attr(cols, "w:sep") === "true") section.columns.sep = true;
      // Unequal columns: w:equalWidth="0" + a w:col per column.
      if (colEls.length === count) {
        const list = colEls.map((cel) => ({
          wTwips: numAttr(cel, "w:w") ?? 0,
          spaceTwips: numAttr(cel, "w:space") ?? 0,
        }));
        if (list.some((c) => c.wTwips > 0)) section.columns.cols = list;
      }
    }
  }
  const pgBorders = el(sectPr, "w:pgBorders");
  if (pgBorders) {
    const offsetFrom = attr(pgBorders, "w:offsetFrom");
    const borders: import("./types").IRPageBorders = {
      offsetFrom: offsetFrom === "text" ? "text" : "page",
    };
    const edge = (name: "top" | "right" | "bottom" | "left"): void => {
      const e = el(pgBorders, "w:" + name);
      if (!e) return;
      const sz = numAttr(e, "w:sz");
      const space = numAttr(e, "w:space");
      const color = attr(e, "w:color");
      borders[name] = {
        style: attr(e, "w:val") ?? "single",
        ...(sz !== undefined ? { sz } : {}),
        ...(space !== undefined ? { space } : {}),
        ...(color !== undefined ? { color } : {}),
      };
    };
    edge("top");
    edge("right");
    edge("bottom");
    edge("left");
    if (borders.top || borders.right || borders.bottom || borders.left) section.pageBorders = borders;
  }
  const pgNumStart = numAttr(el(sectPr, "w:pgNumType"), "w:start");
  if (pgNumStart !== undefined) section.pageNumberStart = pgNumStart;
  const lnNum = decodeLineNumbering(el(sectPr, "w:lnNumType"));
  if (lnNum) section.lineNumbering = lnNum;
  const bodyType = val(sectPr, "w:type");
  if (bodyType === "evenPage" || bodyType === "oddPage") section.breakType = bodyType;
  return section;
}

/** Collect w:headerReference / w:footerReference by w:type (default/first/even). */
function bandRefs(refs: XmlNode[]): BandRefs | undefined {
  if (refs.length === 0) return undefined;
  const out: BandRefs = {};
  for (const ref of refs) {
    const id = attr(ref, "r:id");
    if (!id) continue;
    const type = attr(ref, "w:type") ?? "default";
    if (type === "first") out.first = id;
    else if (type === "even") out.even = id;
    else out.default = id;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

// ---------------------------------------------------------------------------
// Footnotes (footnotes.xml)

/** footnotes.xml → IR block stories keyed by footnote id. The standard
 *  separator/continuation pseudo-notes (negative ids / type attrs) are skipped;
 *  only real notes are returned. Each note's leading w:footnoteRef placeholder
 *  is dropped (parseRun ignores it) — the engine paints the number. */
export function parseFootnotesXml(
  xmlText: string,
  partName: string,
  warnings: WarningSink,
  sdts: Record<string, IRSdtProps> = {},
): Map<string, IRBlock[]> {
  const out = new Map<string, IRBlock[]>();
  const root = rootEl(parseXml(xmlText, partName), "w:footnotes");
  if (!root) return out;
  const ctx: ParseCtx = {
    warnings,
    fieldTokens: false,
    sdts,
    nextSdt: { n: Object.keys(sdts).length },
    blockSdtStack: [],
    inlineSdtStack: [],
    pendingBookmarks: [],
    currentBookmarks: null,
    pendingMarkers: [],
    currentMarkers: null,
    trackFields: false,
    fieldTrack: newFieldTrack(),
  };
  for (const note of els(root, "w:footnote")) {
    const fnId = attr(note, "w:id");
    if (fnId === undefined) continue;
    const type = attr(note, "w:type"); // "separator" | "continuationSeparator" | …
    if (type) continue; // pseudo-notes, not real footnotes
    const blocks: IRBlock[] = [];
    walkBlocks(children(note), blocks, ctx);
    out.set(fnId, blocks);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Endnotes (endnotes.xml) — mirror of parseFootnotesXml. Endnotes lay out at
// the END of the document rather than each page bottom; the IR is identical.

/** endnotes.xml → IR block stories keyed by endnote id. The standard
 *  separator/continuation pseudo-notes (type attrs) are skipped; only real
 *  notes are returned. Each note's leading w:endnoteRef placeholder is dropped
 *  (parseRun ignores it) — the engine paints the number. */
export function parseEndnotesXml(
  xmlText: string,
  partName: string,
  warnings: WarningSink,
  sdts: Record<string, IRSdtProps> = {},
): Map<string, IRBlock[]> {
  const out = new Map<string, IRBlock[]>();
  const root = rootEl(parseXml(xmlText, partName), "w:endnotes");
  if (!root) return out;
  const ctx: ParseCtx = {
    warnings,
    fieldTokens: false,
    sdts,
    nextSdt: { n: Object.keys(sdts).length },
    blockSdtStack: [],
    inlineSdtStack: [],
    pendingBookmarks: [],
    currentBookmarks: null,
    pendingMarkers: [],
    currentMarkers: null,
    trackFields: false,
    fieldTrack: newFieldTrack(),
  };
  for (const note of els(root, "w:endnote")) {
    const enId = attr(note, "w:id");
    if (enId === undefined) continue;
    const type = attr(note, "w:type"); // "separator" | "continuationSeparator" | …
    if (type) continue; // pseudo-notes, not real endnotes
    const blocks: IRBlock[] = [];
    walkBlocks(children(note), blocks, ctx);
    out.set(enId, blocks);
  }
  return out;
}
