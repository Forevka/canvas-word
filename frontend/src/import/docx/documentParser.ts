// document.xml / header / footer body walk → IR. Pure: XML text in, IR + warnings out.
//
// Decodes direct formatting only. Style references (w:pStyle, w:rStyle) are
// recorded in the IR unresolved — styles.ts resolves them into effective
// properties during mapping. Everything the parser understands goes into the
// IR; mapToModel decides what the model can hold.

import type { CharStyle, FieldDef, FieldSpec } from "@cw/shared";
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
  IRTable,
  IRTableCell,
  IRTableRow,
} from "./types";
import { decodeBorders, decodeShdFill } from "./borders";
import { decodeParaProps, decodeRunProps } from "./props";
import { attr, children, el, els, findDeep, numAttr, parseXml, rootEl, textOf, val, type XmlNode } from "./xml";

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

function walkBlocks(nodes: XmlNode[], out: IRBlock[], ctx: ParseCtx): void {
  for (const node of nodes) {
    // A custom field open here marks the block being built; one that opens/closes
    // mid-block is captured during the parse (decideCustomField sets markBlock).
    if (ctx.trackFields) ctx.fieldTrack.markBlock = ctx.fieldTrack.openId;
    switch (node.tagName) {
      case "w:p": {
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
          if (idAttr) ctx.pendingMarkers.push({ id: idAttr, name, kind: "start", offset: 0 });
        }
        break;
      }
      case "w:bookmarkEnd": {
        const idAttr = attr(node, "w:id");
        if (idAttr) ctx.pendingMarkers.push({ id: idAttr, kind: "end", offset: 0 });
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
            if (inline.kind !== "run") continue;
            if (relId) inline.props.linkRelId = relId;
            else if (anchor) inline.props.linkAnchor = anchor;
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
          ctx.currentMarkers.push({ id: idAttr, name, kind: "start", offset: inlineOffset(out) });
        }
        break;
      }
      case "w:bookmarkEnd": {
        const idAttr = attr(node, "w:id");
        if (idAttr && ctx.currentMarkers) ctx.currentMarkers.push({ id: idAttr, kind: "end", offset: inlineOffset(out) });
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
        case "w:footnoteRef":
          break; // the auto-number placeholder inside a footnote BODY — engine paints it
        default:
          break; // w:rPr, w:lastRenderedPageBreak, w:endnoteReference, …
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

/** DrawingML: w:drawing → wp:inline|wp:anchor → … → a:blip r:embed. The exact
 *  nesting varies by producer, so the blip is found by deep search. */
function parseDrawing(drawing: XmlNode, ctx: ParseCtx): IRInline | undefined {
  const anchor = el(drawing, "wp:anchor");
  const container = el(drawing, "wp:inline") ?? anchor;
  if (!container) return undefined;
  const blip = findDeep(container, "a:blip");
  const relId = blip && attr(blip, "r:embed");
  if (!relId) {
    ctx.warnings.add("images-skipped", "A drawing without an embedded image reference was skipped.");
    return undefined;
  }
  const extent = el(container, "wp:extent");
  const image: IRInline = { kind: "image", relId, anchored: !!anchor };
  const cx = numAttr(extent, "cx");
  if (cx !== undefined) image.widthEmu = cx;
  const cy = numAttr(extent, "cy");
  if (cy !== undefined) image.heightEmu = cy;
  if (anchor) {
    // Square/tight/through wrap maps onto the model's "square" float; wrapNone
    // (behind/in-front of text) becomes an absolutely-positioned anchor; only
    // topAndBottom/none-without-wrapNone fall back to block flow.
    const square = el(anchor, "wp:wrapSquare") ?? el(anchor, "wp:wrapTight") ?? el(anchor, "wp:wrapThrough");
    const wrapNone = el(anchor, "wp:wrapNone");
    image.anchorWrap = square ? "square" : "block";
    const posH = el(anchor, "wp:positionH");
    const posV = el(anchor, "wp:positionV");
    const alignText = posH && el(posH, "wp:align");
    const align = alignText && textOf(alignText);
    if (align === "left" || align === "right" || align === "center") image.anchorAlign = align;
    if (wrapNone) {
      const dec = findDeep(anchor, "adec:decorative");
      const z = numAttr(anchor, "relativeHeight");
      image.anchorFloat = {
        behind: attr(anchor, "behindDoc") === "1",
        offsetXEmu: posOffsetEmu(posH),
        offsetYEmu: posOffsetEmu(posV),
        relFromH: relFromH(posH),
        relFromV: relFromV(posV),
        decorative: !!dec && attr(dec, "val") !== "0",
        ...(z !== undefined ? { z } : {}),
      };
    } else if (!square) {
      ctx.warnings.add(
        "images-anchored",
        "Some floating images (overlapping or top-and-bottom wrap) were placed in the text flow.",
      );
    }
  }
  return image;
}

/** wp:positionH|V → wp:posOffset (EMU, signed). Absent/non-numeric → 0. */
function posOffsetEmu(pos: XmlNode | undefined): number {
  const off = pos && el(pos, "wp:posOffset");
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

/** Legacy VML (w:pict → v:shape → v:imagedata r:id), still produced for some
 *  content. Size comes from the shape's style="width:…pt;height:…pt". */
function parseVmlPict(pict: XmlNode, ctx: ParseCtx): IRInline | undefined {
  const imagedata = findDeep(pict, "v:imagedata");
  const relId = imagedata && attr(imagedata, "r:id");
  if (!relId) {
    ctx.warnings.add("images-skipped", "A legacy picture without an image reference was skipped.");
    return undefined;
  }
  const image: IRInline = { kind: "image", relId, anchored: false };
  const shape = findDeep(pict, "v:shape");
  const style = (shape && attr(shape, "style")) ?? "";
  const dim = (name: string): number | undefined => {
    const m = style.match(new RegExp(`(?:^|;)\\s*${name}:([\\d.]+)pt`));
    return m ? Number(m[1]) * 12700 : undefined; // pt → EMU (12700 EMU/pt)
  };
  const w = dim("width");
  if (w !== undefined) image.widthEmu = w;
  const h = dim("height");
  if (h !== undefined) image.heightEmu = h;
  return image;
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

function parseTable(tbl: XmlNode, ctx: ParseCtx): IRTable {
  const rows: IRTableRow[] = [];
  for (const tr of els(tbl, "w:tr")) {
    const cells: IRTableCell[] = [];
    for (const tc of els(tr, "w:tc")) {
      cells.push(parseCell(tc, ctx));
    }
    rows.push({ cells });
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
    const shd = decodeShdFill(el(tblPr, "w:shd"));
    if (shd) table.shd = shd;
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
    const shd = decodeShdFill(el(tcPr, "w:shd"));
    if (shd) cell.shd = shd;
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
