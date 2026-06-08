// document.xml / header / footer body walk → IR. Pure: XML text in, IR + warnings out.
//
// Decodes direct formatting only. Style references (w:pStyle, w:rStyle) are
// recorded in the IR unresolved — styles.ts resolves them into effective
// properties during mapping. Everything the parser understands goes into the
// IR; mapToModel decides what the model can hold.

import { ImportError, WarningSink } from "./types";
import type {
  IRBlock,
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
import { decodeParaProps, decodeRunProps } from "./props";
import { attr, children, el, els, findDeep, numAttr, parseXml, rootEl, textOf, type XmlNode } from "./xml";

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
  /** Bookmark names seen between paragraphs (block-level w:bookmarkStart),
   *  attached to the NEXT paragraph. */
  pendingBookmarks: string[];
  /** Bookmark names collected inside the paragraph currently being walked. */
  currentBookmarks: string[] | null;
}

export function parseDocumentXml(xmlText: string, partName: string, warnings: WarningSink): IRDocument {
  const doc = rootEl(parseXml(xmlText, partName), "w:document");
  const body = doc && el(doc, "w:body");
  if (!body) {
    throw new ImportError("MALFORMED_XML", `${partName} has no w:document/w:body root.`);
  }
  const sdts: Record<string, IRSdtProps> = {};
  const ctx: ParseCtx = { warnings, fieldTokens: false, sdts, nextSdt: { n: 0 }, pendingBookmarks: [], currentBookmarks: null };

  const blocks: IRBlock[] = [];
  walkBlocks(children(body), blocks, ctx);

  const sectPr = el(body, "w:sectPr");
  const section = sectPr ? parseSection(sectPr, warnings) : null;

  return { blocks, section, sdts };
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
  const ctx: ParseCtx = { warnings, fieldTokens: true, sdts, nextSdt: { n: Object.keys(sdts).length }, pendingBookmarks: [], currentBookmarks: null };
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
    switch (node.tagName) {
      case "w:p":
        out.push(parseParagraph(node, ctx));
        break;
      case "w:tbl":
        out.push(parseTable(node, ctx));
        break;
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
        const sdtId = `sdt${ctx.nextSdt.n++}`;
        ctx.sdts[sdtId] = props;
        const inner: IRBlock[] = [];
        walkBlocks(children(content), inner, ctx);
        for (const b of inner) tagBlockSdt(b, sdtId);
        out.push(...inner);
        break;
      }
      case "w:bookmarkStart": {
        // Block-level bookmark — attach to the next paragraph.
        const name = bookmarkName(node);
        if (name) ctx.pendingBookmarks.push(name);
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
  // Complex fields (w:fldChar begin → instr → separate → result → end) span
  // multiple runs; the state lives at paragraph scope.
  const field: FieldState = { depth: 0, instr: "", suppressResult: false };
  walkInlines(children(p), inlines, ctx, field);
  ctx.currentBookmarks = null;
  const para: IRParagraph = { kind: "paragraph", props, inlines };
  if (bookmarks.length > 0) para.bookmarks = bookmarks;
  return para;
}

interface FieldState {
  depth: number;
  instr: string;
  /** True between "separate" and "end" when the field became a token — the
   *  cached result runs must not ALSO be emitted. */
  suppressResult: boolean;
}

/** " PAGE \\* MERGEFORMAT " → "{page}". Layout substitutes per page. */
function fieldToken(instr: string): string | undefined {
  if (/\bNUMPAGES\b/.test(instr)) return "{pages}"; // before PAGE — substring!
  if (/\bPAGE\b/.test(instr)) return "{page}";
  return undefined;
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
        const sdtId = `sdt${ctx.nextSdt.n++}`;
        ctx.sdts[sdtId] = props;
        const inner: IRInline[] = [];
        walkInlines(children(content), inner, ctx, field);
        for (const inline of inner) {
          if (inline.kind === "run") inline.sdtId = sdtId;
        }
        out.push(...inner);
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
        if (name && ctx.currentBookmarks) ctx.currentBookmarks.push(name);
        break;
      }
      default:
        break; // w:pPr, w:proofErr, w:bookmarkEnd, comment ranges, …
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
  let text = "";
  const flush = (): void => {
    if (text.length > 0) out.push({ kind: "run", text, props });
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
        case "w:br":
          flush();
          if (attr(node, "w:type") === "page") out.push({ kind: "break", page: true });
          else out.push({ kind: "break" });
          break;
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
          if (field.depth > 0) field.instr += textOf(node);
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
}

function handleFldChar(
  node: XmlNode,
  runProps: IRRunProps,
  out: IRInline[],
  ctx: ParseCtx,
  field: FieldState,
): void {
  switch (attr(node, "w:fldCharType")) {
    case "begin":
      field.depth++;
      if (field.depth === 1) {
        field.instr = "";
        field.suppressResult = false;
      }
      break;
    case "separate": {
      if (field.depth !== 1) break;
      const token = ctx.fieldTokens ? fieldToken(field.instr) : undefined;
      if (token) {
        out.push({ kind: "run", text: token, props: runProps });
        field.suppressResult = true; // the cached result would duplicate the token
      }
      break;
    }
    case "end":
      if (field.depth > 0) field.depth--;
      if (field.depth === 0) {
        // Field without a "separate" (no cached result): emit the token now.
        if (!field.suppressResult) {
          const token = ctx.fieldTokens ? fieldToken(field.instr) : undefined;
          if (token) out.push({ kind: "run", text: token, props: runProps });
        }
        field.instr = "";
        field.suppressResult = false;
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
    // Square/tight/through wrap maps onto the model's "square" float; the rest
    // (wrapNone = behind/in-front, topAndBottom) fall back to block flow.
    const square = el(anchor, "wp:wrapSquare") ?? el(anchor, "wp:wrapTight") ?? el(anchor, "wp:wrapThrough");
    image.anchorWrap = square ? "square" : "block";
    if (!square) {
      ctx.warnings.add(
        "images-anchored",
        "Some floating images (overlapping or top-and-bottom wrap) were placed in the text flow.",
      );
    }
    const posH = el(anchor, "wp:positionH");
    const alignText = posH && el(posH, "wp:align");
    const align = alignText && textOf(alignText);
    if (align === "left" || align === "right" || align === "center") image.anchorAlign = align;
  }
  return image;
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
  walkBlocks(children(tc), blocks, ctx);
  return { blocks, gridSpan, vMergeContinue };
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

/** Block-level control: every run in the wrapped paragraphs (including table
 *  cell paragraphs) joins the control. */
function tagBlockSdt(block: IRBlock, sdtId: string): void {
  if (block.kind === "paragraph") {
    for (const inline of block.inlines) {
      if (inline.kind === "run" && inline.sdtId === undefined) inline.sdtId = sdtId;
    }
    return;
  }
  for (const row of block.rows) {
    for (const cell of row.cells) for (const b of cell.blocks) tagBlockSdt(b, sdtId);
  }
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
  }
  const headerRelId = pickHeaderFooterRef(els(sectPr, "w:headerReference"), warnings);
  if (headerRelId) section.headerRelId = headerRelId;
  const footerRelId = pickHeaderFooterRef(els(sectPr, "w:footerReference"), warnings);
  if (footerRelId) section.footerRelId = footerRelId;

  const cols = el(sectPr, "w:cols");
  if (cols) {
    const count = numAttr(cols, "w:num") ?? 1;
    if (count > 1) {
      section.columns = { count };
      const space = numAttr(cols, "w:space");
      if (space !== undefined) section.columns.spaceTwips = space;
    }
  }
  const pgNumStart = numAttr(el(sectPr, "w:pgNumType"), "w:start");
  if (pgNumStart !== undefined) section.pageNumberStart = pgNumStart;
  return section;
}

/** The model has ONE header/footer story; Word allows default/first/even
 *  variants. Use default; fall back to whatever exists with a warning. */
function pickHeaderFooterRef(refs: XmlNode[], warnings: WarningSink): string | undefined {
  if (refs.length === 0) return undefined;
  const byType = (t: string): string | undefined => {
    const ref = refs.find((r) => attr(r, "w:type") === t);
    return ref ? attr(ref, "r:id") : undefined;
  };
  const def = byType("default");
  if (def) {
    if (refs.length > 1) {
      warnings.add("header-variants", "First/even-page header/footer variants were ignored (default used).");
    }
    return def;
  }
  warnings.add("header-variants", "No default header/footer — using the first/even variant for all pages.");
  return attr(refs[0]!, "r:id");
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
    pendingBookmarks: [],
    currentBookmarks: null,
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
