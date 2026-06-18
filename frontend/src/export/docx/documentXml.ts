// Document model -> WordprocessingML. The inverse of the importer's
// documentParser + mapToModel: each model field maps back to the w:* element the
// decoder reads, so a written file re-imports to an equal model.
//
// Runs carry FULL direct formatting (every toggle explicit on/off, font, size,
// color) so a paragraph's w:pStyle can't leak run properties back through the
// cascade on re-import. Paragraph styles ride as w:pStyle only.

import type {
  Block,
  CellBorders,
  CharStyle,
  FieldDef,
  Paragraph,
  ParaStyle,
  Run,
  SdtProps,
  SectionPatch,
  SectionProps,
  TableBlock,
  TableCell,
  TableCondOverrides,
} from "@cw/shared";
import { inRange, parseTocInstruction } from "@cw/shared";
import { pxToEighthPoints, pxToEmu, pxToTwips } from "../units";
import { hexColor as hex } from "./mappings";
import { MediaManager } from "./mediaPack";
import { REL, RelManager } from "./relationships";
import { paraCoreXml, runPropsXml as rPrXml } from "./styleProps";
import { el, escapeText, textEl, WML_NS, XML_DECL } from "./xmlWrite";
export interface ExportBookmarkMark {
  id: number;
  name?: string;
  kind: "start" | "end";
  offset: number;
}

export interface PartCtx {
  rels: RelManager;
  media: MediaManager;
  warn: (code: string, detail?: string) => void;
  sdts: Record<string, SdtProps>;
  /** Doc-wide unique ids for w:bookmarkStart / w:footnoteReference fallbacks. */
  nextId: () => number;
  /** model blockId -> bookmark start/end markers anchored there (body + bands).
   *  `id` pairs a start with its end; `offset` is the UTF-16 position in the
   *  paragraph text where the marker sits. */
  bookmarksByBlock: Map<string, ExportBookmarkMark[]>;
  /** model list id -> Word-valid integer numId (shared with numbering.xml). */
  listIdMap: Map<string, number>;
  /** Header/footer parts: {page}/{pages} run text becomes live PAGE/NUMPAGES
   *  fields (the inverse of the importer's header/footer fieldTokens mode). In the
   *  body those tokens stay literal text — symmetric with import, which only
   *  tokenizes fields in bands. */
  fieldTokens?: boolean;
  /** Heading blockId → displayed page number, from a layout pass — the cached
   *  result baked into each TOC entry's PAGEREF field so a non-updating viewer
   *  shows correct numbers (Word recomputes on F9 regardless). */
  tocPages?: Map<string, number>;
  /** Ensure the heading block has an in-document bookmark (synthesizing one if the
   *  doc has none) and return its name, so a live PAGEREF can point at it. Returns
   *  undefined if the target block doesn't exist (dangling TOC entry). */
  ensureTocBookmark?: (blockId: string) => string | undefined;
  /** The document's preserved `TOC` field instruction, re-emitted verbatim (and its
   *  switches honored). Absent → a sensible default ` TOC \o "1-3" \h \z \u `. */
  tocInstruction?: string;
  /** Custom (non-built-in) fields keyed by id — a contiguous run of body blocks
   *  sharing a `fieldId` is re-emitted as the field's complex region (begin /
   *  instrText / separate … end), wrapping the result blocks. */
  fields?: Record<string, FieldDef>;
}

// ---------------------------------------------------------------------------
// Runs

/** w:t / w:tab content for a run's text. */
function runContent(text: string): string {
  if (text.length === 0) return "";
  const parts = text.split("\t");
  let out = "";
  parts.forEach((piece, i) => {
    if (i > 0) out += el("w:tab");
    if (piece.length > 0) out += textEl(piece);
  });
  return out;
}

// Page-number tokens the layout substitutes per page ({page}, {page:roman}, …,
// {pages}). On export they become live PAGE/NUMPAGES fields — the inverse of the
// importer's fieldToken (documentParser.ts) — so a re-rendered footer shows the real
// page number instead of literal "{page}" text.
const PAGE_TOKEN_RE = /\{(page(?::(?:roman|Roman|alpha|Alpha))?|pages)\}/g;
const PAGE_TOKEN_TEST = /\{(?:page(?::(?:roman|Roman|alpha|Alpha))?|pages)\}/;

/** OOXML field instruction for a page-number token. */
function fieldInstr(token: string): string {
  if (token === "pages") return " NUMPAGES \\* MERGEFORMAT ";
  const fmt = token.includes(":") ? token.slice(token.indexOf(":") + 1) : undefined;
  const sw =
    fmt === "roman" ? " \\* roman"
    : fmt === "Roman" ? " \\* ROMAN"
    : fmt === "alpha" ? " \\* alphabetic"
    : fmt === "Alpha" ? " \\* ALPHABETIC"
    : "";
  return ` PAGE${sw} \\* MERGEFORMAT `;
}

/** Serialize a run's text, emitting each page-number token as a w:fldSimple field
 *  (carrying the run's rPr) and the literal segments around it as plain runs. The
 *  field's cached result is an empty run — Word recomputes it, and re-import reads
 *  the instruction, not the result. */
function runBodyWithFields(text: string, rPr: string): string {
  let out = "";
  let last = 0;
  PAGE_TOKEN_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = PAGE_TOKEN_RE.exec(text)) !== null) {
    if (m.index > last) out += el("w:r", undefined, rPr + runContent(text.slice(last, m.index)));
    out += el("w:fldSimple", { "w:instr": fieldInstr(m[1]!) }, el("w:r", undefined, rPr));
    last = m.index + m[0].length;
  }
  if (last < text.length) out += el("w:r", undefined, rPr + runContent(text.slice(last)));
  return out;
}

function singleRun(run: Run, ctx: PartCtx): string {
  const s = run.style;
  const rPr = rPrXml(s);
  if (s.footnoteRef) {
    // model footnoteRef "fn<docxId>" -> w:footnoteReference w:id="<docxId>".
    const docxId = s.footnoteRef.replace(/^fn/, "");
    return el("w:r", undefined, rPr + el("w:footnoteReference", { "w:id": docxId }));
  }
  // Page-number tokens -> live PAGE/NUMPAGES fields (header/footer only). Skipped
  // inside hyperlinks (page tokens never live in a link) so the wrapper logic
  // below stays simple.
  if (ctx.fieldTokens && !s.link && PAGE_TOKEN_TEST.test(run.text)) {
    return runBodyWithFields(run.text, rPr);
  }
  const r = el("w:r", undefined, rPr + runContent(run.text));
  // Hyperlink wrapper. Anchor (#name) is in-document; otherwise an external rel.
  if (s.link) {
    if (s.link.startsWith("#")) return el("w:hyperlink", { "w:anchor": s.link.slice(1) }, r);
    const id = ctx.rels.add(REL.hyperlink, s.link, true);
    return el("w:hyperlink", { "r:id": id }, r);
  }
  return r;
}

/** Serialize a run list: group consecutive runs sharing an sdtId into a w:sdt
 *  content control (outer), and within each segment group runs sharing a built-in
 *  field's id into a complex field (inner). */
function runsXml(runs: Run[], ctx: PartCtx): string {
  let out = "";
  let i = 0;
  while (i < runs.length) {
    const sdtId = runs[i]!.style.sdtId;
    if (sdtId) {
      let j = i;
      const seg: Run[] = [];
      while (j < runs.length && runs[j]!.style.sdtId === sdtId) seg.push(runs[j++]!);
      out += el("w:sdt", undefined, sdtPrXml(ctx.sdts[sdtId]) + el("w:sdtContent", undefined, fieldRunsXml(seg, ctx)));
      i = j;
    } else {
      let j = i;
      const seg: Run[] = [];
      while (j < runs.length && !runs[j]!.style.sdtId) seg.push(runs[j++]!);
      out += fieldRunsXml(seg, ctx);
      i = j;
    }
  }
  return out;
}

/** Group runs sharing a built-in field's id into a complex field
 *  (begin/instrText/separate/result/end). PAGE/NUMPAGES emit an EMPTY cached
 *  result (their run text is our internal {page} token — Word recomputes; on
 *  re-import we re-materialize the token); DATE/TIME/IF emit the materialized
 *  result text. Untagged runs pass through as plain runs. */
function fieldRunsXml(runs: Run[], ctx: PartCtx): string {
  let out = "";
  let i = 0;
  while (i < runs.length) {
    const fid = runs[i]!.style.fieldId;
    const def = fid ? ctx.fields?.[fid] : undefined;
    if (fid && def && def.kind === "builtin") {
      let j = i;
      const seg: Run[] = [];
      while (j < runs.length && runs[j]!.style.fieldId === fid) seg.push(runs[j++]!);
      const isToken = def.spec?.type === "PAGE" || def.spec?.type === "NUMPAGES";
      const result = seg.map((r) => el("w:r", undefined, rPrXml(r.style) + (isToken ? "" : runContent(r.text)))).join("");
      out += fldRun("begin") + instrRun(def.instruction) + fldRun("separate") + result + fldRun("end");
      i = j;
    } else {
      out += singleRun(runs[i]!, ctx);
      i++;
    }
  }
  return out;
}

function sdtPrXml(props: SdtProps | undefined): string {
  if (!props) return el("w:sdtPr");
  const c: string[] = [];
  if (props.alias) c.push(el("w:alias", { "w:val": props.alias }));
  if (props.tag) c.push(el("w:tag", { "w:val": props.tag }));
  if (props.placeholder) c.push(el("w:showingPlcHdr"));
  if (props.type === "plainText") c.push(el("w:text"));
  else if (props.type === "dropDown" || props.type === "comboBox") {
    const items = (props.listItems ?? [])
      .map((it) => el("w:listItem", { "w:displayText": it.display, "w:value": it.value }))
      .join("");
    c.push(el(props.type === "dropDown" ? "w:dropDownList" : "w:comboBox", undefined, items));
  } else if (props.type === "date") {
    c.push(el("w:date", undefined, props.dateFormat ? el("w:dateFormat", { "w:val": props.dateFormat }) : ""));
  } else if (props.type === "checkbox") {
    c.push(el("w14:checkbox", undefined, el("w14:checked", { "w14:val": props.checked ? "1" : "0" })));
  }
  if (props.lockContent || props.lockControl) {
    const v =
      props.lockContent && props.lockControl ? "sdtContentLocked"
      : props.lockContent ? "contentLocked"
      : "sdtLocked";
    c.push(el("w:lock", { "w:val": v }));
  }
  return el("w:sdtPr", undefined, c.join(""));
}

// ---------------------------------------------------------------------------
// Paragraphs

function pPrXml(style: ParaStyle, ctx: PartCtx, markRun?: CharStyle): string {
  const c: string[] = [];
  if (style.namedStyle) c.push(el("w:pStyle", { "w:val": style.namedStyle }));
  if (style.list) {
    const numId = ctx.listIdMap.get(style.list.listId) ?? 0;
    c.push(el("w:numPr", undefined, el("w:ilvl", { "w:val": style.list.level }) + el("w:numId", { "w:val": numId })));
  }
  if (style.pageBreakBefore) c.push(el("w:pageBreakBefore"));
  if (style.keepWithNext) c.push(el("w:keepNext"));
  if (style.keepLinesTogether) c.push(el("w:keepLines"));

  // spacing / indent / jc / tabs — shared with the TOC generator.
  c.push(paraCoreXml(style));

  // The paragraph mark's run props size empty lines on re-import.
  if (markRun) c.push(rPrXml(markRun));

  // A mid-document section break rides the paragraph that ENDS the section.
  if (style.sectionBreak) c.push(sectPrXml(style.sectionBreak.props, ctx, () => "", false));

  return el("w:pPr", undefined, c.join(""));
}

const bookmarkEl = (m: ExportBookmarkMark): string =>
  m.kind === "start"
    ? el("w:bookmarkStart", { "w:id": m.id, "w:name": m.name ?? "" })
    : el("w:bookmarkEnd", { "w:id": m.id });

function paragraphXml(p: Paragraph, ctx: PartCtx): string {
  const isEmpty = p.runs.every((r) => r.text.length === 0) && !p.runs.some((r) => r.style.footnoteRef);
  const markRun = isEmpty && p.runs[0] ? p.runs[0].style : undefined;
  const pPr = pPrXml(p.style, ctx, markRun);
  const runs = isEmpty ? [] : p.runs;
  const marks = ctx.bookmarksByBlock.get(p.id);

  if (!marks || marks.length === 0) {
    return el("w:p", undefined, pPr + runsXml(runs, ctx));
  }
  // A paragraph whose runs join an sdt content control can't have markers spliced
  // between its runs without breaking the w:sdt grouping — bracket the whole
  // paragraph at block level instead (start before w:p, end after).
  if (runs.some((r) => r.style.sdtId)) {
    const pre = marks.filter((m) => m.kind === "start").map(bookmarkEl).join("");
    const post = marks.filter((m) => m.kind === "end").map(bookmarkEl).join("");
    return pre + el("w:p", undefined, pPr + runsXml(runs, ctx)) + post;
  }
  // Offset-aware: emit each run, inserting markers at run boundaries (ends before
  // starts at a shared offset so a point bookmark nests correctly).
  const sorted = [...marks].sort((a, b) => a.offset - b.offset || (a.kind === "end" ? -1 : 1));
  let body = pPr;
  let cum = 0;
  let mi = 0;
  const flush = (upto: number): void => {
    while (mi < sorted.length && sorted[mi]!.offset <= upto) body += bookmarkEl(sorted[mi++]!);
  };
  let ri = 0;
  while (ri < runs.length) {
    flush(cum);
    // Keep built-in inline fields whole: emit the contiguous fieldId-tagged runs as
    // one complex field (begin/instr/separate/result/end) via fieldRunsXml, exactly
    // as the no-marker path does — otherwise splicing run-by-run drops the field
    // wrapping and the result re-imports as plain text. A bookmark that falls
    // between fields still lands at the right boundary; one inside a field result is
    // flushed just after it (a field result is atomic on re-import anyway).
    const fid = runs[ri]!.style.fieldId;
    const def = fid ? ctx.fields?.[fid] : undefined;
    if (fid && def && def.kind === "builtin") {
      const seg: Run[] = [];
      while (ri < runs.length && runs[ri]!.style.fieldId === fid) { seg.push(runs[ri]!); cum += runs[ri]!.text.length; ri++; }
      body += fieldRunsXml(seg, ctx);
    } else {
      body += singleRun(runs[ri]!, ctx);
      cum += runs[ri]!.text.length;
      ri++;
    }
  }
  flush(cum);
  while (mi < sorted.length) body += bookmarkEl(sorted[mi++]!); // trailing markers past text end
  return el("w:p", undefined, body);
}

// ---------------------------------------------------------------------------
// Tables

function bordersXml(tag: string, b: NonNullable<TableCell["borders"]>): string {
  const edge = (name: string, spec: { color: string; widthPx: number; style?: string } | undefined): string => {
    if (!spec) return "";
    const val = spec.style === "double" ? "double" : spec.style === "dashed" ? "dashed" : spec.style === "dotted" ? "dotted" : "single";
    return el("w:" + name, { "w:val": val, "w:sz": pxToEighthPoints(spec.widthPx), "w:space": 0, "w:color": hex(spec.color) });
  };
  const inner = edge("top", b.top) + edge("left", b.left) + edge("bottom", b.bottom) + edge("right", b.right);
  // Always emit the element when a borders object is present, even with no edges:
  // an empty <w:tcBorders/> is the author's explicit "no borders", which the
  // importer keys on (bordersSpecified) to suppress the renderer's gray default
  // grid. Dropping it reverts a borderless table to that grid on reopen.
  return el(tag, undefined, inner);
}

function cellXml(cell: TableCell, ctx: PartCtx, vMergeRestart = false): string {
  const pr: string[] = [];
  if (cell.colSpan && cell.colSpan > 1) pr.push(el("w:gridSpan", { "w:val": cell.colSpan }));
  // vMerge restart opens a vertical merge whose continue cells are synthesized in
  // the rows below (see tableXml). gridSpan precedes vMerge per CT_TcPr.
  if (vMergeRestart) pr.push(el("w:vMerge", { "w:val": "restart" }));
  if (cell.shading) pr.push(el("w:shd", { "w:val": "clear", "w:color": "auto", "w:fill": hex(cell.shading) }));
  if (cell.borders) pr.push(bordersXml("w:tcBorders", cell.borders));
  // Cell padding (w:tcMar). The importer resolves every side onto cell.margin, so
  // omitting it would re-import as Word's default (0 top/bottom) — silently
  // shrinking every row and drifting page count. Emit all four sides to round-trip.
  if (cell.margin) {
    const m = cell.margin;
    pr.push(el("w:tcMar", undefined,
      el("w:top", { "w:w": pxToTwips(m.top), "w:type": "dxa" }) +
      el("w:left", { "w:w": pxToTwips(m.left), "w:type": "dxa" }) +
      el("w:bottom", { "w:w": pxToTwips(m.bottom), "w:type": "dxa" }) +
      el("w:right", { "w:w": pxToTwips(m.right), "w:type": "dxa" }),
    ));
  }
  const tcPr = el("w:tcPr", undefined, pr.join(""));
  // A cell must contain at least one paragraph.
  const content = cell.blocks.length > 0 ? cell.blocks.map((b) => blockXml(b, ctx)).join("") : el("w:p");
  return el("w:tc", undefined, tcPr + content);
}

/** An active vertical merge, keyed by its START column: how many continue rows
 *  remain, how many grid columns it spans, and the owner's resolved borders. */
interface PendingMerge {
  remaining: number;
  span: number;
  borders: CellBorders | undefined;
}

/** A synthesized w:vMerge="continue" cell for one band of an open vertical merge.
 *  Word renders a merged cell's edges from its constituent cells, NOT the restart
 *  cell alone: the side borders repeat on every band, and the merged region's
 *  bottom is taken from the FINAL continue cell. We mirror the owner's borders
 *  onto each band so Word draws the same box the engine paints from the owner. */
function continueCellXml(m: PendingMerge): string {
  const pr: string[] = [];
  if (m.span > 1) pr.push(el("w:gridSpan", { "w:val": m.span })); // gridSpan precedes vMerge per CT_TcPr
  pr.push(el("w:vMerge", { "w:val": "continue" }));
  if (m.borders) {
    const isFinalBand = m.remaining === 1; // last continue row owns the merge's bottom edge
    const b: CellBorders = {};
    if (m.borders.left) b.left = m.borders.left;
    if (m.borders.right) b.right = m.borders.right;
    if (isFinalBand && m.borders.bottom) b.bottom = m.borders.bottom;
    if (b.left || b.right || b.bottom) pr.push(bordersXml("w:tcBorders", b));
  }
  return el("w:tc", undefined, el("w:tcPr", undefined, pr.join("")) + el("w:p"));
}

/** Build per-row cells, re-synthesizing the w:vMerge "continue" cells the
 *  importer dropped: a rowSpan=N owner needs N-1 continue cells stacked below it
 *  in the following rows. A continue cell carries the owner's gridSpan (so a
 *  colSpan+rowSpan merge stays one logical cell — emitting one continue per
 *  column would make the importer bump rowSpan once PER column) and the owner's
 *  borders (so Word draws the merged region's sides/bottom rather than its gray
 *  gridlines). */
function tableXml(table: TableBlock, ctx: PartCtx): string {
  const colCount = Math.max(
    1,
    ...table.rows.map((r) => r.cells.reduce((s, c) => s + (c.colSpan ?? 1), 0)),
  );
  const grid = (table.colFractions ?? Array.from({ length: colCount }, () => 1 / colCount))
    .map((f) => el("w:gridCol", { "w:w": Math.max(1, Math.round(f * 9000)) }))
    .join("");

  // Active vertical merges, indexed by their START column (covered columns stay
  // undefined — we emit one gridSpan'd continue cell and skip past them).
  const pending = new Array<PendingMerge | undefined>(colCount).fill(undefined);

  // Emit every continue cell occupying columns from `col` onward, stopping at a
  // free column. Returns the advanced column index.
  const emitContinues = (col: number, out: string[]): number => {
    while (col < colCount) {
      const m = pending[col];
      if (!m || m.remaining <= 0) break;
      out.push(continueCellXml(m));
      m.remaining--;
      if (m.remaining <= 0) pending[col] = undefined;
      col += m.span;
    }
    return col;
  };

  const rowsXml: string[] = [];
  for (const row of table.rows) {
    let col = 0;
    const out: string[] = [];
    for (const cell of row.cells) {
      // Continue cells for merges still spanning down from a row above.
      col = emitContinues(col, out);
      const span = cell.colSpan ?? 1;
      if (cell.rowSpan && cell.rowSpan > 1) {
        out.push(cellXml(cell, ctx, true));
        pending[col] = { remaining: cell.rowSpan - 1, span, borders: cell.borders };
      } else {
        out.push(cellXml(cell, ctx));
      }
      col += span;
    }
    // Trailing columns still under a span.
    col = emitContinues(col, out);
    rowsXml.push(el("w:tr", undefined, out.join("")));
  }

  // w:tblStyle (style reference) must precede the rest of tblPr; w:tblLook records
  // which conditional bands are active. Cells still carry concrete (baked) tcPr, so
  // the table renders correctly even without the style.
  const styleRef = table.styleId ? el("w:tblStyle", { "w:val": table.styleId }) : "";
  const look = table.styleId && table.condOverrides ? tblLookXml(table.condOverrides) : "";
  const tblPr = el(
    "w:tblPr",
    undefined,
    styleRef + el("w:tblW", { "w:w": 0, "w:type": "auto" }) + el("w:tblLayout", { "w:type": "fixed" }) + look,
  );
  return el("w:tbl", undefined, tblPr + el("w:tblGrid", undefined, grid) + rowsXml.join(""));
}

function tblLookXml(o: TableCondOverrides): string {
  return el("w:tblLook", {
    "w:firstRow": o.firstRow ? "1" : "0",
    "w:lastRow": o.lastRow ? "1" : "0",
    "w:firstColumn": o.firstCol ? "1" : "0",
    "w:lastColumn": o.lastCol ? "1" : "0",
    "w:noHBand": o.bandRows ? "0" : "1",
    "w:noVBand": o.bandCols ? "0" : "1",
  });
}

// ---------------------------------------------------------------------------
// Blocks / sections

export function blockXml(block: Block, ctx: PartCtx): string {
  if (block.kind === "paragraph") return paragraphXml(block, ctx);
  if (block.kind === "table") return tableXml(block, ctx);
  return imageParagraphXml(block, ctx);
}

// ---------------------------------------------------------------------------
// Live TOC field. A model TOC is a run of paragraphs carrying style.tocEntry
// (the editor's generated table of contents). On export we wrap them in a real
// Word `TOC` complex field and each entry's page number in a `PAGEREF` field —
// so the result is a live, F9-refreshable table of contents, not frozen text.

const fldRun = (type: "begin" | "separate" | "end"): string =>
  el("w:r", undefined, el("w:fldChar", { "w:fldCharType": type }));
const instrRun = (instr: string): string =>
  el("w:r", undefined, el("w:instrText", { "xml:space": "preserve" }, escapeText(instr)));

/** A `PAGEREF <anchor> \h` field whose cached result is `numText` (the page). */
function pagerefFieldXml(anchor: string, numText: string, rPr: string): string {
  return (
    fldRun("begin") +
    instrRun(` PAGEREF ${anchor} \\h `) +
    fldRun("separate") +
    el("w:r", undefined, rPr + (numText.length > 0 ? textEl(numText) : "")) +
    fldRun("end")
  );
}

const isTocEntry = (b: Block): b is Paragraph => b.kind === "paragraph" && !!b.style.tocEntry;

/** Emit a contiguous run of tocEntry paragraphs as one block-level `TOC` field.
 *  Each entry is "<label><tab><PAGEREF>"; the field's begin/instr/separate ride
 *  the first entry and its end rides the last. The instruction is the preserved
 *  one (so imported TOCs round-trip with their switches), honoring `\h` (wrap
 *  entries as hyperlinks) and `\n` (omit page numbers for levels). Headings get a
 *  bookmark (synthesized if needed) so the PAGEREF resolves. */
function tocFieldXml(entries: Paragraph[], ctx: PartCtx): string {
  const instr = ctx.tocInstruction ?? ' TOC \\o "1-3" \\h \\z \\u ';
  const sw = parseTocInstruction(instr);
  const hyperlink = sw.hyperlinks || ctx.tocInstruction === undefined; // default TOC is \h
  return entries
    .map((p, i) => {
      const rPr = p.runs[0] ? rPrXml(p.runs[0].style) : "";
      const targetId = p.style.tocEntry!.targetId;
      const level = p.style.tocEntry!.level;
      const wantNumber = !inRange(sw.hidePageNumberRange, level);
      const anchor = wantNumber || hyperlink ? ctx.ensureTocBookmark?.(targetId) : undefined;
      let inner = "";
      if (i === 0) inner += fldRun("begin") + instrRun(instr) + fldRun("separate");
      let entry = runsXml(p.runs, ctx); // the entry label
      if (anchor && wantNumber) {
        const numText = String(ctx.tocPages?.get(targetId) ?? "");
        entry += el("w:r", undefined, rPr + el("w:tab")) + pagerefFieldXml(anchor, numText, rPr);
      }
      if (hyperlink && anchor) entry = el("w:hyperlink", { "w:anchor": anchor, "w:history": "1" }, entry);
      inner += entry;
      if (i === entries.length - 1) inner += fldRun("end");
      return el("w:p", undefined, pPrXml(p.style, ctx) + inner);
    })
    .join("");
}

// ---------------------------------------------------------------------------
// Generic custom field. A run of contiguous blocks sharing a `fieldId` is the
// result region of a host-defined field (Document.fields). On export we wrap it
// in a real complex field — the begin/instrText/separate ride the FIRST result
// paragraph, the end rides the LAST — so it re-imports as the same field and the
// host can refresh it. The verbatim instruction is re-emitted unchanged.

/** Insert `frag` immediately after the paragraph's w:pPr (before its first run),
 *  so field begin/instr/separate lead the paragraph content. */
function injectAfterPPr(pXml: string, frag: string): string {
  const at = pXml.indexOf("</w:pPr>");
  if (at >= 0) return pXml.slice(0, at + 8) + frag + pXml.slice(at + 8);
  const m = pXml.match(/<w:p(?:\s[^>]*)?>/);
  if (m) { const i = m.index! + m[0].length; return pXml.slice(0, i) + frag + pXml.slice(i); }
  return frag + pXml;
}

/** Insert `frag` just before the paragraph's closing tag, so the field end is the
 *  paragraph's last run. */
function injectBeforePClose(pXml: string, frag: string): string {
  const at = pXml.lastIndexOf("</w:p>");
  return at >= 0 ? pXml.slice(0, at) + frag + pXml.slice(at) : pXml + frag;
}

function fieldRegionXml(blocks: Block[], ctx: PartCtx): string {
  const def = ctx.fields?.[blocks[0]!.fieldId!];
  const prefix = fldRun("begin") + instrRun(def?.instruction ?? " ") + fldRun("separate");
  const suffix = fldRun("end");
  const rendered = blocks.map((b) => blockXml(b, ctx));
  // The field boundary runs must live inside a paragraph; if the first/last result
  // block is a table/image, carry the boundary on a synthetic empty paragraph
  // (which re-imports as part of the field, so a later round-trip is stable).
  let head = "";
  let tail = "";
  if (blocks[0]!.kind === "paragraph") rendered[0] = injectAfterPPr(rendered[0]!, prefix);
  else head = el("w:p", undefined, prefix);
  const last = rendered.length - 1;
  if (blocks[last]!.kind === "paragraph") rendered[last] = injectBeforePClose(rendered[last]!, suffix);
  else tail = el("w:p", undefined, suffix);
  return head + rendered.join("") + tail;
}

function imageParagraphXml(img: Extract<Block, { kind: "image" }>, ctx: PartCtx): string {
  const target = ctx.media.resolve(img.src);
  if (!target) {
    ctx.warn("image-unresolved", img.src);
    return el("w:p");
  }
  const relId = ctx.rels.add(REL.image, target);
  const cx = pxToEmu(img.widthPx);
  const cy = pxToEmu(img.heightPx);
  const graphic = el(
    "a:graphic",
    undefined,
    el(
      "a:graphicData",
      { uri: "http://schemas.openxmlformats.org/drawingml/2006/picture" },
      el(
        "pic:pic",
        undefined,
        el(
          "pic:nvPicPr",
          undefined,
          el("pic:cNvPr", { id: 0, name: "image" }) + el("pic:cNvPicPr"),
        ) +
          el("pic:blipFill", undefined, el("a:blip", { "r:embed": relId }) + el("a:stretch", undefined, el("a:fillRect"))) +
          el(
            "pic:spPr",
            undefined,
            el("a:xfrm", undefined, el("a:off", { x: 0, y: 0 }) + el("a:ext", { cx, cy })) +
              el("a:prstGeom", { prst: "rect" }, el("a:avLst")),
          ),
      ),
    ),
  );

  // Anchored behind/in-front image (wrapNone): re-emit as a wp:anchor so the
  // absolute position and z-order survive the round-trip (inline would lose them).
  if (img.anchor) {
    const a = img.anchor;
    const anchor = el(
      "wp:anchor",
      {
        distT: 0, distB: 0, distL: 0, distR: 0, simplePos: 0,
        // relativeHeight is an unsigned int in OOXML; clamp the (possibly
        // negative, from repeated "send to back") z into the valid range.
        relativeHeight: Math.max(0, Math.round(a.z ?? 0)), behindDoc: a.behind ? 1 : 0, locked: 0,
        layoutInCell: 1, allowOverlap: 1,
      },
      el("wp:simplePos", { x: 0, y: 0 }) +
        el("wp:positionH", { relativeFrom: a.relFromH }, el("wp:posOffset", undefined, String(pxToEmu(a.offsetXPx)))) +
        el("wp:positionV", { relativeFrom: a.relFromV }, el("wp:posOffset", undefined, String(pxToEmu(a.offsetYPx)))) +
        el("wp:extent", { cx, cy }) +
        el("wp:wrapNone") +
        el("wp:docPr", { id: ctx.nextId(), name: "image" }) +
        graphic,
    );
    return el("w:p", undefined, el("w:r", undefined, el("w:drawing", undefined, anchor)));
  }

  const drawing = el(
    "w:drawing",
    undefined,
    el(
      "wp:inline",
      { distT: 0, distB: 0, distL: 0, distR: 0 },
      el("wp:extent", { cx, cy }) + el("wp:docPr", { id: ctx.nextId(), name: "image" }) + graphic,
    ),
  );
  const align = img.align === "center" ? "center" : img.align === "right" ? "right" : "left";
  return el("w:p", undefined, el("w:pPr", undefined, el("w:jc", { "w:val": align })) + el("w:r", undefined, drawing));
}

/** Band reference + part registration callback: returns the new relationship id. */
export type AddBandPart = (blocks: Block[], kind: "header" | "footer") => string;

export function sectPrXml(
  s: SectionProps | SectionPatch,
  ctx: PartCtx,
  addBand: AddBandPart,
  isBody: boolean,
): string {
  const c: string[] = [];
  // Header/footer references (default/first/even).
  const band = (blocks: Block[] | undefined, kind: "header" | "footer", type: string): void => {
    if (!blocks) return;
    const id = addBand(blocks, kind);
    c.push(el(`w:${kind}Reference`, { "w:type": type, "r:id": id }));
  };
  band(s.header, "header", "default");
  band(s.headerFirst, "header", "first");
  band(s.headerEven, "header", "even");
  band(s.footer, "footer", "default");
  band(s.footerFirst, "footer", "first");
  band(s.footerEven, "footer", "even");
  if (s.headerFirst || s.footerFirst) c.push(el("w:titlePg"));

  if (s.pageWidthPx !== undefined && s.pageHeightPx !== undefined) {
    c.push(el("w:pgSz", { "w:w": pxToTwips(s.pageWidthPx), "w:h": pxToTwips(s.pageHeightPx) }));
  }
  if (s.marginPx) {
    const m = s.marginPx;
    c.push(el("w:pgMar", {
      "w:top": pxToTwips(m.top), "w:right": pxToTwips(m.right),
      "w:bottom": pxToTwips(m.bottom), "w:left": pxToTwips(m.left),
      "w:header": 720, "w:footer": 720,
    }));
  }
  if (s.columns && s.columns.count > 1) {
    c.push(el("w:cols", { "w:num": s.columns.count, "w:space": pxToTwips(s.columns.gapPx) }));
  }
  if (s.pageNumberStart !== undefined) c.push(el("w:pgNumType", { "w:start": s.pageNumberStart }));
  // A non-body sectPr (on a paragraph) implies a Next Page break.
  if (!isBody) c.push(el("w:type", { "w:val": "nextPage" }));
  return el("w:sectPr", undefined, c.join(""));
}

export interface BuildResult {
  xml: string;
}

export function buildDocumentXml(
  blocks: Block[],
  section: SectionProps,
  ctx: PartCtx,
  addBand: AddBandPart,
): string {
  const parts: string[] = [];
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i]!;
    if (isTocEntry(b)) {
      // Wrap the whole contiguous run of TOC entries in one live TOC field.
      const group: Paragraph[] = [];
      while (i < blocks.length && isTocEntry(blocks[i]!)) group.push(blocks[i++] as Paragraph);
      i--; // the for-loop will ++ past the last consumed block
      parts.push(tocFieldXml(group, ctx));
    } else if (b.fieldId && ctx.fields?.[b.fieldId]) {
      // Wrap the contiguous run of blocks sharing this fieldId in a complex field.
      const id = b.fieldId;
      const group: Block[] = [];
      while (i < blocks.length && blocks[i]!.fieldId === id) group.push(blocks[i++]!);
      i--;
      parts.push(fieldRegionXml(group, ctx));
    } else {
      parts.push(blockXml(b, ctx));
    }
  }
  const body = parts.join("") + sectPrXml(section, ctx, addBand, true);
  return XML_DECL + el("w:document", WML_NS, el("w:body", undefined, body));
}
