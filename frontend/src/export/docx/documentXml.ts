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
  Paragraph,
  ParaStyle,
  Run,
  SdtProps,
  SectionPatch,
  SectionProps,
  TableBlock,
  TableCell,
} from "@cw/shared";
import {
  multiplierToLine,
  pxToEighthPoints,
  pxToEmu,
  pxToHalfPoints,
  pxToTwips,
} from "../units";
import { MediaManager } from "./mediaPack";
import { REL, RelManager } from "./relationships";
import { el, textEl, WML_NS, XML_DECL } from "./xmlWrite";

// model verticalAlign -> w:vertAlign; jc; highlight hex -> Word color name.
const HIGHLIGHT_NAME: Record<string, string> = {
  "#ffff00": "yellow", "#00ff00": "green", "#00ffff": "cyan", "#ff00ff": "magenta",
  "#0000ff": "blue", "#ff0000": "red", "#000080": "darkBlue", "#008080": "darkCyan",
  "#008000": "darkGreen", "#800080": "darkMagenta", "#800000": "darkRed", "#808000": "darkYellow",
  "#808080": "darkGray", "#c0c0c0": "lightGray", "#000000": "black", "#ffffff": "white",
};
const JC: Record<ParaStyle["align"], string> = { left: "left", center: "center", right: "right", justify: "both" };
const TAB_VAL: Record<string, string> = { left: "left", center: "center", right: "right", decimal: "decimal" };
const TAB_LEADER: Record<string, string> = { dot: "dot", dash: "hyphen", underscore: "underscore" };

export interface PartCtx {
  rels: RelManager;
  media: MediaManager;
  warn: (code: string, detail?: string) => void;
  sdts: Record<string, SdtProps>;
  /** Doc-wide unique ids for w:bookmarkStart / w:footnoteReference fallbacks. */
  nextId: () => number;
  /** model blockId -> bookmark names anchored there (body only). */
  bookmarksByBlock: Map<string, string[]>;
  /** model list id -> Word-valid integer numId (shared with numbering.xml). */
  listIdMap: Map<string, number>;
}

const hex = (color: string): string => color.replace(/^#/, "").toLowerCase();

// ---------------------------------------------------------------------------
// Runs

function rPrXml(s: CharStyle): string {
  const family = (s.fontFamily.split(",")[0] ?? "").trim();
  const children: string[] = [];
  if (family) children.push(el("w:rFonts", { "w:ascii": family, "w:hAnsi": family, "w:cs": family }));
  // Toggles explicit on/off so they fully override any inherited style.
  children.push(el("w:b", { "w:val": s.bold ? "1" : "0" }));
  children.push(el("w:i", { "w:val": s.italic ? "1" : "0" }));
  if (s.strikethrough) children.push(el("w:strike", { "w:val": "1" }));
  if (s.hidden) children.push(el("w:vanish", { "w:val": "1" })); // preserved hidden text
  children.push(el("w:color", { "w:val": hex(s.color) }));
  children.push(el("w:sz", { "w:val": pxToHalfPoints(s.fontSizePx) }));
  children.push(el("w:szCs", { "w:val": pxToHalfPoints(s.fontSizePx) }));
  children.push(el("w:u", { "w:val": s.underline ? "single" : "none" }));
  if (s.letterSpacingPx !== undefined) children.push(el("w:spacing", { "w:val": pxToTwips(s.letterSpacingPx) }));
  if (s.highlightColor) {
    const name = HIGHLIGHT_NAME[s.highlightColor.toLowerCase()];
    if (name) children.push(el("w:highlight", { "w:val": name }));
  }
  if (s.verticalAlign) children.push(el("w:vertAlign", { "w:val": s.verticalAlign === "super" ? "superscript" : "subscript" }));
  return el("w:rPr", undefined, children.join(""));
}

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

function singleRun(run: Run, ctx: PartCtx): string {
  const s = run.style;
  let body: string;
  if (s.footnoteRef) {
    // model footnoteRef "fn<docxId>" -> w:footnoteReference w:id="<docxId>".
    const docxId = s.footnoteRef.replace(/^fn/, "");
    body = el("w:footnoteReference", { "w:id": docxId });
  } else {
    body = runContent(run.text);
  }
  const r = el("w:r", undefined, rPrXml(s) + body);
  // Hyperlink wrapper. Anchor (#name) is in-document; otherwise an external rel.
  if (s.link) {
    if (s.link.startsWith("#")) return el("w:hyperlink", { "w:anchor": s.link.slice(1) }, r);
    const id = ctx.rels.add(REL.hyperlink, s.link, true);
    return el("w:hyperlink", { "r:id": id }, r);
  }
  return r;
}

/** Serialize a run list, grouping consecutive runs that share an sdtId into one
 *  w:sdt content control. */
function runsXml(runs: Run[], ctx: PartCtx): string {
  let out = "";
  let i = 0;
  while (i < runs.length) {
    const sdtId = runs[i]!.style.sdtId;
    if (sdtId) {
      let j = i;
      let inner = "";
      while (j < runs.length && runs[j]!.style.sdtId === sdtId) inner += singleRun(runs[j++]!, ctx);
      out += el("w:sdt", undefined, sdtPrXml(ctx.sdts[sdtId]) + el("w:sdtContent", undefined, inner));
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

  // w:spacing folds before/after/line into one element.
  const sp: Record<string, number | string> = {
    "w:before": pxToTwips(style.spaceBeforePx),
    "w:after": pxToTwips(style.spaceAfterPx),
    "w:line": multiplierToLine(style.lineHeight),
    "w:lineRule": "auto",
  };
  c.push(el("w:spacing", sp));

  // indent: hanging is a negative first-line indent.
  const ind: Record<string, number> = {};
  if (style.indentLeftPx) ind["w:left"] = pxToTwips(style.indentLeftPx);
  if (style.indentRightPx) ind["w:right"] = pxToTwips(style.indentRightPx);
  if (style.indentFirstLinePx > 0) ind["w:firstLine"] = pxToTwips(style.indentFirstLinePx);
  else if (style.indentFirstLinePx < 0) ind["w:hanging"] = pxToTwips(-style.indentFirstLinePx);
  if (Object.keys(ind).length > 0) c.push(el("w:ind", ind));

  c.push(el("w:jc", { "w:val": JC[style.align] }));

  if (style.tabStops && style.tabStops.length > 0) {
    const tabs = style.tabStops
      .map((t) =>
        el("w:tab", {
          "w:val": t.align ? (TAB_VAL[t.align] ?? "left") : "left",
          "w:pos": pxToTwips(t.posPx),
          "w:leader": t.leader && t.leader !== "none" ? TAB_LEADER[t.leader] : undefined,
        }),
      )
      .join("");
    c.push(el("w:tabs", undefined, tabs));
  }

  // The paragraph mark's run props size empty lines on re-import.
  if (markRun) c.push(rPrXml(markRun));

  // A mid-document section break rides the paragraph that ENDS the section.
  if (style.sectionBreak) c.push(sectPrXml(style.sectionBreak.props, ctx, () => "", false));

  return el("w:pPr", undefined, c.join(""));
}

function paragraphXml(p: Paragraph, ctx: PartCtx): string {
  const isEmpty = p.runs.every((r) => r.text.length === 0) && !p.runs.some((r) => r.style.footnoteRef);
  const markRun = isEmpty && p.runs[0] ? p.runs[0].style : undefined;
  const body = pPrXml(p.style, ctx, markRun) + runsXml(isEmpty ? [] : p.runs, ctx);
  const para = el("w:p", undefined, body);
  // Bookmarks anchored to this block bracket the paragraph.
  const names = ctx.bookmarksByBlock.get(p.id);
  if (names && names.length > 0) {
    let pre = "";
    let post = "";
    for (const name of names) {
      const id = ctx.nextId();
      pre += el("w:bookmarkStart", { "w:id": id, "w:name": name });
      post += el("w:bookmarkEnd", { "w:id": id });
    }
    return pre + para + post;
  }
  return para;
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

function cellXml(cell: TableCell, ctx: PartCtx): string {
  const pr: string[] = [];
  if (cell.colSpan && cell.colSpan > 1) pr.push(el("w:gridSpan", { "w:val": cell.colSpan }));
  if (cell.shading) pr.push(el("w:shd", { "w:val": "clear", "w:color": "auto", "w:fill": hex(cell.shading) }));
  if (cell.borders) pr.push(bordersXml("w:tcBorders", cell.borders));
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
        out.push(injectVMergeRestart(cellXml(cell, ctx)));
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

  const tblPr = el(
    "w:tblPr",
    undefined,
    el("w:tblW", { "w:w": 0, "w:type": "auto" }) + el("w:tblLayout", { "w:type": "fixed" }),
  );
  return el("w:tbl", undefined, tblPr + el("w:tblGrid", undefined, grid) + rowsXml.join(""));
}

/** Splice a w:vMerge restart into a cell's existing w:tcPr (or add one). */
function injectVMergeRestart(tc: string): string {
  const merge = el("w:vMerge", { "w:val": "restart" });
  if (tc.startsWith("<w:tc><w:tcPr>")) {
    return tc.replace("<w:tc><w:tcPr>", `<w:tc><w:tcPr>${merge}`);
  }
  if (tc.startsWith("<w:tc><w:tcPr/>")) {
    return tc.replace("<w:tc><w:tcPr/>", `<w:tc><w:tcPr>${merge}</w:tcPr>`);
  }
  return tc.replace("<w:tc>", `<w:tc>${el("w:tcPr", undefined, merge)}`);
}

// ---------------------------------------------------------------------------
// Blocks / sections

export function blockXml(block: Block, ctx: PartCtx): string {
  if (block.kind === "paragraph") return paragraphXml(block, ctx);
  if (block.kind === "table") return tableXml(block, ctx);
  return imageParagraphXml(block, ctx);
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
  const drawing = el(
    "w:drawing",
    undefined,
    el(
      "wp:inline",
      { distT: 0, distB: 0, distL: 0, distR: 0 },
      el("wp:extent", { cx, cy }) +
        el("wp:docPr", { id: ctx.nextId(), name: "image" }) +
        el(
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
        ),
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
  const body = blocks.map((b) => blockXml(b, ctx)).join("") + sectPrXml(section, ctx, addBand, true);
  return XML_DECL + el("w:document", WML_NS, el("w:body", undefined, body));
}
