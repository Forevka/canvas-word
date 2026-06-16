// Clipboard bridge.
// Copy/cut: serialize the selected model fragment to text/html + text/plain —
// the HTML flavor is what makes paste into real Word / Google Docs keep styles.
// Paste: whitelist-map incoming HTML onto runs + paragraph styles via DOMParser.

import type { Block, CharStyle, GridRect, ParaStyle, Paragraph, Run, TableBlock } from "@cw/shared";
import type { DocPosition } from "@cw/shared";
import { buildTableGrid, normalizeRect, normalizeRuns, sliceRuns } from "@cw/shared";
import { textOfRuns } from "@cw/shared";

export interface FragmentBlock {
  runs: Run[];
  style: ParaStyle;
}

export interface DocFragment {
  blocks: FragmentBlock[];
  /** True for a partial single paragraph (inline paste, no paragraph breaks). */
  inline: boolean;
}

const DEFAULT_CHAR: CharStyle = {
  fontFamily: "Georgia, serif",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#202124",
};

const DEFAULT_PARA: ParaStyle = {
  align: "left",
  lineHeight: 1.5,
  spaceBeforePx: 0,
  spaceAfterPx: 12,
  indentFirstLinePx: 0,
  indentLeftPx: 0,
};

// ---------------------------------------------------------------------------
// Extraction

/** `paras` is the story's paragraph list (body incl. cells, or one band). */
export function extractFragment(paras: Paragraph[], from: DocPosition, to: DocPosition): DocFragment {
  const blocks = paras;
  const fi = paras.findIndex((b) => b.id === from.blockId);
  const ti = paras.findIndex((b) => b.id === to.blockId);
  const out: FragmentBlock[] = [];
  for (let i = fi; i <= ti; i++) {
    const block = blocks[i]!;
    const len = textOfRuns(block.runs).length;
    const start = i === fi ? from.offset : 0;
    const end = i === ti ? to.offset : len;
    out.push({
      runs: sliceRuns(block.runs, start, end),
      style: { ...block.style },
    });
  }
  return { blocks: out, inline: out.length === 1 };
}

// ---------------------------------------------------------------------------
// Serialization (export)

const escapeHtml = (s: string): string =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

function runToHtml(run: Run): string {
  const s = run.style;
  const css: string[] = [
    `font-family:${s.fontFamily}`,
    `font-size:${s.fontSizePx}px`,
    `color:${s.color}`,
  ];
  if (s.bold) css.push("font-weight:700");
  if (s.italic) css.push("font-style:italic");
  if (s.highlightColor) css.push(`background-color:${s.highlightColor}`);
  if (s.verticalAlign) css.push(`vertical-align:${s.verticalAlign}`);
  const deco: string[] = [];
  if (s.underline) deco.push("underline");
  if (s.strikethrough) deco.push("line-through");
  if (deco.length) css.push(`text-decoration:${deco.join(" ")}`);
  // Soft line breaks ("\v") become <br> inside the span.
  const inner = run.text.split("\v").map(escapeHtml).join("<br>");
  const span = `<span style="${css.join(";")}">${inner}</span>`;
  return s.link ? `<a href="${escapeHtml(s.link)}">${span}</a>` : span;
}

export function fragmentToHtml(fragment: DocFragment): string {
  const paras = fragment.blocks.map((b) => {
    const css = `text-align:${b.style.align};line-height:${b.style.lineHeight}`;
    return `<p style="${css}">${b.runs.map(runToHtml).join("")}</p>`;
  });
  return paras.join("");
}

export function fragmentToPlainText(fragment: DocFragment): string {
  // Soft breaks ("\v") read as newlines in plain text.
  return fragment.blocks.map((b) => textOfRuns(b.runs).replace(/\v/g, "\n")).join("\n");
}

// --- rectangular table-cell selection --------------------------------------
// A whole-cell / multi-cell copy serializes to an HTML <table> (so paste into
// Word / Google Docs / Sheets keeps the grid) plus tab-separated plain text
// (rows on their own line — the de-facto format spreadsheets parse into cells).

const cellParas = (blocks: Block[]): FragmentBlock[] =>
  blocks.filter((b): b is Paragraph => b.kind === "paragraph").map((p) => ({ runs: p.runs, style: p.style }));

function cellToHtml(blocks: Block[]): string {
  return fragmentToHtml({ blocks: cellParas(blocks), inline: false });
}

function cellToPlainText(blocks: Block[]): string {
  // Multiple paragraphs in one cell collapse to spaces so each table row stays a
  // single tab-delimited line (keeps columns aligned when pasted into a sheet).
  return cellParas(blocks)
    .map((b) => textOfRuns(b.runs).replace(/\v/g, " "))
    .join(" ")
    .trim();
}

/** Serialize a (possibly merged-cell-spanning) grid rectangle of a table to the
 *  text/html + text/plain clipboard flavors. `rect` is the raw anchor/focus
 *  rectangle; it is normalized to whole merged cells here. */
export function tableRectToClipboard(table: TableBlock, rect: GridRect): { html: string; text: string } {
  const grid = buildTableGrid(table);
  const r = normalizeRect(grid, rect);
  const trs: string[] = [];
  const lines: string[] = [];
  for (let ri = r.r0; ri <= r.r1; ri++) {
    const tds: string[] = [];
    const cells: string[] = [];
    for (let ci = r.c0; ci <= r.c1; ci++) {
      const s = grid.slots[ri]![ci];
      // Emit each cell once, at its grid origin; continuation slots of a
      // colSpan/rowSpan cell contribute nothing (HTML) / an empty column (text).
      const atOrigin = s !== undefined && s.originRow === ri && s.originCol === ci;
      if (atOrigin) {
        const cspan = s.colSpan > 1 ? ` colspan="${s.colSpan}"` : "";
        const rspan = s.rowSpan > 1 ? ` rowspan="${s.rowSpan}"` : "";
        tds.push(`<td${cspan}${rspan}>${cellToHtml(s.cell.blocks)}</td>`);
        cells.push(cellToPlainText(s.cell.blocks));
      } else if (!s) {
        tds.push("<td></td>"); // ragged hole
        cells.push("");
      } else {
        cells.push(""); // continuation slot: blank column keeps tab alignment
      }
    }
    trs.push(`<tr>${tds.join("")}</tr>`);
    lines.push(cells.join("\t"));
  }
  return { html: `<table>${trs.join("")}</table>`, text: lines.join("\n") };
}

// ---------------------------------------------------------------------------
// Parsing (import) — whitelist mapping, everything else degrades to plain text

const BLOCK_TAGS = new Set(["P", "DIV", "H1", "H2", "H3", "H4", "H5", "H6", "LI", "BLOCKQUOTE", "TR", "PRE"]);

function parseInlineStyle(el: HTMLElement, base: CharStyle): CharStyle {
  const out = { ...base };
  const tag = el.tagName;
  if (tag === "B" || tag === "STRONG") out.bold = true;
  if (tag === "I" || tag === "EM") out.italic = true;
  if (tag === "U") out.underline = true;
  if (tag === "S" || tag === "STRIKE" || tag === "DEL") out.strikethrough = true;
  if (tag === "CODE") out.fontFamily = "Consolas, monospace";
  if (tag === "A") {
    const href = el.getAttribute("href");
    if (href) out.link = href;
  }
  if (tag === "SUP") out.verticalAlign = "super";
  if (tag === "SUB") out.verticalAlign = "sub";
  if (tag === "MARK") out.highlightColor = "#ffeb3b";
  const st = el.style;
  if (st.backgroundColor && st.backgroundColor !== "transparent") out.highlightColor = st.backgroundColor;
  if (st.fontWeight === "bold" || Number(st.fontWeight) >= 600) out.bold = true;
  if (st.fontWeight === "normal" || (st.fontWeight && Number(st.fontWeight) < 600 && Number(st.fontWeight) > 0)) out.bold = false;
  if (st.fontStyle === "italic") out.italic = true;
  if (st.textDecoration.includes("underline") || st.textDecorationLine.includes("underline")) out.underline = true;
  if (st.textDecoration.includes("line-through") || st.textDecorationLine.includes("line-through")) out.strikethrough = true;
  if (st.color) out.color = st.color;
  if (st.fontFamily) out.fontFamily = st.fontFamily;
  const size = Number.parseFloat(st.fontSize);
  if (st.fontSize.endsWith("px") && Number.isFinite(size)) out.fontSizePx = Math.min(96, Math.max(6, size));
  if (tag.match(/^H[1-6]$/)) {
    out.bold = true;
    out.fontSizePx = [32, 26, 22, 19, 17, 16][Number(tag[1]) - 1]!;
  }
  return out;
}

export function htmlToFragment(html: string): DocFragment | null {
  const parsed = new DOMParser().parseFromString(html, "text/html");
  const blocks: FragmentBlock[] = [];
  let current: Run[] = [];
  let currentAlign: ParaStyle["align"] = "left";

  const flush = (): void => {
    if (current.length === 0) return;
    blocks.push({
      runs: normalizeRuns(current, DEFAULT_CHAR),
      style: { ...DEFAULT_PARA, align: currentAlign },
    });
    current = [];
    currentAlign = "left";
  };

  const walk = (node: Node, style: CharStyle): void => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = (node.textContent ?? "").replace(/\s+/g, " ");
      if (text.length > 0 && text !== " ") current.push({ text, style });
      else if (text === " " && current.length > 0) current.push({ text, style });
      return;
    }
    if (!(node instanceof HTMLElement)) return;
    if (node.tagName === "BR") {
      // <br> = soft line break within the paragraph (matches our "\v" export).
      current.push({ text: "\v", style });
      return;
    }
    if (node.tagName === "STYLE" || node.tagName === "SCRIPT" || node.tagName === "HEAD") return;
    const isBlock = BLOCK_TAGS.has(node.tagName);
    if (isBlock) {
      flush();
      const ta = node.style.textAlign;
      if (ta === "center" || ta === "right" || ta === "justify") currentAlign = ta;
    }
    const childStyle = parseInlineStyle(node, style);
    for (const child of Array.from(node.childNodes)) walk(child, childStyle);
    if (isBlock) flush();
  };

  walk(parsed.body, DEFAULT_CHAR);
  flush();
  if (blocks.length === 0) return null;
  return { blocks, inline: blocks.length === 1 };
}
