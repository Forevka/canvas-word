// Milestone-1 demo document: mixed inline styles, alignments, a multi-page
// paragraph (line-level page break), and a CJK paragraph (Intl.Segmenter breaking).

import type { CharStyle, Document, ImageBlock, ParaStyle, Paragraph, Run, TableBlock } from "@cw/shared";
import { defaultStylesheet } from "@cw/shared";
import { defaultListDefinition, DEFAULT_BULLET_LIST_ID, DEFAULT_NUMBER_LIST_ID } from "@cw/shared";

const BODY: CharStyle = {
  fontFamily: "Georgia, serif",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#202124",
};

const PARA: ParaStyle = {
  align: "left",
  lineHeight: 1.5,
  spaceBeforePx: 0,
  spaceAfterPx: 12,
  indentFirstLinePx: 0,
  indentLeftPx: 0,
};

let nextId = 0;
const id = (): string => `b${nextId++}`;

const run = (text: string, patch: Partial<CharStyle> = {}): Run => ({
  text,
  style: { ...BODY, ...patch },
});

const para = (runs: Run[], patch: Partial<ParaStyle> = {}): Paragraph => ({
  kind: "paragraph",
  id: id(),
  revision: 0,
  runs,
  style: { ...PARA, ...patch },
});

const SAMPLE_SVG =
  "data:image/svg+xml," +
  encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" width="360" height="120">` +
      `<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1">` +
      `<stop offset="0" stop-color="#1a73e8"/><stop offset="1" stop-color="#9c27b0"/>` +
      `</linearGradient></defs>` +
      `<rect width="360" height="120" rx="8" fill="url(#g)"/>` +
      `<text x="180" y="68" font-family="Arial" font-size="22" fill="#fff" text-anchor="middle">an image block</text>` +
      `</svg>`,
  );

const image = (): ImageBlock => ({
  kind: "image",
  id: id(),
  revision: 0,
  src: SAMPLE_SVG,
  widthPx: 360,
  heightPx: 120,
  align: "center",
});

const cellPara = (text: string, patch: Partial<CharStyle> = {}): Paragraph => ({
  kind: "paragraph",
  id: id(),
  revision: 0,
  runs: [run(text, { fontSizePx: 14, ...patch })],
  style: { ...PARA, lineHeight: 1.35, spaceAfterPx: 0 },
});

const table = (): TableBlock => ({
  kind: "table",
  id: id(),
  revision: 0,
  rows: [
    {
      cells: [
        { id: id(), blocks: [cellPara("Layer", { bold: true })] },
        { id: id(), blocks: [cellPara("Responsibility", { bold: true })] },
        { id: id(), blocks: [cellPara("Measures text?", { bold: true })] },
      ],
    },
    {
      cells: [
        { id: id(), blocks: [cellPara("layout")] },
        { id: id(), blocks: [cellPara("pretext line breaking, pagination, geometry queries — the only module importing pretext")] },
        { id: id(), blocks: [cellPara("yes", { color: "#188038" })] },
      ],
    },
    {
      cells: [
        { id: id(), blocks: [cellPara("paint")] },
        { id: id(), blocks: [cellPara("virtualized canvases, one fillText per fragment")] },
        { id: id(), blocks: [cellPara("never", { color: "#a4373a", bold: true })] },
      ],
    },
  ],
});

const LOREM =
  "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor " +
  "incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud " +
  "exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute " +
  "irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla " +
  "pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia " +
  "deserunt mollit anim id est laborum. ";

export function sampleDoc(): Document {
  return {
    stylesheet: defaultStylesheet(),
    lists: {
      [DEFAULT_BULLET_LIST_ID]: defaultListDefinition("bullet"),
      [DEFAULT_NUMBER_LIST_ID]: defaultListDefinition("decimal"),
    },
    section: {
      // US Letter at 96 dpi, 1in margins
      pageWidthPx: 816,
      pageHeightPx: 1056,
      marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
      // Margin bands are full block stories — any block kind works here.
      header: [
        para(
          [
            run("canvas-word", { fontFamily: "Arial, sans-serif", fontSizePx: 11, bold: true, color: "#5f6368" }),
            run("  ·  pretext-on-canvas demo", { fontFamily: "Arial, sans-serif", fontSizePx: 11, color: "#9aa0a6" }),
          ],
          { spaceAfterPx: 0 },
        ),
      ],
      footer: [
        para(
          [run("Page {page} of {pages}", { fontFamily: "Arial, sans-serif", fontSizePx: 11, color: "#9aa0a6" })],
          { align: "center", spaceAfterPx: 0 },
        ),
      ],
    },
    blocks: [
      para(
        [run("canvas-word", { fontFamily: "Arial, sans-serif", fontSizePx: 32, bold: true, color: "#1a1a2e" })],
        { align: "center", spaceAfterPx: 4, namedStyle: "Title" },
      ),
      para(
        [run("a page-accurate document editor rendered with pretext on canvas", { italic: true, color: "#5f6368" })],
        { align: "center", spaceAfterPx: 28, namedStyle: "Subtitle" },
      ),
      para([
        run("Every line on this page was broken by "),
        run("@chenglou/pretext", { fontFamily: "Consolas, monospace", fontSizePx: 15, color: "#0b57d0" }),
        run(", not by the browser. The paint layer below this paragraph never measures text — it receives absolutely-positioned fragments and calls "),
        run("fillText", { fontFamily: "Consolas, monospace", fontSizePx: 15 }),
        run(" once per same-styled slice. That separation is the whole architecture."),
      ]),
      para([
        run("Inline styles flow through a single rich-inline prepare pass: "),
        run("bold", { bold: true }),
        run(", "),
        run("italic", { italic: true }),
        run(", "),
        run("underlined", { underline: true }),
        run(", "),
        run("struck through", { strikethrough: true }),
        run(", and "),
        run("mixed sizes inside one paragraph", { fontSizePx: 22, bold: true, italic: true, color: "#a4373a" }),
        run(" — all measured once, line-broken together, cached by paragraph revision."),
      ]),
      image(),
      para(
        [run("First-line indents fall out of pretext's variable-width API: the engine simply passes a narrower maxWidth for the first line of this paragraph. " + LOREM)],
        { indentFirstLinePx: 36 },
      ),
      table(),
      para([run("The layers, as a multilevel numbered list:")], { spaceAfterPx: 4 }),
      para([run("Model — invertible ops, position mapping")], { list: { listId: DEFAULT_NUMBER_LIST_ID, level: 0 }, spaceAfterPx: 2 }),
      para([run("Layout — pretext, pagination, geometry")], { list: { listId: DEFAULT_NUMBER_LIST_ID, level: 0 }, spaceAfterPx: 2 }),
      para([run("line caches keyed by (revision, width)")], { list: { listId: DEFAULT_NUMBER_LIST_ID, level: 1 }, spaceAfterPx: 2 }),
      para([run("break rules: widows, orphans, keep-with-next")], { list: { listId: DEFAULT_NUMBER_LIST_ID, level: 1 }, spaceAfterPx: 2 }),
      para([run("counters reset when a higher level increments")], { list: { listId: DEFAULT_NUMBER_LIST_ID, level: 2 }, spaceAfterPx: 2 }),
      para([run("Paint — never measures, one fillText per fragment")], { list: { listId: DEFAULT_NUMBER_LIST_ID, level: 0 }, spaceAfterPx: 2 }),
      para([run("And bullets:")], { spaceBeforePx: 8, spaceAfterPx: 4 }),
      para([run("markers are paint-only — they live in the hanging indent")], { list: { listId: DEFAULT_BULLET_LIST_ID, level: 0 }, spaceAfterPx: 2 }),
      para([run("so the line-break caches survive renumbering")], { list: { listId: DEFAULT_BULLET_LIST_ID, level: 1 }, spaceAfterPx: 2 }),
      para(
        [run("This paragraph is centered.", { fontSizePx: 18 })],
        { align: "center" },
      ),
      para(
        [run("And this one is right-aligned.", { fontSizePx: 18 })],
        { align: "right" },
      ),
      para([
        run("複雑なスクリプトの行分割も Intl.Segmenter が処理します。日本語のテキストは単語間にスペースがありませんが、pretext は文節の境界を正しく検出して行を折り返します。"),
        run("中文文本同样可以正确换行，因为分词由浏览器内置的国际化引擎完成。", { color: "#0b57d0" }),
      ]),
      // Long paragraph: forces a line-level page break (spans pages mid-paragraph),
      // and is JUSTIFIED — slack is distributed into spaces via ctx.wordSpacing.
      para(
        [run("The paragraph you are reading now is long enough to cross a page boundary, which exercises line-level pagination — the thing contenteditable-based editors cannot do honestly. It is also fully justified: every line except the last is stretched to the margin by distributing leftover width across its spaces. " + LOREM.repeat(14))],
        { align: "justify" },
      ),
      para([
        run("Page two continues seamlessly because pagination operates on line boxes, not on blocks. ", { bold: true }),
        run("Heights of every page are known before anything is painted, so the scrollbar is exact and virtualization mounts only the canvases that intersect the viewport."),
      ]),
      para([run(LOREM.repeat(6))]),
      para(
        [run("— end of milestone 1 —", { italic: true, color: "#5f6368" })],
        { align: "center", spaceBeforePx: 24 },
      ),
    ],
  };
}
