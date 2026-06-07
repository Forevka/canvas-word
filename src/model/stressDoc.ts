// Milestone-2 perf probe: a ~N-page document with a keep-with-next heading every
// 10 paragraphs, so the break-rule scan in the browser has real material to check.

import type { CharStyle, Document, ParaStyle, Paragraph, Run } from "./document";

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

const LOREM =
  "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor " +
  "incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud " +
  "exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute " +
  "irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla " +
  "pariatur. ";

export function stressDoc(targetPages: number): Document {
  const blocks: Paragraph[] = [];
  // ~36 content lines per page; body paragraph below is ~7 lines + spacing → ~4.5
  // paragraphs per page. Vary length per index so line counts differ honestly.
  const paragraphCount = Math.max(1, Math.round(targetPages * 4.5));

  for (let i = 0; i < paragraphCount; i++) {
    if (i % 10 === 0) {
      blocks.push({
        kind: "paragraph",
        id: `h${i}`,
        revision: 0,
        runs: [
          {
            text: `${1 + i / 10}. Section heading ${1 + i / 10}`,
            style: { ...BODY, fontFamily: "Arial, sans-serif", fontSizePx: 22, bold: true },
          },
        ],
        style: { ...PARA, spaceBeforePx: 18, spaceAfterPx: 8, keepWithNext: true },
      });
    }
    const runs: Run[] = [
      { text: `¶${i} — `, style: { ...BODY, bold: true, color: "#0b57d0" } },
      { text: LOREM.repeat(1 + (i % 3)), style: BODY },
    ];
    blocks.push({ kind: "paragraph", id: `p${i}`, revision: 0, runs, style: PARA });
  }

  return {
    section: {
      pageWidthPx: 816,
      pageHeightPx: 1056,
      marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
    },
    blocks,
  };
}
