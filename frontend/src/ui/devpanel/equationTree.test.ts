// The document inspector (dev panel) exposes equations as a "∑ equation" node
// that expands into the MathML AST — for both display blocks and inline runs.

import { describe, expect, it } from "vitest";
import type { CharStyle, Document, EquationBlock, MathEquation, Paragraph } from "@cw/shared";
import { buildModelTree } from "./modelTree";

const SECTION = { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } };
const CS: CharStyle = { fontFamily: "x", fontSizePx: 12, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };

// 1/2 fraction.
const frac: MathEquation = {
  display: true,
  root: { type: "row", children: [{ type: "frac", num: { type: "number", text: "1" }, den: { type: "number", text: "2" } }] },
};

const flatten = (nodes: ReturnType<typeof buildModelTree>): { label: string; kind: string }[] => {
  const out: { label: string; kind: string }[] = [];
  const walk = (ns: typeof nodes): void => {
    for (const n of ns) { out.push({ label: n.label, kind: n.kind }); walk(n.children); }
  };
  walk(nodes);
  return out;
};

describe("inspector equation nodes", () => {
  it("shows a display equation as a non-table node expanding into its AST", () => {
    const block: EquationBlock = { kind: "equation", id: "eq", revision: 0, align: "center", equation: frac };
    const tree = buildModelTree({ section: SECTION, blocks: [block] } as Document);
    const labels = flatten(tree).map((n) => n.label);
    expect(labels).toContain("∑ equation");
    expect(labels).not.toContain("▦ table"); // the bug: equation was treated as a table
    expect(labels).toContain("frac"); // AST is inspectable
    expect(labels).toContain('number "1"');
    expect(labels).toContain('number "2"');
  });

  it("shows an inline equation run as ∑ equation (inline) with its AST", () => {
    const para: Paragraph = {
      kind: "paragraph", id: "p", revision: 0,
      style: { align: "left", lineHeight: 1, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 },
      runs: [
        { text: "x", style: CS },
        { text: "￼", style: { ...CS, equation: { ...frac, display: false } } },
      ],
    };
    const labels = flatten(buildModelTree({ section: SECTION, blocks: [para] } as Document)).map((n) => n.label);
    expect(labels).toContain("∑ equation (inline)");
    expect(labels).toContain("frac");
  });
});
