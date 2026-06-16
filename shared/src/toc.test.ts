import { describe, expect, it } from "vitest";
import { buildTocInstruction, parseTocInstruction, type TocSwitches } from "./toc";

describe("TOC switches ⇄ instruction round-trip", () => {
  const cases: { name: string; sw: TocSwitches }[] = [
    { name: "level range + hyperlinks + web-hide", sw: { outlineRange: { from: 1, to: 3 }, useOutlineLevels: false, hyperlinks: true, hideInWeb: true } },
    { name: "wide range, outline levels", sw: { outlineRange: { from: 1, to: 5 }, useOutlineLevels: true, hyperlinks: false, hideInWeb: false } },
    { name: "separator + hide-numbers range", sw: { outlineRange: { from: 1, to: 2 }, useOutlineLevels: false, hyperlinks: true, hideInWeb: false, separator: "—", hidePageNumberRange: { from: 3, to: 4 } } },
    { name: "no range (flags only)", sw: { useOutlineLevels: true, hyperlinks: true, hideInWeb: false } },
  ];

  for (const c of cases) {
    it(`re-parses to the same switches: ${c.name}`, () => {
      expect(parseTocInstruction(buildTocInstruction(c.sw))).toEqual(c.sw);
    });
  }

  it("builds a Word-shaped instruction string", () => {
    expect(buildTocInstruction({ outlineRange: { from: 1, to: 3 }, useOutlineLevels: false, hyperlinks: true, hideInWeb: true }))
      .toBe(' TOC \\o "1-3" \\h \\z ');
  });
});
