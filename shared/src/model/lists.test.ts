import { describe, expect, it } from "vitest";
import { bulletListDefinition, markerText, numberListDefinition } from "./lists";

describe("bulletListDefinition", () => {
  it("uses the given glyph at every level", () => {
    const def = bulletListDefinition("b", "▪");
    expect(def.id).toBe("b");
    expect(def.levels).toHaveLength(9);
    expect(markerText(def, 0, [1])).toBe("▪");
    expect(markerText(def, 3, [1, 1, 1, 1])).toBe("▪");
  });
});

describe("numberListDefinition", () => {
  it("renders the chosen format and suffix", () => {
    expect(markerText(numberListDefinition("n", "decimal", ")"), 0, [3])).toBe("3)");
    expect(markerText(numberListDefinition("n", "lowerLetter", "."), 0, [1])).toBe("a.");
    expect(markerText(numberListDefinition("n", "upperLetter", "."), 0, [2])).toBe("B.");
    expect(markerText(numberListDefinition("n", "lowerRoman", "."), 0, [4])).toBe("iv.");
    expect(markerText(numberListDefinition("n", "upperRoman", "."), 0, [9])).toBe("IX.");
  });
});
