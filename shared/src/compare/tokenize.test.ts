import { describe, expect, it } from "vitest";
import { tokenize } from "./tokenize";

describe("tokenize — word granularity", () => {
  it("splits words, whitespace, and punctuation with correct offsets", () => {
    const toks = tokenize("hi, world");
    expect(toks.map((t) => t.text)).toEqual(["hi", ",", " ", "world"]);
    // offsets are contiguous and cover the whole string
    expect(toks[0]).toMatchObject({ start: 0, end: 2 });
    expect(toks[3]).toMatchObject({ start: 4, end: 9 });
    for (let i = 1; i < toks.length; i++) expect(toks[i]!.start).toBe(toks[i - 1]!.end);
  });

  it("reconstructs the original string from token slices", () => {
    const text = "The  quick—brown fox.";
    const toks = tokenize(text);
    expect(toks.map((t) => text.slice(t.start, t.end)).join("")).toBe(text);
  });

  it("empty string yields no tokens", () => {
    expect(tokenize("")).toEqual([]);
  });
});

describe("tokenize — char granularity", () => {
  it("one token per code point with UTF-16 offsets", () => {
    const toks = tokenize("ab", "char");
    expect(toks.map((t) => t.text)).toEqual(["a", "b"]);
  });

  it("keeps a surrogate pair as one token spanning two code units", () => {
    const toks = tokenize("a😀b", "char");
    expect(toks.map((t) => t.text)).toEqual(["a", "😀", "b"]);
    expect(toks[1]).toMatchObject({ start: 1, end: 3 }); // emoji is 2 UTF-16 units
    expect(toks[2]).toMatchObject({ start: 3, end: 4 });
  });
});
