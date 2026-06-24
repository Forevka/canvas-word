import { describe, expect, it } from "vitest";
import type { Block, CharStyle, Paragraph } from "./document";
import {
  ancestryThrough,
  blockInnermostSdtId,
  blockInSdt,
  commonPrefixLen,
  fullSdtChain,
  innermostSdtId,
  inSdt,
  normSdtPath,
  pushSdt,
  removeSdt,
  sdtPathEq,
} from "./sdt";

const CHAR: CharStyle = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const styleWith = (sdtPath?: string[]): CharStyle => ({ ...CHAR, sdtPath });
const blockWith = (sdtPath?: string[]): Block =>
  ({ kind: "paragraph", id: "p", revision: 0, runs: [], style: {} as Paragraph["style"], sdtPath }) as Block;

describe("innermostSdtId", () => {
  it("returns the last id of a non-empty inline path", () => {
    expect(innermostSdtId(styleWith(["a", "b", "c"]))).toBe("c");
  });

  it("returns undefined for an empty or absent path", () => {
    expect(innermostSdtId(styleWith([]))).toBeUndefined();
    expect(innermostSdtId(styleWith(undefined))).toBeUndefined();
  });
});

describe("blockInnermostSdtId", () => {
  it("returns the last id of a block path, undefined when absent", () => {
    expect(blockInnermostSdtId(blockWith(["x", "y"]))).toBe("y");
    expect(blockInnermostSdtId(blockWith(undefined))).toBeUndefined();
  });
});

describe("fullSdtChain", () => {
  it("concatenates block prefix then inline suffix (outer→inner)", () => {
    expect(fullSdtChain(blockWith(["b1", "b2"]), styleWith(["i1"]))).toEqual(["b1", "b2", "i1"]);
  });

  it("returns the inline path alone when there is no block path", () => {
    expect(fullSdtChain(blockWith(undefined), styleWith(["i1", "i2"]))).toEqual(["i1", "i2"]);
    expect(fullSdtChain(undefined, styleWith(["i1"]))).toEqual(["i1"]);
  });

  it("returns a fresh array (does not alias the inline path)", () => {
    const style = styleWith(["i1"]);
    const out = fullSdtChain(blockWith(undefined), style);
    expect(out).not.toBe(style.sdtPath);
  });

  it("returns empty array when neither has a path", () => {
    expect(fullSdtChain(blockWith(undefined), styleWith(undefined))).toEqual([]);
  });
});

describe("sdtPathEq", () => {
  it("treats undefined and [] as equal (empty path)", () => {
    expect(sdtPathEq(undefined, [])).toBe(true);
    expect(sdtPathEq([], undefined)).toBe(true);
    expect(sdtPathEq(undefined, undefined)).toBe(true);
  });

  it("compares element-wise", () => {
    expect(sdtPathEq(["a", "b"], ["a", "b"])).toBe(true);
    expect(sdtPathEq(["a", "b"], ["a", "c"])).toBe(false);
    expect(sdtPathEq(["a"], ["a", "b"])).toBe(false);
  });
});

describe("commonPrefixLen", () => {
  it("counts the shared outer prefix", () => {
    expect(commonPrefixLen(["a", "b", "c"], ["a", "b", "z"])).toBe(2);
    expect(commonPrefixLen(["a"], ["b"])).toBe(0);
    expect(commonPrefixLen(["a", "b"], ["a", "b"])).toBe(2);
  });

  it("is bounded by the shorter path length", () => {
    expect(commonPrefixLen(["a", "b"], ["a", "b", "c"])).toBe(2);
    expect(commonPrefixLen([], ["a"])).toBe(0);
  });
});

describe("inSdt / blockInSdt", () => {
  it("detects membership at any nesting level", () => {
    expect(inSdt(styleWith(["a", "b", "c"]), "b")).toBe(true);
    expect(inSdt(styleWith(["a", "b", "c"]), "z")).toBe(false);
    expect(inSdt(styleWith(undefined), "a")).toBe(false);
  });

  it("blockInSdt mirrors inSdt for block paths", () => {
    expect(blockInSdt(blockWith(["x", "y"]), "x")).toBe(true);
    expect(blockInSdt(blockWith(["x", "y"]), "q")).toBe(false);
    expect(blockInSdt(blockWith(undefined), "x")).toBe(false);
  });
});

describe("normSdtPath", () => {
  it("returns undefined for empty/absent and a fresh copy otherwise", () => {
    expect(normSdtPath(undefined)).toBeUndefined();
    expect(normSdtPath([])).toBeUndefined();
    const src = ["a", "b"];
    const out = normSdtPath(src)!;
    expect(out).toEqual(["a", "b"]);
    expect(out).not.toBe(src);
  });
});

describe("pushSdt", () => {
  it("appends to an existing path without mutating it", () => {
    const src = ["a"];
    const out = pushSdt(src, "b");
    expect(out).toEqual(["a", "b"]);
    expect(src).toEqual(["a"]);
  });

  it("starts a new single-element path from empty/undefined", () => {
    expect(pushSdt(undefined, "a")).toEqual(["a"]);
    expect(pushSdt([], "a")).toEqual(["a"]);
  });
});

describe("removeSdt", () => {
  it("removes a single id, preserving the rest", () => {
    expect(removeSdt(["a", "b", "c"], "b")).toEqual(["a", "c"]);
  });

  it("returns undefined when removal empties the path", () => {
    expect(removeSdt(["a"], "a")).toBeUndefined();
    expect(removeSdt(undefined, "a")).toBeUndefined();
  });

  it("removes ALL matching occurrences (filter semantics)", () => {
    expect(removeSdt(["a", "b", "a"], "a")).toEqual(["b"]);
  });

  it("returns a new array and leaves the input intact", () => {
    const src = ["a", "b"];
    const out = removeSdt(src, "a");
    expect(out).toEqual(["b"]);
    expect(src).toEqual(["a", "b"]);
  });
});

describe("ancestryThrough", () => {
  it("returns the prefix up to and including the id", () => {
    expect(ancestryThrough(["a", "b", "c"], "b")).toEqual(["a", "b"]);
    expect(ancestryThrough(["a", "b", "c"], "a")).toEqual(["a"]);
    expect(ancestryThrough(["a", "b", "c"], "c")).toEqual(["a", "b", "c"]);
  });

  it("returns undefined when the id is absent or path is empty", () => {
    expect(ancestryThrough(["a", "b"], "z")).toBeUndefined();
    expect(ancestryThrough(undefined, "a")).toBeUndefined();
  });
});
