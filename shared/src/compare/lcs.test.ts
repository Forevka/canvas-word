import { describe, expect, it } from "vitest";
import { align } from "./lcs";

const id = (s: string): string => s;
const tags = <T>(hunks: { tag: string; a: T[]; b: T[] }[]): string[] => hunks.map((h) => h.tag);

describe("align — LCS hunks", () => {
  it("identical sequences are one equal hunk", () => {
    const h = align(["a", "b", "c"], ["a", "b", "c"], id);
    expect(tags(h)).toEqual(["equal"]);
    expect(h[0]!.a).toEqual(["a", "b", "c"]);
  });

  it("pure insertion", () => {
    const h = align(["a", "c"], ["a", "b", "c"], id);
    expect(tags(h)).toEqual(["equal", "insert", "equal"]);
    expect(h[1]!.b).toEqual(["b"]);
    expect(h[1]!.a).toEqual([]);
  });

  it("pure deletion", () => {
    const h = align(["a", "b", "c"], ["a", "c"], id);
    expect(tags(h)).toEqual(["equal", "delete", "equal"]);
    expect(h[1]!.a).toEqual(["b"]);
  });

  it("adjacent delete+insert coalesce into replace", () => {
    const h = align(["a", "x", "c"], ["a", "y", "c"], id);
    expect(tags(h)).toEqual(["equal", "replace", "equal"]);
    expect(h[1]!.a).toEqual(["x"]);
    expect(h[1]!.b).toEqual(["y"]);
    expect(h[1]!.ai).toBe(1);
    expect(h[1]!.bi).toBe(1);
  });

  it("everything different is one replace", () => {
    expect(tags(align(["a", "b"], ["c", "d"], id))).toEqual(["replace"]);
  });

  it("empty inputs", () => {
    expect(align<string>([], [], id)).toEqual([]);
    expect(tags(align([], ["a"], id))).toEqual(["insert"]);
    expect(tags(align(["a"], [], id))).toEqual(["delete"]);
  });
});
