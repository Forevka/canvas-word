import { describe, it, expect } from "vitest";
import { anchorBeforeId, resolveInsertIndex } from "./ribbon";

describe("resolveInsertIndex", () => {
  const order = ["a", "b", "c"];

  it("appends when no anchor is given", () => {
    expect(resolveInsertIndex(order)).toBe(3);
    expect(resolveInsertIndex(order, {})).toBe(3);
  });

  it("inserts before an existing anchor", () => {
    expect(resolveInsertIndex(order, { before: "b" })).toBe(1);
    expect(resolveInsertIndex(order, { before: "a" })).toBe(0);
  });

  it("inserts after an existing anchor", () => {
    expect(resolveInsertIndex(order, { after: "b" })).toBe(2);
    expect(resolveInsertIndex(order, { after: "c" })).toBe(3);
  });

  it("prefers before over after when both given", () => {
    expect(resolveInsertIndex(order, { before: "a", after: "c" })).toBe(0);
  });

  it("appends when the anchor id is unknown", () => {
    expect(resolveInsertIndex(order, { before: "nope" })).toBe(3);
    expect(resolveInsertIndex(order, { after: "nope" })).toBe(3);
  });
});

describe("anchorBeforeId", () => {
  const order = ["a", "b", "c"];

  it("returns the id to insert before, or null to append", () => {
    expect(anchorBeforeId(order, { before: "b" })).toBe("b");
    expect(anchorBeforeId(order, { after: "c" })).toBeNull();
    expect(anchorBeforeId(order, undefined)).toBeNull();
  });

  it("excludes the moving id so a move targets the post-removal list", () => {
    // Moving "a" to after "b": with "a" removed the list is [b,c]; after b → index 1 → "c".
    expect(anchorBeforeId(order, { after: "b" }, "a")).toBe("c");
    // Moving "a" to the end: list without a is [b,c]; after c → append → null.
    expect(anchorBeforeId(order, { after: "c" }, "a")).toBeNull();
    // Moving "c" to before "b": list without c is [a,b]; before b → "b".
    expect(anchorBeforeId(order, { before: "b" }, "c")).toBe("b");
  });
});
