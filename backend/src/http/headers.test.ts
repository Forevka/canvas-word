import { describe, expect, it } from "vitest";
import { headerStr, safeParseUser, sanitizeFilename } from "./headers";

describe("headerStr", () => {
  it("returns the value unchanged for a single string", () => {
    expect(headerStr("text/plain")).toBe("text/plain");
  });

  it("returns the first element for a multi-valued header", () => {
    expect(headerStr(["a", "b", "c"])).toBe("a");
  });

  it("returns undefined for an undefined header", () => {
    expect(headerStr(undefined)).toBeUndefined();
  });

  it("returns undefined for an empty array", () => {
    expect(headerStr([])).toBeUndefined();
  });
});

describe("safeParseUser", () => {
  it("parses a full valid user object", () => {
    const raw = JSON.stringify({ id: "u1", firstName: "Ed", lastName: "Itor" });
    expect(safeParseUser(raw)).toEqual({ id: "u1", firstName: "Ed", lastName: "Itor" });
  });

  it("defaults missing first/last names to empty strings", () => {
    expect(safeParseUser(JSON.stringify({ id: "u2" }))).toEqual({ id: "u2", firstName: "", lastName: "" });
  });

  it("returns undefined when id is missing", () => {
    expect(safeParseUser(JSON.stringify({ firstName: "No", lastName: "Id" }))).toBeUndefined();
  });

  it("returns undefined when id is not a string", () => {
    expect(safeParseUser(JSON.stringify({ id: 42 }))).toBeUndefined();
  });

  it("returns undefined for malformed JSON", () => {
    expect(safeParseUser("{ not json")).toBeUndefined();
  });

  it("returns undefined for a JSON primitive (non-object)", () => {
    expect(safeParseUser("null")).toBeUndefined();
    expect(safeParseUser("\"just a string\"")).toBeUndefined();
  });

  it("ignores extra unexpected fields, keeping only the known shape", () => {
    const raw = JSON.stringify({ id: "u3", firstName: "A", lastName: "B", admin: true });
    expect(safeParseUser(raw)).toEqual({ id: "u3", firstName: "A", lastName: "B" });
  });
});

describe("sanitizeFilename", () => {
  it("replaces quotes, slashes and path separators with underscores", () => {
    expect(sanitizeFilename('a"b/c\\d:e')).toBe("a_b_c_d_e");
  });

  it("replaces wildcard and angle-bracket control chars", () => {
    expect(sanitizeFilename("a*b?c<d>e|f")).toBe("a_b_c_d_e_f");
  });

  it("collapses a run of forbidden characters into a single underscore", () => {
    expect(sanitizeFilename('a"""b')).toBe("a_b");
  });

  it("trims surrounding whitespace", () => {
    expect(sanitizeFilename("  report  ")).toBe("report");
  });

  it("caps the result at 120 characters", () => {
    const out = sanitizeFilename("x".repeat(200));
    expect(out).toHaveLength(120);
  });

  it("falls back to 'document' for an empty string", () => {
    expect(sanitizeFilename("")).toBe("document");
  });

  it("falls back to 'document' for a whitespace-only string", () => {
    expect(sanitizeFilename("   ")).toBe("document");
  });

  it("falls back to 'document' for a string of only forbidden chars", () => {
    // The replace collapses to "_", which is non-empty, so this is NOT "document".
    expect(sanitizeFilename("///")).toBe("_");
  });

  it("keeps a normal title intact", () => {
    expect(sanitizeFilename("My Report 2026")).toBe("My Report 2026");
  });

  it("replaces hyphens too (they are inside the forbidden char class)", () => {
    expect(sanitizeFilename("2026-06-24")).toBe("2026_06_24");
  });
});
