// Issue #244 — the shared MRU list backing the shape gallery's "Recently used" row
// (A3) and the colour popover's "Recent colours" row (D1).
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { readRecent, pushRecent } from "./recentList";

// Minimal in-memory localStorage (the node test env has none).
function installStorage(): void {
  const store = new Map<string, string>();
  vi.stubGlobal("localStorage", {
    getItem: (k: string) => (store.has(k) ? store.get(k)! : null),
    setItem: (k: string, v: string) => void store.set(k, v),
    removeItem: (k: string) => void store.delete(k),
    clear: () => store.clear(),
  });
}

describe("recentList", () => {
  beforeEach(installStorage);
  afterEach(() => vi.unstubAllGlobals());

  it("returns [] for an empty/unknown key", () => {
    expect(readRecent("k", 8)).toEqual([]);
  });

  it("pushes newest-first and reads back", () => {
    pushRecent("k", "a", 8);
    pushRecent("k", "b", 8);
    expect(readRecent("k", 8)).toEqual(["b", "a"]);
  });

  it("de-duplicates, moving a repeat to the front", () => {
    pushRecent("k", "a", 8);
    pushRecent("k", "b", 8);
    pushRecent("k", "a", 8);
    expect(readRecent("k", 8)).toEqual(["a", "b"]);
  });

  it("clamps to the max length (oldest dropped)", () => {
    for (const c of ["a", "b", "c", "d"]) pushRecent("k", c, 3);
    expect(readRecent("k", 3)).toEqual(["d", "c", "b"]);
  });

  it("survives a malformed stored value and a thrown storage", () => {
    localStorage.setItem("k", "not json");
    expect(readRecent("k", 8)).toEqual([]);
    vi.stubGlobal("localStorage", {
      getItem: () => { throw new Error("blocked"); },
      setItem: () => { throw new Error("blocked"); },
    });
    expect(readRecent("k", 8)).toEqual([]);
    expect(() => pushRecent("k", "a", 8)).not.toThrow();
  });
});
