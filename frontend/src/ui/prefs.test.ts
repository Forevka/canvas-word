import { afterEach, describe, expect, it, vi } from "vitest";
import { readPref, writePref } from "./prefs";

// Minimal in-memory localStorage (the node test env has none) — mirrors recentList.test.ts.
function stubStorage(): Map<string, string> {
  const store = new Map<string, string>();
  vi.stubGlobal("localStorage", {
    getItem: (k: string) => (store.has(k) ? store.get(k)! : null),
    setItem: (k: string, v: string) => void store.set(k, v),
    removeItem: (k: string) => void store.delete(k),
  });
  return store;
}

afterEach(() => vi.unstubAllGlobals());

describe("readPref / writePref", () => {
  const THEMES = ["light", "dark", "system"] as const;

  it("round-trips an allowed value", () => {
    stubStorage();
    expect(readPref("cw:pref:theme", THEMES)).toBeNull(); // absent → null
    writePref("cw:pref:theme", "dark");
    expect(readPref("cw:pref:theme", THEMES)).toBe("dark");
  });

  it("rejects a value outside the allowed set", () => {
    const store = stubStorage();
    store.set("cw:pref:theme", "chartreuse"); // not one of THEMES
    expect(readPref("cw:pref:theme", THEMES)).toBeNull();
  });

  it("returns null (never throws) when storage is unavailable", () => {
    vi.stubGlobal("localStorage", {
      getItem: () => { throw new Error("denied"); },
      setItem: () => { throw new Error("denied"); },
    });
    expect(readPref("cw:pref:theme", THEMES)).toBeNull();
    expect(() => writePref("cw:pref:theme", "dark")).not.toThrow();
  });
});
