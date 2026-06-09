import { describe, expect, it } from "vitest";
import { colorForId, deterministicUserId, userDisplayName } from "./change";

describe("identity helpers", () => {
  it("deterministicUserId is stable and name-normalized", () => {
    const a = deterministicUserId("Alice Smith");
    expect(a).toBe(deterministicUserId("alice smith")); // case-insensitive
    expect(a).toBe(deterministicUserId("  Alice   Smith  ")); // whitespace-normalized
    expect(a).not.toBe(deterministicUserId("Bob Jones"));
    expect(a).toMatch(/^u[0-9a-f]{8}$/);
  });

  it("colorForId is stable per id and a valid HSL color", () => {
    expect(colorForId("u123")).toBe(colorForId("u123"));
    expect(colorForId("u123")).toMatch(/^hsl\(\d+, 70%, 45%\)$/);
    expect(colorForId("u123")).not.toBe(colorForId("u456"));
  });

  it("userDisplayName handles missing parts", () => {
    expect(userDisplayName({ firstName: "Alice", lastName: "Smith" })).toBe("Alice Smith");
    expect(userDisplayName({ firstName: "Cher", lastName: "" })).toBe("Cher");
    expect(userDisplayName({ firstName: "", lastName: "" })).toBe("Anonymous");
  });
});
