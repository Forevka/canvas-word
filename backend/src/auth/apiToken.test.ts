import { createHash } from "node:crypto";
import { describe, expect, it } from "vitest";
import { hashToken, newApiToken } from "./apiToken";

describe("hashToken", () => {
  it("produces a 64-char lowercase hex sha256 digest", () => {
    const h = hashToken("cwk_example");
    expect(h).toMatch(/^[0-9a-f]{64}$/);
  });

  it("is deterministic for the same input", () => {
    expect(hashToken("same")).toBe(hashToken("same"));
  });

  it("differs for different inputs", () => {
    expect(hashToken("a")).not.toBe(hashToken("b"));
  });

  it("matches a reference sha256 computed independently", () => {
    const ref = createHash("sha256").update("token-123").digest("hex");
    expect(hashToken("token-123")).toBe(ref);
  });
});

describe("newApiToken", () => {
  it("returns a token with the cwk_ prefix", () => {
    expect(newApiToken().token.startsWith("cwk_")).toBe(true);
  });

  it("uses the first 12 chars of the token as the display prefix", () => {
    const t = newApiToken();
    expect(t.prefix).toBe(t.token.slice(0, 12));
    expect(t.prefix).toHaveLength(12);
  });

  it("sets tokenHash to hashToken(token)", () => {
    const t = newApiToken();
    expect(t.tokenHash).toBe(hashToken(t.token));
  });

  it("generates a fresh random token on every call", () => {
    const a = newApiToken();
    const b = newApiToken();
    expect(a.token).not.toBe(b.token);
    expect(a.tokenHash).not.toBe(b.tokenHash);
  });

  it("emits 24 random bytes as 48 hex chars after the prefix", () => {
    // cwk_ (4) + 48 hex = 52 chars total.
    expect(newApiToken().token).toHaveLength(52);
  });
});
