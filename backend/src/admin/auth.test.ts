import { describe, expect, it } from "vitest";
import { checkCredentials, issueToken, verifyToken } from "./auth";

// Uses the dev defaults (ADMIN_USERNAME=admin / ADMIN_PASSWORD=admin) unless the
// environment overrides them.

describe("admin auth", () => {
  it("accepts the seeded credentials and rejects others", () => {
    expect(checkCredentials("admin", "admin")).toBe(true);
    expect(checkCredentials("admin", "wrong")).toBe(false);
    expect(checkCredentials("nope", "admin")).toBe(false);
    expect(checkCredentials("", "")).toBe(false);
  });

  it("issues a token that verifies back to the username", () => {
    const { token, expiresAt } = issueToken("admin");
    expect(expiresAt).toBeGreaterThan(Date.now());
    const payload = verifyToken(token);
    expect(payload?.username).toBe("admin");
  });

  it("rejects a tampered token", () => {
    const { token } = issueToken("admin");
    const tampered = token.slice(0, -2) + (token.endsWith("a") ? "b" : "a") + token.slice(-1);
    expect(verifyToken(tampered)).toBeNull();
    expect(verifyToken("garbage")).toBeNull();
    expect(verifyToken("")).toBeNull();
  });
});
