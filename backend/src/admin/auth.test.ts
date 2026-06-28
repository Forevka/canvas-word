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
    // Flip the second-to-last character to one guaranteed to differ from it,
    // so the tamper is never a no-op (the old logic keyed the replacement off
    // the last char while overwriting the second-to-last — see issue #40).
    const i = token.length - 2;
    const replacement = token[i] === "a" ? "b" : "a";
    const tampered = token.slice(0, i) + replacement + token.slice(i + 1);
    expect(verifyToken(tampered)).toBeNull();
    expect(verifyToken("garbage")).toBeNull();
    expect(verifyToken("")).toBeNull();
  });
});
