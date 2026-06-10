// Integration API tokens. A token is a random `cwk_…` string handed to a
// 3rd-party system; the server stores only its SHA-256 hash, so a DB leak never
// exposes a usable credential. The plaintext is shown once at creation.

import { createHash, randomBytes } from "node:crypto";

export const hashToken = (token: string): string =>
  createHash("sha256").update(token).digest("hex");

export interface NewApiToken {
  /** Plaintext token — returned to the caller ONCE, never stored. */
  token: string;
  /** sha256(token) hex — what we persist and look up by. */
  tokenHash: string;
  /** Leading chars, stored for display so a token is identifiable in the UI. */
  prefix: string;
}

export function newApiToken(): NewApiToken {
  const token = `cwk_${randomBytes(24).toString("hex")}`;
  return { token, tokenHash: hashToken(token), prefix: token.slice(0, 12) };
}
