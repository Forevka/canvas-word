// Dashboard auth — a single pre-seeded admin (username + password from env) and
// a stateless signed bearer token. No session table: the token is its own proof
// (HMAC-signed payload), so any server instance with the same secret accepts it.
//
//   POST /admin/login { username, password }  ->  { token, expiresAt }
//   Authorization: Bearer <token>             ->  requireAdmin gate
//
// Env (with dev-friendly defaults; SET THESE IN PRODUCTION):
//   ADMIN_USERNAME      (default "admin")
//   ADMIN_PASSWORD      (default "admin")
//   ADMIN_TOKEN_SECRET  (default a fixed dev string)

import { createHmac, timingSafeEqual } from "node:crypto";
import type { IncomingHttpHeaders } from "node:http";

/** Anything carrying request headers — a raw IncomingMessage or a Fastify request. */
type HasHeaders = { headers: IncomingHttpHeaders };

const ADMIN_USERNAME = process.env.ADMIN_USERNAME ?? "admin";
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD ?? "admin";
const TOKEN_SECRET = process.env.ADMIN_TOKEN_SECRET ?? "dev-insecure-token-secret-change-me";
const TOKEN_TTL_MS = Number(process.env.ADMIN_TOKEN_TTL_MS ?? 12 * 60 * 60 * 1000); // 12h

function b64url(buf: Buffer): string {
  return buf.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function sign(payloadB64: string): string {
  return b64url(createHmac("sha256", TOKEN_SECRET).update(payloadB64).digest());
}

/** Constant-time string compare that never throws on length mismatch. */
function safeEqual(a: string, b: string): boolean {
  const ab = Buffer.from(a);
  const bb = Buffer.from(b);
  if (ab.length !== bb.length) return false;
  return timingSafeEqual(ab, bb);
}

export function checkCredentials(username: string, password: string): boolean {
  // Compare both so timing doesn't reveal which field was wrong.
  const okUser = safeEqual(username, ADMIN_USERNAME);
  const okPass = safeEqual(password, ADMIN_PASSWORD);
  return okUser && okPass;
}

export interface AdminToken {
  username: string;
  exp: number;
}

export function issueToken(username: string): { token: string; expiresAt: number } {
  const exp = Date.now() + TOKEN_TTL_MS;
  const payload: AdminToken = { username, exp };
  const payloadB64 = b64url(Buffer.from(JSON.stringify(payload), "utf8"));
  const token = `${payloadB64}.${sign(payloadB64)}`;
  return { token, expiresAt: exp };
}

export function verifyToken(token: string): AdminToken | null {
  const dot = token.indexOf(".");
  if (dot <= 0) return null;
  const payloadB64 = token.slice(0, dot);
  const sig = token.slice(dot + 1);
  if (!safeEqual(sig, sign(payloadB64))) return null;
  try {
    const payload = JSON.parse(
      Buffer.from(payloadB64.replace(/-/g, "+").replace(/_/g, "/"), "base64").toString("utf8"),
    ) as AdminToken;
    if (typeof payload.exp !== "number" || payload.exp < Date.now()) return null;
    return payload;
  } catch {
    return null;
  }
}

/** Extract a bearer token from the Authorization header and verify it. */
export function authFromRequest(req: HasHeaders): AdminToken | null {
  const header = req.headers["authorization"];
  if (!header || Array.isArray(header)) return null;
  const m = /^Bearer\s+(.+)$/i.exec(header.trim());
  if (!m) return null;
  return verifyToken(m[1]!);
}
