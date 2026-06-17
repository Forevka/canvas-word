// HTTP-layer authorization: the admin gate and the upload gate, expressed as
// Fastify preHandler hooks over the lower-level token primitives in admin/auth
// and auth/apiToken.
import type { IncomingHttpHeaders } from "node:http";
import type { FastifyReply, FastifyRequest } from "fastify";
import { authFromRequest } from "../admin/auth";
import { hashToken } from "../auth/apiToken";
import { UPLOAD_REQUIRES_AUTH } from "../config";
import { headerStr } from "./headers";
import type { ChangeStore } from "../store/ChangeStore";

/** Admin gate: a valid dashboard bearer token, else 401. */
export async function requireAdmin(req: FastifyRequest, reply: FastifyReply): Promise<void> {
  if (!authFromRequest(req)) await reply.code(401).send({ error: "unauthorized" });
}

/** Anything carrying request headers — a raw IncomingMessage or a Fastify request. */
type HasHeaders = { headers: IncomingHttpHeaders };

/** Bearer value from the Authorization header (admin token OR an integration token). */
function bearerToken(req: HasHeaders): string | undefined {
  const h = req.headers["authorization"];
  if (!h || Array.isArray(h)) return undefined;
  const m = /^Bearer\s+(.+)$/i.exec(h.trim());
  return m ? m[1] : undefined;
}

/** Upload auth: a dashboard admin token (UI) OR a 3rd-party integration token
 *  (sent as X-API-Key or Bearer). */
export async function authorizeUpload(req: FastifyRequest, store: ChangeStore): Promise<boolean> {
  if (authFromRequest(req)) return true; // admin (dashboard) token
  const raw = headerStr(req.headers["x-api-key"]) ?? bearerToken(req);
  if (!raw) return false;
  return store.verifyApiToken(hashToken(raw));
}

/** Build the upload preHandler bound to a store (no-op when uploads are public). */
export function uploadAuthHook(store: ChangeStore) {
  return async (req: FastifyRequest, reply: FastifyReply): Promise<void> => {
    if (UPLOAD_REQUIRES_AUTH && !(await authorizeUpload(req, store))) {
      await reply.code(401).send({ error: "unauthorized" });
    }
  };
}
