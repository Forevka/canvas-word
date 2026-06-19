import type { FastifyCorsOptions } from "@fastify/cors";

// Origin policy: CORS_ORIGIN env wins (comma-separated → array). If unset, dev/test
// stays permissive ("*"); production defaults closed (false) so an unconfigured
// deploy doesn't expose itself cross-origin — set CORS_ORIGIN to open it.
function resolveOrigin(): NonNullable<FastifyCorsOptions["origin"]> {
  const env = process.env.CORS_ORIGIN?.trim();
  if (env) {
    const list = env.split(",").map((o) => o.trim()).filter(Boolean);
    return list.length > 1 ? list : (list[0] ?? "*");
  }
  return process.env.NODE_ENV === "production" ? false : "*";
}

// Permissive CORS for every route + preflight (replaces the hand-rolled OPTIONS
// branch). x-api-key is allowed so browser integrations can upload with a key.
export const corsOptions: FastifyCorsOptions = {
  origin: resolveOrigin(),
  methods: ["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
  allowedHeaders: ["content-type", "content-encoding", "authorization", "x-filename", "x-user", "x-api-key"],
  exposedHeaders: ["content-disposition"],
};
