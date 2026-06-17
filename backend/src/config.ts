// Centralised environment configuration — read once at module load so the rest
// of the codebase never touches process.env directly.

/** HTTP port for the collaboration server. */
export const PORT = Number(process.env.BACKEND_PORT ?? 8787);

/** Uploads are token-guarded by default; set UPLOAD_PUBLIC=1 for anonymous upload. */
export const UPLOAD_REQUIRES_AUTH = process.env.UPLOAD_PUBLIC !== "1";

// Bodies have no fixed ceiling in this collaboration server: a base snapshot or a
// media blob can be large. Keep a generous limit (overridable) instead of Fastify's
// 1MB default — the raw http server it replaced had no limit at all.
export const BODY_LIMIT = Number(process.env.BACKEND_BODY_LIMIT ?? 256 * 1024 * 1024);
