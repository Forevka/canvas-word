import type { FastifyInstance } from "fastify";

// API docs. The interactive UI is served by @fastify/swagger-ui at /swagger; this
// keeps the canonical machine-readable spec at the stable /openapi.json path,
// generated from the route schemas (hidden from the spec itself).
export function registerMetaRoutes(app: FastifyInstance): void {
  app.get("/openapi.json", { schema: { hide: true } }, async () => app.swagger());
}
