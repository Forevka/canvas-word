// OpenAPI generation + Swagger UI. @fastify/swagger introspects each route's
// `schema` (params/query/body/response + tags/summary/security/consumes/produces)
// and assembles the document; @fastify/swagger-ui serves it (self-hosted assets,
// no CDN). Both must be registered BEFORE the routes so the onRoute hook captures
// their schemas. The spec is no longer hand-authored — the routes ARE the spec.
import swagger from "@fastify/swagger";
import swaggerUi from "@fastify/swagger-ui";
import type { FastifyInstance } from "fastify";
import { PORT } from "../config";
import { registerSchemas } from "../schemas/models";

export async function registerOpenapi(app: FastifyInstance): Promise<void> {
  registerSchemas(app);

  // Advertised base URL: PUBLIC_BASE_URL in deployed envs, else local dev on the
  // configured port (so it tracks BACKEND_PORT instead of a hardcoded 8787).
  const serverUrl = process.env.PUBLIC_BASE_URL ?? `http://localhost:${PORT}`;

  await app.register(swagger, {
    // Keep shared-schema component names as their $id (Error, Change, …) instead
    // of the default def-0/def-1/… so the spec and Swagger UI read clearly.
    refResolver: {
      buildLocalReference: (json, _baseUri, _fragment, i) => (typeof json.$id === "string" ? json.$id : `def-${i}`),
    },
    openapi: {
      openapi: "3.1.0",
      info: {
        title: "canvas-word collaboration API",
        version: "0.0.1",
        description:
          "Event-sourced document store: a base snapshot plus an append-only, totally-ordered " +
          "change log. Reconstruct any version = snapshot + replay. Live editing also flows over a " +
          "WebSocket channel (not described by OpenAPI): connect to `ws://<host>/ws?doc=<id>`, send " +
          "`{type:'submit', change}` and receive `{type:'change', change}` broadcasts.",
      },
      servers: [{ url: serverUrl, description: "local dev" }],
      tags: [
        { name: "docs", description: "Documents, change log, review overlay, export" },
        { name: "media", description: "Content-addressed image blobs" },
        { name: "render", description: "Stateless headless .docx → PDF" },
        { name: "admin", description: "Dashboard API (token-guarded)" },
      ],
      components: {
        securitySchemes: {
          ApiKeyAuth: {
            type: "apiKey",
            in: "header",
            name: "X-API-Key",
            description: "Integration API key. Mint one via POST /admin/tokens (admin-only). Click Authorize and paste it.",
          },
          BearerAuth: {
            type: "http",
            scheme: "bearer",
            description: "An admin token (from POST /admin/login) OR an integration API key, sent as `Authorization: Bearer <token>`.",
          },
        },
      },
    },
  });

  await app.register(swaggerUi, {
    routePrefix: "/swagger",
    uiConfig: { docExpansion: "list", deepLinking: true },
  });
}
