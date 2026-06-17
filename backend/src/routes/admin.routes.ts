// Admin API (dashboard) — token-guarded except the login route.
import { randomBytes } from "node:crypto";
import type { FastifyInstance, FastifySchema } from "fastify";
import { checkCredentials, issueToken } from "../admin/auth";
import { newApiToken } from "../auth/apiToken";
import type { AppContext } from "../context";
import { requireAdmin } from "../http/auth";
import { SESSION_ENDED_EVENT } from "../webhooks/dispatcher";

const randomSecret = (): string => randomBytes(24).toString("hex");

const idParam = { type: "object", properties: { id: { type: "string" } }, required: ["id"] } as const;
const okId = { type: "object", properties: { id: { type: "string" } }, additionalProperties: true } as const;
const listOf = (key: string) =>
  ({ type: "object", properties: { [key]: { type: "array", items: { type: "object", additionalProperties: true } } }, additionalProperties: true }) as const;

/** Shared bits for every guarded admin route: the tag, BearerAuth, and a 401
 *  merged into whatever response codes the route declares. */
const guarded = (extra: FastifySchema): { preHandler: typeof requireAdmin; schema: FastifySchema } => {
  const { response, ...rest } = extra;
  return {
    preHandler: requireAdmin,
    schema: {
      tags: ["admin"],
      security: [{ BearerAuth: [] }],
      ...rest,
      response: { 401: { $ref: "Error#" }, ...(response as Record<string, unknown>) },
    },
  };
};

export function registerAdminRoutes(app: FastifyInstance, { store }: AppContext): void {
  app.post(
    "/admin/login",
    {
      schema: {
        tags: ["admin"],
        summary: "Exchange username + password for a bearer token",
        body: { type: "object", properties: { username: { type: "string" }, password: { type: "string" } } },
        response: {
          200: { type: "object", properties: { token: { type: "string" }, expiresAt: { type: "integer" } }, additionalProperties: true },
          401: { $ref: "Error#" },
        },
      },
    },
    async (req, reply) => {
      const { username, password } = (req.body ?? {}) as { username?: string; password?: string };
      if (!username || !password || !checkCredentials(username, password)) {
        return reply.code(401).send({ error: "invalid credentials" });
      }
      return issueToken(username);
    },
  );

  app.get("/admin/docs", guarded({ summary: "List all documents", response: { 200: listOf("documents") } }), async () => ({
    documents: await store.listDocuments(),
  }));

  // GET /admin/docs/:id/activity — per-document edit history with author names.
  app.get(
    "/admin/docs/:id/activity",
    guarded({ summary: "Per-document edit history", params: idParam, response: { 200: { type: "object", additionalProperties: true }, 404: { $ref: "Error#" } } }),
    async (req, reply) => {
      const { id } = req.params as { id: string };
      const activity = await store.getActivity(id);
      if (!activity) return reply.code(404).send({ error: "document not found" });
      return activity;
    },
  );

  app.get("/admin/users", guarded({ summary: "List known users", response: { 200: listOf("users") } }), async () => ({
    users: await store.listUsers(),
  }));

  app.get("/admin/webhooks", guarded({ summary: "List webhooks", response: { 200: listOf("webhooks") } }), async () => ({
    webhooks: await store.listWebhooks(),
  }));
  app.post(
    "/admin/webhooks",
    guarded({
      summary: "Register a webhook",
      body: {
        type: "object",
        properties: { url: { type: "string" }, secret: { type: "string" }, events: { type: "array", items: { type: "string" } } },
      },
      response: { 201: { type: "object", additionalProperties: true }, 400: { $ref: "Error#" } },
    }),
    async (req, reply) => {
      const b = (req.body ?? {}) as { url?: string; secret?: string; events?: string[] };
      if (!b.url) return reply.code(400).send({ error: "url required" });
      const secret = b.secret && b.secret.length > 0 ? b.secret : randomSecret();
      const events = b.events && b.events.length > 0 ? b.events : [SESSION_ENDED_EVENT];
      return reply.code(201).send(await store.createWebhook({ url: b.url, secret, events }));
    },
  );
  app.delete("/admin/webhooks/:id", guarded({ summary: "Delete a webhook", params: idParam, response: { 200: okId } }), async (req) => {
    const { id } = req.params as { id: string };
    await store.deleteWebhook(id);
    return { id };
  });
  app.patch(
    "/admin/webhooks/:id",
    guarded({ summary: "Enable/disable a webhook", params: idParam, body: { type: "object", properties: { active: { type: "boolean" } } }, response: { 200: okId } }),
    async (req) => {
      const { id } = req.params as { id: string };
      const b = (req.body ?? {}) as { active?: boolean };
      if (typeof b.active === "boolean") await store.setWebhookActive(id, b.active);
      return { id };
    },
  );
  app.get(
    "/admin/webhooks/:id/deliveries",
    guarded({ summary: "Recent delivery attempts for a webhook", params: idParam, response: { 200: listOf("deliveries") } }),
    async (req) => {
      const { id } = req.params as { id: string };
      return { deliveries: await store.listDeliveries(id) };
    },
  );

  // Integration tokens (API keys) for 3rd-party /upload.
  app.get("/admin/tokens", guarded({ summary: "List integration API tokens", response: { 200: listOf("tokens") } }), async () => ({
    tokens: await store.listApiTokens(),
  }));
  app.post(
    "/admin/tokens",
    guarded({
      summary: "Mint an integration API token (plaintext returned once)",
      body: { type: "object", properties: { name: { type: "string" } } },
      response: { 201: { type: "object", additionalProperties: true } },
    }),
    async (req, reply) => {
      const b = (req.body ?? {}) as { name?: string };
      const name = b.name?.trim() || "integration";
      const { token, tokenHash, prefix } = newApiToken();
      const rec = await store.createApiToken({ name, tokenHash, prefix });
      // The plaintext token is returned ONCE here and never stored.
      return reply.code(201).send({ ...rec, token });
    },
  );
  app.delete("/admin/tokens/:id", guarded({ summary: "Revoke an integration API token", params: idParam, response: { 200: okId } }), async (req) => {
    const { id } = req.params as { id: string };
    await store.revokeApiToken(id);
    return { id };
  });
}
