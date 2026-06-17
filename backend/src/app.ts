// Composition root: assemble the Fastify app from plugins + route modules and
// wire the runtime dependencies (store, broadcaster, webhook dispatcher) that the
// route layer consumes through AppContext.
//
// The store assigns each change its canonical seq; the server is the single source
// of order (the ShareDB-style seam where OT rebasing lands in Phase 4).
// Reconstructing a version is base snapshot + replay (shared).
import type { Server } from "node:http";
import cors from "@fastify/cors";
import websocket from "@fastify/websocket";
import Fastify, { type FastifyInstance } from "fastify";
import { BODY_LIMIT } from "./config";
import type { AppContext } from "./context";
import { corsOptions } from "./plugins/cors";
import { registerErrorHandlers } from "./plugins/errorHandler";
import { registerOpenapi } from "./plugins/openapi";
import { registerRawBody } from "./plugins/rawBody";
import { Broadcaster } from "./realtime/Broadcaster";
import { registerAdminRoutes } from "./routes/admin.routes";
import { registerDocsRoutes } from "./routes/docs.routes";
import { registerMediaRoutes } from "./routes/media.routes";
import { registerMetaRoutes } from "./routes/meta.routes";
import { registerRenderRoutes } from "./routes/render.routes";
import { registerUploadRoutes } from "./routes/upload.routes";
import { registerWsRoutes } from "./routes/ws.routes";
import type { ChangeStore } from "./store/ChangeStore";
import { WebhookDispatcher } from "./webhooks/dispatcher";

export async function createApp(
  store: ChangeStore,
  opts: { dispatcher?: WebhookDispatcher } = {},
): Promise<{ app: FastifyInstance; server: Server; bcast: Broadcaster }> {
  const dispatcher = opts.dispatcher ?? new WebhookDispatcher(store);
  const bcast = new Broadcaster((docId, ctx) => {
    void dispatcher.onSessionEnded(docId, ctx);
  });
  const ctx: AppContext = { store, bcast, dispatcher };

  const app = Fastify({ bodyLimit: BODY_LIMIT, routerOptions: { ignoreTrailingSlash: true } });

  // Infrastructure: CORS + WebSocket support, body handling, error envelopes.
  await app.register(cors, corsOptions);
  await app.register(websocket);
  registerRawBody(app);
  registerErrorHandlers(app);

  // OpenAPI generation + Swagger UI — must precede the routes so their schemas
  // are captured into the spec.
  await registerOpenapi(app);

  // Presentation: one module per endpoint group.
  registerMetaRoutes(app);
  registerAdminRoutes(app, ctx);
  registerUploadRoutes(app, ctx);
  registerRenderRoutes(app, ctx);
  registerDocsRoutes(app, ctx);
  registerMediaRoutes(app, ctx);
  registerWsRoutes(app, ctx);

  await app.ready();
  return { app, server: app.server, bcast };
}
