// Document CRUD + change log + review overlay + server-side export.
import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import {
  DOCX_MIME,
  PDF_MIME,
  reconstruct,
  rehydrateReview,
  type Change,
  type ReviewOpEnvelope,
  type SerializedDocument,
  type UserInfo,
} from "@cw/shared";
import type { ResolvedFontsConfig } from "@forevka/wordcanvas/export";
import type { AppContext } from "../context";
import { exportDoc } from "../export/serverExport";
import { sanitizeFilename } from "../http/headers";

// Reusable schema fragments.
const docIdParam = { type: "object", properties: { id: { type: "string", format: "uuid" } }, required: ["id"] } as const;
const freeObject = { type: "object", additionalProperties: true } as const;

export function registerDocsRoutes(app: FastifyInstance, { store, bcast }: AppContext): void {
  // POST /docs — create a document from a base snapshot. Body is either a bare
  // SerializedDocument, or { snapshot, createdBy?, user? } for attribution.
  app.post(
    "/docs",
    {
      schema: {
        tags: ["docs"],
        summary: "Create a document from a base snapshot",
        description:
          "Mints a new document id (server-side UUID) and stores the snapshot as version 0. Body is " +
          "either a bare SerializedDocument, or { snapshot, createdBy?, user?, fontsConfig? } — the " +
          "fontsConfig (custom families with face URLs + sizing) is saved so server-side export embeds them.",
        body: { type: "object", additionalProperties: true },
        response: {
          201: {
            type: "object",
            properties: { docId: { type: "string", format: "uuid" }, version: { type: "integer" } },
            additionalProperties: true,
          },
        },
      },
    },
    async (req, reply) => {
      const body = req.body as
        | SerializedDocument
        | { snapshot: SerializedDocument; createdBy?: string; user?: UserInfo; fontsConfig?: ResolvedFontsConfig };
      const wrapped = !!body && typeof body === "object" && "snapshot" in body;
      const snapshot = (wrapped ? body.snapshot : body) as SerializedDocument;
      const createdBy = wrapped ? body.createdBy : undefined;
      const user = wrapped ? body.user : undefined;
      const fontsConfig = wrapped ? body.fontsConfig : undefined;
      if (user) await store.upsertUser(user);
      const created = await store.createDocument(snapshot, {
        ...(createdBy ? { createdBy } : {}),
        ...(fontsConfig ? { fontsConfig } : {}),
      });
      return reply.code(201).send(created);
    },
  );

  // GET /docs/:id — newest snapshot + head version.
  app.get(
    "/docs/:id",
    {
      schema: {
        tags: ["docs"],
        summary: "Newest snapshot + head version",
        params: docIdParam,
        response: {
          200: {
            type: "object",
            properties: {
              docId: { type: "string" },
              version: { type: "integer", description: "version the snapshot represents" },
              head: { type: "integer", description: "latest version (accepted change count)" },
              snapshot: { $ref: "SerializedDocument#" },
            },
            additionalProperties: true,
          },
          404: { $ref: "Error#" },
        },
      },
    },
    async (req, reply) => {
      const { id } = req.params as { id: string };
      const snap = await store.getSnapshot(id);
      if (!snap) return reply.code(404).send({ error: "document not found" });
      const head = await store.getHead(id);
      return { docId: id, version: snap.version, head, snapshot: snap.snapshot };
    },
  );

  // GET /docs/:id/changes?since=N
  app.get(
    "/docs/:id/changes",
    {
      schema: {
        tags: ["docs"],
        summary: "Changes since a sequence number (ordered)",
        params: docIdParam,
        querystring: {
          type: "object",
          properties: { since: { type: "integer", default: 0, description: "return changes with seq >= since" } },
        },
        response: {
          200: {
            type: "object",
            properties: {
              docId: { type: "string" },
              since: { type: "integer" },
              changes: { type: "array", items: { $ref: "Change#" } },
            },
            additionalProperties: true,
          },
        },
      },
    },
    async (req) => {
      const { id } = req.params as { id: string };
      const since = Number((req.query as { since?: string }).since ?? 0);
      const changes = await store.getChanges(id, since);
      return { docId: id, since, changes };
    },
  );

  // POST /docs/:id/changes — append a change (OT-rebased onto head), broadcast it.
  app.post(
    "/docs/:id/changes",
    {
      schema: {
        tags: ["docs"],
        summary: "Append a change (OT-rebased onto head), returns the accepted change",
        description:
          "If the change's baseVersion is behind head, the server transforms its ops against the changes " +
          "committed since, assigns the canonical seq, and persists + broadcasts it. Idempotent on change.id.",
        params: docIdParam,
        body: { $ref: "Change#" },
        response: { 200: { $ref: "Change#" } },
      },
    },
    async (req) => {
      const { id } = req.params as { id: string };
      const accepted = await store.appendChange(id, req.body as Change);
      bcast.publish(id, accepted);
      return accepted;
    },
  );

  // GET /docs/:id/document — reconstruct the latest version (convenience).
  app.get(
    "/docs/:id/document",
    {
      schema: {
        tags: ["docs"],
        summary: "Reconstruct the latest document (snapshot + replayed log)",
        params: docIdParam,
        response: {
          200: {
            type: "object",
            properties: { docId: { type: "string" }, version: { type: "integer" }, doc: { $ref: "Document#" } },
            additionalProperties: true,
          },
          404: { $ref: "Error#" },
        },
      },
    },
    async (req, reply) => {
      const { id } = req.params as { id: string };
      const snap = await store.getSnapshot(id);
      if (!snap) return reply.code(404).send({ error: "document not found" });
      const changes = await store.getChanges(id, snap.version);
      const doc = reconstruct(snap.snapshot, changes);
      return { docId: id, version: snap.version + changes.length, doc };
    },
  );

  // GET /docs/:id/review — the review overlay (track changes + comments),
  // rehydrated to HEAD: the op log folded with anchors fast-forwarded through
  // the change log, so a joiner sees every suggestion/comment correctly placed.
  app.get(
    "/docs/:id/review",
    {
      schema: {
        tags: ["docs"],
        summary: "Review overlay (track changes + comments) rehydrated to HEAD",
        params: docIdParam,
        response: {
          200: {
            type: "object",
            properties: { docId: { type: "string" }, version: { type: "integer" }, layer: freeObject },
            additionalProperties: true,
          },
          404: { $ref: "Error#" },
        },
      },
    },
    async (req, reply) => {
      const { id } = req.params as { id: string };
      const snap = await store.getSnapshot(id);
      if (!snap) return reply.code(404).send({ error: "document not found" });
      const [changes, reviewOps] = await Promise.all([
        store.getChanges(id, snap.version),
        store.getReviewOps(id),
      ]);
      const layer = rehydrateReview(id, snap.snapshot, snap.version, changes, reviewOps);
      return { docId: id, version: snap.version + changes.length, layer };
    },
  );

  // POST /docs/:id/review/ops — append a review op (REST fallback for offline
  // clients; the live path is the WebSocket "review" message).
  app.post(
    "/docs/:id/review/ops",
    {
      schema: {
        tags: ["docs"],
        summary: "Append a review op (suggestion/comment) — REST fallback",
        params: docIdParam,
        body: { type: "object", additionalProperties: true, description: "a ReviewOpEnvelope (see @cw/shared)" },
        response: { 200: { type: "object", properties: { ok: { type: "boolean" } }, additionalProperties: true } },
      },
    },
    async (req) => {
      const { id } = req.params as { id: string };
      await store.appendReviewOp(id, req.body as ReviewOpEnvelope);
      return { ok: true };
    },
  );

  // GET /docs/:id/activity — edit history with author names (attribution).
  app.get(
    "/docs/:id/activity",
    {
      schema: {
        tags: ["docs"],
        summary: "Edit history with author names (attribution)",
        params: docIdParam,
        response: {
          200: {
            type: "object",
            properties: {
              docId: { type: "string" },
              createdBy: { type: ["string", "null"] },
              creatorFirstName: { type: "string" },
              creatorLastName: { type: "string" },
              createdAt: { type: "integer" },
              entries: { type: "array", items: freeObject },
            },
            additionalProperties: true,
          },
          404: { $ref: "Error#" },
        },
      },
    },
    async (req, reply) => {
      const { id } = req.params as { id: string };
      const activity = await store.getActivity(id);
      if (!activity) return reply.code(404).send({ error: "document not found" });
      return activity;
    },
  );

  // GET /docs/:id/export.docx | /docs/:id/export.pdf — server-side download.
  // These are the URLs the session-ended webhook hands to integrations. The 200
  // is a binary stream (documented via `produces`), so no response schema is set
  // — that would make fast-json-stringify mangle the bytes.
  const exportSchema = (format: "docx" | "pdf") =>
    ({
      tags: ["docs"],
      summary: `Export the document as ${format.toUpperCase()}`,
      params: docIdParam,
      querystring: {
        type: "object",
        properties: { tracked: { type: "string", enum: ["accept", "reject"], description: "fold suggestions (default reject)" } },
      },
      produces: [format === "pdf" ? PDF_MIME : DOCX_MIME],
      response: { 200: { type: "string", format: "binary" }, 404: { $ref: "Error#" } },
    }) as const;
  const exportHandler = (format: "docx" | "pdf") => async (req: FastifyRequest, reply: FastifyReply) => {
    const { id } = req.params as { id: string };
    // ?tracked=accept folds suggestions as accepted; default rejects (baseline).
    const bakeMode = (req.query as { tracked?: string }).tracked === "accept" ? "accept" : "reject";
    const exported = await exportDoc(id, format, store, bakeMode);
    if (!exported) return reply.code(404).send({ error: "document not found" });
    const filename = `${sanitizeFilename(exported.title ?? id)}.${format}`;
    const mime = format === "pdf" ? PDF_MIME : DOCX_MIME;
    return reply
      .code(200)
      .header("content-disposition", `attachment; filename="${filename}"`)
      .type(mime)
      .send(Buffer.from(exported.bytes));
  };
  app.get("/docs/:id/export.docx", { schema: exportSchema("docx") }, exportHandler("docx"));
  app.get("/docs/:id/export.pdf", { schema: exportSchema("pdf") }, exportHandler("pdf"));
}
