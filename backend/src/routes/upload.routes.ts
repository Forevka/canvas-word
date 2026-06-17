// POST /upload — parse a .docx server-side, store it, return its docId so the
// caller can redirect into WordCanvas({ docId }). Body = raw file bytes (described
// via consumes, not a body schema, since our parser yields a Buffer); filename +
// optional user come from headers.
import type { FastifyInstance } from "fastify";
import type { AppContext } from "../context";
import { uploadAuthHook } from "../http/auth";
import { headerStr, safeParseUser } from "../http/headers";
import { parseAndStore } from "../import/serverImport";

export function registerUploadRoutes(app: FastifyInstance, { store }: AppContext): void {
  app.post(
    "/upload",
    {
      preHandler: uploadAuthHook(store),
      schema: {
        tags: ["docs"],
        summary: "Upload a .docx; parse + store it server-side, returns its docId",
        description:
          "Body = raw .docx bytes. Auth: dashboard admin token OR an integration API key (X-API-Key or " +
          "Bearer). Filename comes from the X-Filename header; optional X-User (JSON) attributes the creator.",
        security: [{ ApiKeyAuth: [] }, { BearerAuth: [] }],
        consumes: ["application/vnd.openxmlformats-officedocument.wordprocessingml.document", "application/octet-stream"],
        response: {
          201: { type: "object", properties: { docId: { type: "string", format: "uuid" } }, additionalProperties: true },
          400: { $ref: "Error#" },
          401: { $ref: "Error#" },
        },
      },
    },
    async (req, reply) => {
      const bytes = (req.body as Buffer | undefined) ?? Buffer.alloc(0);
      const filename = headerStr(req.headers["x-filename"]) ?? "Untitled.docx";
      const userHeader = headerStr(req.headers["x-user"]);
      const user = userHeader ? safeParseUser(userHeader) : undefined;
      try {
        const result = await parseAndStore(new Uint8Array(bytes), filename, store, user);
        return reply.code(201).send(result);
      } catch (e) {
        return reply.code(400).send({ error: String(e instanceof Error ? e.message : e) });
      }
    },
  );
}
