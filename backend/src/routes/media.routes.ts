// PUT/GET /media/:hash — content-addressed blob store. Bodies/responses are raw
// bytes, so they're described via consumes/produces rather than JSON schemas
// (a body/response schema would make Fastify validate/serialize the Buffer).
import { OCTET_STREAM_MIME } from "@cw/shared";
import type { FastifyInstance } from "fastify";
import type { AppContext } from "../context";
import { headerStr } from "../http/headers";

const hashParam = {
  type: "object",
  properties: { hash: { type: "string", description: "sha256 hex of the bytes" } },
  required: ["hash"],
} as const;

export function registerMediaRoutes(app: FastifyInstance, { store }: AppContext): void {
  app.put(
    "/media/:hash",
    {
      schema: {
        tags: ["media"],
        summary: "Store image bytes under their content hash",
        params: hashParam,
        consumes: [OCTET_STREAM_MIME],
        response: { 200: { type: "object", properties: { hash: { type: "string" } }, additionalProperties: true } },
      },
    },
    async (req) => {
      const { hash } = req.params as { hash: string };
      const bytes = (req.body as Buffer | undefined) ?? Buffer.alloc(0);
      const mime = headerStr(req.headers["content-type"]) ?? OCTET_STREAM_MIME;
      await store.putMedia({ hash, mime, bytes: new Uint8Array(bytes) });
      return { hash };
    },
  );

  app.get(
    "/media/:hash",
    {
      schema: {
        tags: ["media"],
        summary: "Fetch image bytes by content hash",
        params: hashParam,
        produces: [OCTET_STREAM_MIME],
        response: { 200: { type: "string", format: "binary" }, 404: { $ref: "Error#" } },
      },
    },
    async (req, reply) => {
      const { hash } = req.params as { hash: string };
      const rec = await store.getMedia(hash);
      if (!rec) return reply.code(404).send({ error: "media not found" });
      return reply.code(200).type(rec.mime).send(Buffer.from(rec.bytes));
    },
  );
}
