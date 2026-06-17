// POST /render.pdf — stateless headless render for a renderer-less producer (e.g. a
// C# pipeline that emits raw .docx but has no layout engine). We run our layout
// engine to: BUILD the TOC field's entries from the document's headings (styled via
// the optional `toc` part) and auto-compute footer PAGE/NUMPAGES per page, then
// return a PDF. Accepts multipart (file=docx, toc=JSON TocOptions) OR a raw .docx
// body. Nothing is stored. The 200 is binary (described via produces), so it carries
// no response schema; only the JSON error responses do.
import type { FastifyInstance } from "fastify";
import type { TocOptions } from "@cw/shared";
import type { AppContext } from "../context";
import { renderPdfFromDocx } from "../export/serverExport";
import { uploadAuthHook } from "../http/auth";
import { headerStr } from "../http/headers";
import { parseMultipart } from "../http/multipart";

export function registerRenderRoutes(app: FastifyInstance, { store }: AppContext): void {
  app.post(
    "/render.pdf",
    {
      preHandler: uploadAuthHook(store),
      schema: {
        tags: ["render"],
        summary: "Render a .docx to PDF headlessly (TOC + page numbers computed)",
        description:
          "Render a raw .docx to PDF with the server's layout engine — the path for a renderer-less producer " +
          "(e.g. a C# pipeline that emits .docx but cannot compute layout). The engine BUILDS the TOC field's " +
          "entries from the document's headings (styled by the optional `toc` part — see TocOptions in " +
          "@cw/shared) and auto-computes footer PAGE/NUMPAGES per page. No document is stored. Auth: dashboard " +
          "admin token OR an integration API key (X-API-Key or Bearer). Accepts multipart (file + optional toc) " +
          "OR a raw .docx body.",
        security: [{ ApiKeyAuth: [] }, { BearerAuth: [] }],
        consumes: ["multipart/form-data", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
        produces: ["application/pdf"],
        response: { 200: { type: "string", format: "binary" }, 400: { $ref: "Error#" }, 401: { $ref: "Error#" } },
      },
    },
    async (req, reply) => {
      const ctype = headerStr(req.headers["content-type"]) ?? "";
      const body = (req.body as Buffer | undefined) ?? Buffer.alloc(0);
      let docx: Buffer | undefined;
      let opts: TocOptions = {};
      if (/multipart\/form-data/i.test(ctype)) {
        const fields = parseMultipart(ctype, body);
        docx = fields["file"]?.data;
        const tocJson = fields["toc"]?.data?.toString("utf8");
        if (tocJson) {
          try {
            opts = JSON.parse(tocJson) as TocOptions;
          } catch {
            return reply.code(400).send({ error: "invalid toc JSON" });
          }
        }
      } else {
        docx = body; // raw .docx body, default options
      }
      if (!docx || docx.length === 0) {
        return reply.code(400).send({ error: "missing docx (multipart 'file' part or raw body)" });
      }
      try {
        const out = await renderPdfFromDocx(new Uint8Array(docx), opts);
        return reply
          .code(200)
          .header("content-disposition", `attachment; filename="render.pdf"`)
          .type("application/pdf")
          .send(Buffer.from(out));
      } catch (e) {
        return reply.code(400).send({ error: String(e instanceof Error ? e.message : e) });
      }
    },
  );
}
