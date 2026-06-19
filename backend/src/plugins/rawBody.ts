// Body handling shared by every route. The collaboration clients send three body
// shapes: JSON (often gzipped), raw binary (media bytes, a bare .docx), and
// multipart (render.pdf). This installs:
//   - a content-type default so a bare-body request still parses;
//   - transparent gzip inflation;
//   - a raw-Buffer catch-all for everything that isn't JSON.
import { createGunzip } from "node:zlib";
import { OCTET_STREAM_MIME } from "@cw/shared";
import type { FastifyInstance } from "fastify";

export function registerRawBody(app: FastifyInstance): void {
  // Default a body-bearing request with no content-type to octet-stream so the
  // catch-all parser below claims it (e.g. a bare .docx upload).
  app.addHook("onRequest", (req, _reply, done) => {
    if (!req.headers["content-type"] && req.method !== "GET" && req.method !== "HEAD") {
      req.headers["content-type"] = OCTET_STREAM_MIME;
    }
    done();
  });

  // Transparently inflate gzipped bodies (clients gzip the large JSON snapshot).
  // Drop the now-stale length/encoding so the parser reads the inflated stream to
  // its true end instead of rejecting the size mismatch.
  app.addHook("preParsing", (req, _reply, payload, done) => {
    if ((req.headers["content-encoding"] ?? "").toLowerCase() === "gzip") {
      delete req.headers["content-length"];
      delete req.headers["content-encoding"];
      done(null, payload.pipe(createGunzip()));
    } else {
      done(null, payload);
    }
  });

  // Raw-binary catch-all: media bytes, .docx uploads, and multipart (render.pdf).
  // JSON keeps Fastify's built-in parser; this only claims everything else.
  app.addContentTypeParser("*", { parseAs: "buffer" }, (_req, body, done) => done(null, body));
}
