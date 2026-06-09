// Hand-authored OpenAPI 3.1 spec for the collaboration API. Small and stable, so
// a generated spec isn't worth a framework; served as-is at GET /openapi.json and
// rendered by Swagger UI at GET /swagger.

export const OPENAPI_SPEC = {
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
  servers: [{ url: "http://localhost:8787", description: "local dev" }],
  paths: {
    "/docs": {
      post: {
        summary: "Create a document from a base snapshot",
        description: "Mints a new document id (server-side UUID) and stores the snapshot as version 0.",
        requestBody: {
          required: true,
          content: { "application/json": { schema: { $ref: "#/components/schemas/SerializedDocument" } } },
        },
        responses: {
          "201": {
            description: "Created",
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  properties: { docId: { type: "string", format: "uuid" }, version: { type: "integer" } },
                  required: ["docId", "version"],
                },
              },
            },
          },
        },
      },
    },
    "/docs/{id}": {
      get: {
        summary: "Newest snapshot + head version",
        parameters: [{ $ref: "#/components/parameters/DocId" }],
        responses: {
          "200": {
            description: "OK",
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  properties: {
                    docId: { type: "string" },
                    version: { type: "integer", description: "version the snapshot represents" },
                    head: { type: "integer", description: "latest version (accepted change count)" },
                    snapshot: { $ref: "#/components/schemas/SerializedDocument" },
                  },
                },
              },
            },
          },
          "404": { $ref: "#/components/responses/NotFound" },
        },
      },
    },
    "/docs/{id}/changes": {
      get: {
        summary: "Changes since a sequence number (ordered)",
        parameters: [
          { $ref: "#/components/parameters/DocId" },
          { name: "since", in: "query", schema: { type: "integer", default: 0 }, description: "return changes with seq >= since" },
        ],
        responses: {
          "200": {
            description: "OK",
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  properties: {
                    docId: { type: "string" },
                    since: { type: "integer" },
                    changes: { type: "array", items: { $ref: "#/components/schemas/Change" } },
                  },
                },
              },
            },
          },
        },
      },
      post: {
        summary: "Append a change (OT-rebased onto head), returns the accepted change",
        description:
          "If the change's baseVersion is behind head, the server transforms its ops against the " +
          "changes committed since, assigns the canonical seq, and persists + broadcasts it. " +
          "Idempotent on change.id.",
        parameters: [{ $ref: "#/components/parameters/DocId" }],
        requestBody: {
          required: true,
          content: { "application/json": { schema: { $ref: "#/components/schemas/Change" } } },
        },
        responses: {
          "200": {
            description: "Accepted (with assigned seq and possibly rebased ops)",
            content: { "application/json": { schema: { $ref: "#/components/schemas/Change" } } },
          },
        },
      },
    },
    "/docs/{id}/document": {
      get: {
        summary: "Reconstruct the latest document (snapshot + replayed log)",
        parameters: [{ $ref: "#/components/parameters/DocId" }],
        responses: {
          "200": {
            description: "OK",
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  properties: {
                    docId: { type: "string" },
                    version: { type: "integer" },
                    doc: { $ref: "#/components/schemas/Document" },
                  },
                },
              },
            },
          },
          "404": { $ref: "#/components/responses/NotFound" },
        },
      },
    },
    "/docs/{id}/activity": {
      get: {
        summary: "Edit history with author names (attribution)",
        parameters: [{ $ref: "#/components/parameters/DocId" }],
        responses: {
          "200": {
            description: "OK",
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  properties: {
                    docId: { type: "string" },
                    createdBy: { type: ["string", "null"] },
                    creatorFirstName: { type: "string" },
                    creatorLastName: { type: "string" },
                    createdAt: { type: "integer" },
                    entries: {
                      type: "array",
                      items: {
                        type: "object",
                        properties: {
                          seq: { type: "integer" },
                          userId: { type: ["string", "null"] },
                          firstName: { type: "string" },
                          lastName: { type: "string" },
                          ts: { type: "integer" },
                          origin: { type: "string" },
                          opCount: { type: "integer" },
                        },
                      },
                    },
                  },
                },
              },
            },
          },
          "404": { $ref: "#/components/responses/NotFound" },
        },
      },
    },
    "/media/{hash}": {
      put: {
        summary: "Store image bytes under their content hash",
        parameters: [{ $ref: "#/components/parameters/MediaHash" }],
        requestBody: {
          required: true,
          content: { "application/octet-stream": { schema: { type: "string", format: "binary" } } },
        },
        responses: {
          "200": {
            description: "Stored",
            content: { "application/json": { schema: { type: "object", properties: { hash: { type: "string" } } } } },
          },
        },
      },
      get: {
        summary: "Fetch image bytes by content hash",
        parameters: [{ $ref: "#/components/parameters/MediaHash" }],
        responses: {
          "200": { description: "Image bytes", content: { "application/octet-stream": { schema: { type: "string", format: "binary" } } } },
          "404": { $ref: "#/components/responses/NotFound" },
        },
      },
    },
  },
  components: {
    parameters: {
      DocId: { name: "id", in: "path", required: true, schema: { type: "string", format: "uuid" } },
      MediaHash: { name: "hash", in: "path", required: true, schema: { type: "string" }, description: "sha256 hex of the bytes" },
    },
    responses: {
      NotFound: {
        description: "Not found",
        content: { "application/json": { schema: { $ref: "#/components/schemas/Error" } } },
      },
    },
    schemas: {
      Error: { type: "object", properties: { error: { type: "string" } }, required: ["error"] },
      // The document model is large; described as a free-form object (the client
      // owns its exact shape via @cw/shared types).
      Document: { type: "object", description: "canvas-word document model (see @cw/shared Document)", additionalProperties: true },
      SerializedDocument: {
        type: "object",
        properties: { version: { type: "integer", description: "snapshot format version" }, doc: { $ref: "#/components/schemas/Document" } },
        required: ["version", "doc"],
      },
      Op: { type: "object", description: "a single operation (see @cw/shared Op union: insertText, deleteRange, splitParagraph, …)", additionalProperties: true },
      Change: {
        type: "object",
        description: "one entry in the change log (forward ops only)",
        properties: {
          id: { type: "string", description: "client-minted unique id (idempotency key)" },
          docId: { type: "string" },
          baseVersion: { type: "integer", description: "version the change was generated against" },
          seq: { type: "integer", description: "canonical order, server-assigned" },
          siteId: { type: "string" },
          userId: { type: "string", description: "caller-supplied author id (attribution)" },
          origin: { type: "string", enum: ["typing", "command", "paste", "undo", "redo"] },
          ts: { type: "integer", description: "author epoch ms" },
          ops: { type: "array", items: { $ref: "#/components/schemas/Op" } },
          selectionAfter: { type: ["object", "null"], additionalProperties: true },
        },
        required: ["id", "docId", "baseVersion", "siteId", "origin", "ts", "ops"],
      },
    },
  },
} as const;

/** Minimal HTML that renders Swagger UI (from CDN) against /openapi.json. */
export const SWAGGER_HTML = `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>canvas-word API</title>
    <link rel="stylesheet" href="https://unpkg.com/swagger-ui-dist@5/swagger-ui.css" />
  </head>
  <body>
    <div id="swagger"></div>
    <script src="https://unpkg.com/swagger-ui-dist@5/swagger-ui-bundle.js" crossorigin></script>
    <script>
      window.ui = SwaggerUIBundle({ url: "/openapi.json", dom_id: "#swagger" });
    </script>
  </body>
</html>`;
