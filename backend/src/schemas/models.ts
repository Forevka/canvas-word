// Shared JSON Schemas registered once via app.addSchema, so every route can
// reference them by $id ({ $ref: "Change#" }) for both validation/serialization
// AND OpenAPI generation. The document model itself is large and owned by the
// client (@cw/shared), so it's described as a free-form object — additionalProperties
// stays open so serialization never strips fields the frontend reads.
import type { FastifyInstance } from "fastify";

const Error = {
  $id: "Error",
  type: "object",
  properties: { error: { type: "string" } },
  required: ["error"],
} as const;

const Document = {
  $id: "Document",
  type: "object",
  description: "canvas-word document model (see @cw/shared Document)",
  additionalProperties: true,
} as const;

const SerializedDocument = {
  $id: "SerializedDocument",
  type: "object",
  properties: {
    version: { type: "integer", description: "snapshot format version" },
    doc: { $ref: "Document#" },
  },
  required: ["version", "doc"],
  additionalProperties: true,
} as const;

const Op = {
  $id: "Op",
  type: "object",
  description: "a single operation (see @cw/shared Op union: insertText, deleteRange, splitParagraph, …)",
  additionalProperties: true,
} as const;

const Change = {
  $id: "Change",
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
    ops: { type: "array", items: { $ref: "Op#" } },
    selectionAfter: { type: ["object", "null"], additionalProperties: true },
  },
  required: ["id", "docId", "baseVersion", "siteId", "origin", "ts", "ops"],
  additionalProperties: true,
} as const;

/** Register every shared schema. Must run before routes that $ref them. */
export function registerSchemas(app: FastifyInstance): void {
  for (const schema of [Error, Document, SerializedDocument, Op, Change]) {
    app.addSchema(schema);
  }
}
