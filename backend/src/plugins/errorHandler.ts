// Uniform error / 404 envelopes — every failure returns { error: string }. The
// error handler respects a thrown error's statusCode (so a 400 body-parse error
// stays a 400) and defaults everything else to 500.
import type { FastifyInstance } from "fastify";

export function registerErrorHandlers(app: FastifyInstance): void {
  app.setErrorHandler((err: { statusCode?: number; message?: string }, _req, reply) => {
    const code = typeof err.statusCode === "number" && err.statusCode >= 400 ? err.statusCode : 500;
    void reply.code(code).send({ error: String(err instanceof Error ? err.message : (err.message ?? err)) });
  });
  app.setNotFoundHandler((_req, reply) => {
    void reply.code(404).send({ error: "not found" });
  });
}
