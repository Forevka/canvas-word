import type { FastifyCorsOptions } from "@fastify/cors";

// Permissive CORS for every route + preflight (replaces the hand-rolled OPTIONS
// branch). x-api-key is allowed so browser integrations can upload with a key.
export const corsOptions: FastifyCorsOptions = {
  origin: "*",
  methods: ["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
  allowedHeaders: ["content-type", "content-encoding", "authorization", "x-filename", "x-user", "x-api-key"],
  exposedHeaders: ["content-disposition"],
};
