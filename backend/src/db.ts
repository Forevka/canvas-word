// Postgres connection. Reads DATABASE_URL (see env.example); defaults to the
// docker-compose Postgres so `docker compose up -d` + `npm run dev -w @cw/backend`
// works with no configuration.

import { Pool } from "pg";

// 127.0.0.1, not "localhost": Docker Desktop publishes the port on IPv4
// (0.0.0.0), while "localhost" can resolve to IPv6 ::1 first and refuse.
export const DATABASE_URL =
  process.env.DATABASE_URL ?? "postgres://postgres:postgres@127.0.0.1:5544/canvasword";

export function createPool(connectionString: string = DATABASE_URL): Pool {
  return new Pool({ connectionString });
}
