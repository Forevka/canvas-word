// Collaboration server: HTTP for document/change/media CRUD + WebSocket for live
// change broadcast. The store assigns each change its canonical seq; the server
// is the single source of order (the ShareDB-style seam where OT rebasing lands
// in Phase 4). Reconstructing a version is base snapshot + replay (shared).

import { createServer, type IncomingMessage, type Server, type ServerResponse } from "node:http";
import { WebSocketServer, type WebSocket } from "ws";
import { reconstruct, type Change, type SerializedDocument } from "@cw/shared";
import { createPool } from "./db";
import { PgChangeStore, type ChangeStore } from "./store/ChangeStore";

const PORT = Number(process.env.BACKEND_PORT ?? 8787);

// --- tiny request helpers ---------------------------------------------------

function sendJson(res: ServerResponse, code: number, body: unknown): void {
  const text = JSON.stringify(body);
  res.writeHead(code, { "content-type": "application/json", "access-control-allow-origin": "*" });
  res.end(text);
}

async function readBody(req: IncomingMessage): Promise<Buffer> {
  const chunks: Buffer[] = [];
  for await (const c of req) chunks.push(c as Buffer);
  return Buffer.concat(chunks);
}

async function readJson<T>(req: IncomingMessage): Promise<T> {
  const buf = await readBody(req);
  return JSON.parse(buf.toString("utf8")) as T;
}

// --- live broadcast ---------------------------------------------------------

class Broadcaster {
  private readonly rooms = new Map<string, Set<WebSocket>>();

  join(docId: string, ws: WebSocket): void {
    let room = this.rooms.get(docId);
    if (!room) this.rooms.set(docId, (room = new Set()));
    room.add(ws);
    ws.on("close", () => room!.delete(ws));
  }

  publish(docId: string, change: Change): void {
    const room = this.rooms.get(docId);
    if (!room) return;
    const msg = JSON.stringify({ type: "change", change });
    for (const ws of room) if (ws.readyState === ws.OPEN) ws.send(msg);
  }
}

// --- app --------------------------------------------------------------------

export function createApp(store: ChangeStore): { server: Server; bcast: Broadcaster } {
  const bcast = new Broadcaster();

  const server = createServer((req, res) => {
    handle(req, res, store, bcast).catch((e) => {
      sendJson(res, 500, { error: String(e instanceof Error ? e.message : e) });
    });
  });

  const wss = new WebSocketServer({ server, path: "/ws" });
  wss.on("connection", (ws, req) => {
    const docId = new URL(req.url ?? "", "http://x").searchParams.get("doc");
    if (!docId) return ws.close(1008, "missing ?doc");
    bcast.join(docId, ws);
  });

  return { server, bcast };
}

async function handle(
  req: IncomingMessage,
  res: ServerResponse,
  store: ChangeStore,
  bcast: Broadcaster,
): Promise<void> {
  const method = req.method ?? "GET";
  const url = new URL(req.url ?? "/", "http://x");
  const parts = url.pathname.split("/").filter(Boolean);

  if (method === "OPTIONS") {
    res.writeHead(204, {
      "access-control-allow-origin": "*",
      "access-control-allow-methods": "GET,POST,PUT,OPTIONS",
      "access-control-allow-headers": "content-type",
    });
    res.end();
    return;
  }

  // POST /docs  — create a document from a base snapshot
  if (method === "POST" && parts.length === 1 && parts[0] === "docs") {
    const base = await readJson<SerializedDocument>(req);
    const created = await store.createDocument(base);
    return sendJson(res, 201, created);
  }

  // /docs/:id ...
  if (parts[0] === "docs" && parts[1]) {
    const docId = parts[1];

    // GET /docs/:id  — newest snapshot + head version
    if (method === "GET" && parts.length === 2) {
      const snap = await store.getSnapshot(docId);
      if (!snap) return sendJson(res, 404, { error: "document not found" });
      const head = await store.getHead(docId);
      return sendJson(res, 200, { docId, version: snap.version, head, snapshot: snap.snapshot });
    }

    // GET /docs/:id/changes?since=N
    if (method === "GET" && parts[2] === "changes") {
      const since = Number(url.searchParams.get("since") ?? 0);
      const changes = await store.getChanges(docId, since);
      return sendJson(res, 200, { docId, since, changes });
    }

    // POST /docs/:id/changes  — append a change, broadcast it
    if (method === "POST" && parts[2] === "changes") {
      const change = await readJson<Change>(req);
      const accepted = await store.appendChange(docId, change);
      bcast.publish(docId, accepted);
      return sendJson(res, 200, accepted);
    }

    // GET /docs/:id/document  — reconstruct the latest version (convenience)
    if (method === "GET" && parts[2] === "document") {
      const snap = await store.getSnapshot(docId);
      if (!snap) return sendJson(res, 404, { error: "document not found" });
      const changes = await store.getChanges(docId, snap.version);
      const doc = reconstruct(snap.snapshot, changes);
      return sendJson(res, 200, { docId, version: (snap.version + changes.length), doc });
    }
  }

  // /media/:hash
  if (parts[0] === "media" && parts[1]) {
    const hash = parts[1];
    if (method === "PUT") {
      const bytes = await readBody(req);
      const mime = req.headers["content-type"] ?? "application/octet-stream";
      await store.putMedia({ hash, mime, bytes: new Uint8Array(bytes) });
      return sendJson(res, 200, { hash });
    }
    if (method === "GET") {
      const rec = await store.getMedia(hash);
      if (!rec) return sendJson(res, 404, { error: "media not found" });
      res.writeHead(200, { "content-type": rec.mime, "access-control-allow-origin": "*" });
      res.end(Buffer.from(rec.bytes));
      return;
    }
  }

  sendJson(res, 404, { error: "not found" });
}

// Entry point (tsx src/server.ts). Tests import createApp + their own store.
if (import.meta.url === `file://${process.argv[1]}` || process.argv[1]?.endsWith("server.ts")) {
  const pool = createPool();
  const { server } = createApp(new PgChangeStore(pool));
  server.listen(PORT, () => {
    console.log(`[cw-backend] listening on http://localhost:${PORT} (ws: /ws?doc=<id>)`);
  });
}
