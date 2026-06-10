// Collaboration server: HTTP for document/change/media CRUD + WebSocket for live
// change broadcast. The store assigns each change its canonical seq; the server
// is the single source of order (the ShareDB-style seam where OT rebasing lands
// in Phase 4). Reconstructing a version is base snapshot + replay (shared).

import { createServer, type IncomingMessage, type Server, type ServerResponse } from "node:http";
import { gunzipSync } from "node:zlib";
import { WebSocketServer, type WebSocket } from "ws";
import { reconstruct, type Change, type SerializedDocument, type UserInfo } from "@cw/shared";
import { createPool } from "./db";
import { OPENAPI_SPEC, SWAGGER_HTML } from "./openapi";
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
  const raw = Buffer.concat(chunks);
  // Transparently inflate gzipped bodies (clients gzip the large JSON snapshot;
  // Content-Encoding: gzip). Media PUTs send raw bytes — already-compressed images.
  if (raw.length > 0 && (req.headers["content-encoding"] ?? "").toLowerCase() === "gzip") {
    return gunzipSync(raw);
  }
  return raw;
}

async function readJson<T>(req: IncomingMessage): Promise<T> {
  const buf = await readBody(req);
  return JSON.parse(buf.toString("utf8")) as T;
}

// --- live broadcast + presence ----------------------------------------------

interface Conn {
  docId: string;
  siteId?: string;
  user?: UserInfo | undefined;
  selection?: unknown;
}

class Broadcaster {
  private readonly rooms = new Map<string, Set<WebSocket>>();
  private readonly meta = new Map<WebSocket, Conn>();

  join(docId: string, ws: WebSocket): void {
    let room = this.rooms.get(docId);
    if (!room) this.rooms.set(docId, (room = new Set()));
    room.add(ws);
    this.meta.set(ws, { docId });
    ws.on("close", () => {
      room!.delete(ws);
      const m = this.meta.get(ws);
      this.meta.delete(ws);
      // Tell the room this collaborator left so their caret disappears.
      if (m?.siteId) this.toRoom(docId, { type: "leave", siteId: m.siteId }, ws);
    });
  }

  /** A client identified itself: send it the current roster, announce it. */
  hello(ws: WebSocket, siteId: string, user?: UserInfo): void {
    const m = this.meta.get(ws);
    if (!m) return;
    m.siteId = siteId;
    m.user = user;
    const entries = this.roster(m.docId, ws);
    if (ws.readyState === ws.OPEN) ws.send(JSON.stringify({ type: "roster", entries }));
    this.toRoom(m.docId, { type: "presence", siteId, user, selection: m.selection ?? null }, ws);
  }

  /** A client's caret moved: relay to the rest of the room. */
  presence(ws: WebSocket, selection: unknown): void {
    const m = this.meta.get(ws);
    if (!m?.siteId) return;
    m.selection = selection;
    this.toRoom(m.docId, { type: "presence", siteId: m.siteId, user: m.user, selection }, ws);
  }

  /** Accepted change → everyone in the room (incl. sender, who treats it as ack). */
  publish(docId: string, change: Change): void {
    this.toRoom(docId, { type: "change", change });
  }

  private roster(
    docId: string,
    exclude: WebSocket,
  ): Array<{ siteId: string; user?: UserInfo | undefined; selection: unknown }> {
    const room = this.rooms.get(docId);
    const out: Array<{ siteId: string; user?: UserInfo | undefined; selection: unknown }> = [];
    if (!room) return out;
    for (const ws of room) {
      if (ws === exclude) continue;
      const m = this.meta.get(ws);
      if (m?.siteId) out.push({ siteId: m.siteId, user: m.user, selection: m.selection ?? null });
    }
    return out;
  }

  private toRoom(docId: string, msg: unknown, exclude?: WebSocket): void {
    const room = this.rooms.get(docId);
    if (!room) return;
    const s = JSON.stringify(msg);
    for (const ws of room) {
      if (ws === exclude) continue;
      if (ws.readyState === ws.OPEN) ws.send(s);
    }
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
    ws.on("message", (data) => {
      void (async () => {
        let msg: { type?: string; change?: Change; siteId?: string; user?: UserInfo; selection?: unknown };
        try {
          msg = JSON.parse(String(data));
        } catch {
          return;
        }
        if (msg.type === "submit" && msg.change) {
          // Append (with OT rebase) then broadcast to everyone in the room,
          // including the sender — whose client treats the echo as its ack.
          const accepted = await store.appendChange(docId, msg.change);
          bcast.publish(docId, accepted);
        } else if (msg.type === "hello" && msg.siteId) {
          if (msg.user) await store.upsertUser(msg.user);
          bcast.hello(ws, msg.siteId, msg.user);
        } else if (msg.type === "presence") {
          bcast.presence(ws, msg.selection ?? null);
        }
      })();
    });
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
      "access-control-allow-headers": "content-type,content-encoding",
    });
    res.end();
    return;
  }

  // API docs (distinct paths so they don't collide with /docs/:id).
  if (method === "GET" && url.pathname === "/openapi.json") {
    return sendJson(res, 200, OPENAPI_SPEC);
  }
  if (method === "GET" && (url.pathname === "/swagger" || url.pathname === "/swagger/")) {
    res.writeHead(200, { "content-type": "text/html; charset=utf-8", "access-control-allow-origin": "*" });
    res.end(SWAGGER_HTML);
    return;
  }

  // POST /docs  — create a document from a base snapshot. Body is either a bare
  // SerializedDocument, or { snapshot, createdBy?, user? } for attribution.
  if (method === "POST" && parts.length === 1 && parts[0] === "docs") {
    const body = await readJson<
      SerializedDocument | { snapshot: SerializedDocument; createdBy?: string; user?: UserInfo }
    >(req);
    const wrapped = body && typeof body === "object" && "snapshot" in body;
    const snapshot = (wrapped ? body.snapshot : body) as SerializedDocument;
    const createdBy = wrapped ? body.createdBy : undefined;
    const user = wrapped ? body.user : undefined;
    if (user) await store.upsertUser(user);
    const created = await store.createDocument(snapshot, createdBy ? { createdBy } : {});
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

    // GET /docs/:id/activity  — edit history with author names (attribution)
    if (method === "GET" && parts[2] === "activity") {
      const activity = await store.getActivity(docId);
      if (!activity) return sendJson(res, 404, { error: "document not found" });
      return sendJson(res, 200, activity);
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
