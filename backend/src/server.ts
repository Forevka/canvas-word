// Collaboration server: HTTP for document/change/media CRUD + WebSocket for live
// change broadcast. The store assigns each change its canonical seq; the server
// is the single source of order (the ShareDB-style seam where OT rebasing lands
// in Phase 4). Reconstructing a version is base snapshot + replay (shared).

import { randomBytes } from "node:crypto";
import { createServer, type IncomingMessage, type Server, type ServerResponse } from "node:http";
import { gunzipSync } from "node:zlib";
import { WebSocketServer, type WebSocket } from "ws";
import { reconstruct, rehydrateReview, type Change, type ReviewOpEnvelope, type SerializedDocument, type UserInfo } from "@cw/shared";
import { authFromRequest, checkCredentials, issueToken, requireAdmin } from "./admin/auth";
import { hashToken, newApiToken } from "./auth/apiToken";
import { createPool } from "./db";
import { exportDoc } from "./export/serverExport";
import { parseAndStore } from "./import/serverImport";
import { OPENAPI_SPEC, SWAGGER_HTML } from "./openapi";
import { PgChangeStore, type ChangeStore } from "./store/ChangeStore";
import { SESSION_ENDED_EVENT, WebhookDispatcher } from "./webhooks/dispatcher";

const PORT = Number(process.env.BACKEND_PORT ?? 8787);
// Uploads are token-guarded by default; set UPLOAD_PUBLIC=1 for anonymous upload.
const UPLOAD_REQUIRES_AUTH = process.env.UPLOAD_PUBLIC !== "1";

// --- small header / value helpers -------------------------------------------

function headerStr(h: string | string[] | undefined): string | undefined {
  return Array.isArray(h) ? h[0] : h;
}

function safeParseUser(raw: string): UserInfo | undefined {
  try {
    const u = JSON.parse(raw) as Partial<UserInfo>;
    if (u && typeof u.id === "string") {
      return { id: u.id, firstName: u.firstName ?? "", lastName: u.lastName ?? "" };
    }
  } catch {
    // ignore malformed header
  }
  return undefined;
}

/** Make a title safe for a Content-Disposition filename (no quotes/paths/control). */
function sanitizeFilename(name: string): string {
  const cleaned = name.replace(/[-"\\/:*?<>|]+/g, "_").trim();
  return cleaned.length > 0 ? cleaned.slice(0, 120) : "document";
}

const randomSecret = (): string => randomBytes(24).toString("hex");

/** Bearer value from the Authorization header (admin token OR an integration token). */
function bearerToken(req: IncomingMessage): string | undefined {
  const h = req.headers["authorization"];
  if (!h || Array.isArray(h)) return undefined;
  const m = /^Bearer\s+(.+)$/i.exec(h.trim());
  return m ? m[1] : undefined;
}

/** Upload auth: a dashboard admin token (UI) OR a 3rd-party integration token
 *  (sent as X-API-Key or Bearer). */
async function authorizeUpload(req: IncomingMessage, store: ChangeStore): Promise<boolean> {
  if (authFromRequest(req)) return true; // admin (dashboard) token
  const raw = headerStr(req.headers["x-api-key"]) ?? bearerToken(req);
  if (!raw) return false;
  return store.verifyApiToken(hashToken(raw));
}

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
  // Pending session-end timers, keyed by doc — set when a room empties, cleared
  // if someone (re)joins before the debounce elapses.
  private readonly pendingEnd = new Map<string, ReturnType<typeof setTimeout>>();

  constructor(
    private readonly onSessionEnd?: (docId: string, ctx: { lastEditor?: UserInfo }) => void,
    private readonly debounceMs: number = Number(process.env.SESSION_END_DEBOUNCE_MS ?? 8000),
  ) {}

  join(docId: string, ws: WebSocket): void {
    // A (re)join means the editing session is still alive — cancel any pending
    // session-end webhook for this document.
    const pending = this.pendingEnd.get(docId);
    if (pending) {
      clearTimeout(pending);
      this.pendingEnd.delete(docId);
    }
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
      // The last editor just left → arm the debounced session-end signal.
      if (room!.size === 0) this.scheduleSessionEnd(docId, m);
    });
  }

  /** Arm (or re-arm) the debounced session-end for a now-empty room. The last
   *  connection's user is the best-effort "last editor"; the dispatcher falls
   *  back to the most recent author if it's absent. */
  private scheduleSessionEnd(docId: string, m: Conn | undefined): void {
    if (!this.onSessionEnd) return;
    const existing = this.pendingEnd.get(docId);
    if (existing) clearTimeout(existing);
    const lastEditor = m?.user;
    const timer = setTimeout(() => {
      this.pendingEnd.delete(docId);
      this.onSessionEnd!(docId, lastEditor ? { lastEditor } : {});
    }, this.debounceMs);
    this.pendingEnd.set(docId, timer);
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

  /** A review op (suggestion/comment) → the rest of the room. Excludes the
   *  sender, who already applied it locally (apply is idempotent regardless). */
  publishReview(docId: string, env: ReviewOpEnvelope, sender: WebSocket): void {
    this.toRoom(docId, { type: "review", review: env }, sender);
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

export function createApp(
  store: ChangeStore,
  opts: { dispatcher?: WebhookDispatcher } = {},
): { server: Server; bcast: Broadcaster } {
  const dispatcher = opts.dispatcher ?? new WebhookDispatcher(store);
  const bcast = new Broadcaster((docId, ctx) => {
    void dispatcher.onSessionEnded(docId, ctx);
  });

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
        let msg: { type?: string; change?: Change; review?: ReviewOpEnvelope; siteId?: string; user?: UserInfo; selection?: unknown };
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
        } else if (msg.type === "review" && msg.review) {
          // Persist the review op (append log) and relay to the rest of the room.
          await store.appendReviewOp(docId, msg.review);
          bcast.publishReview(docId, msg.review, ws);
          // Notify third-party systems about any @-mentions in the op.
          void dispatcher.onReviewOp(docId, msg.review.op);
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
      "access-control-allow-methods": "GET,POST,PUT,PATCH,DELETE,OPTIONS",
      "access-control-allow-headers": "content-type,content-encoding,authorization,x-filename,x-user",
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

  // --- admin API (dashboard) — token-guarded except the login route ----------
  if (parts[0] === "admin") {
    if (method === "POST" && parts[1] === "login" && parts.length === 2) {
      const { username, password } = await readJson<{ username?: string; password?: string }>(req);
      if (!username || !password || !checkCredentials(username, password)) {
        return sendJson(res, 401, { error: "invalid credentials" });
      }
      return sendJson(res, 200, issueToken(username));
    }
    if (!requireAdmin(req, res)) return;

    if (method === "GET" && parts[1] === "docs" && parts.length === 2) {
      return sendJson(res, 200, { documents: await store.listDocuments() });
    }
    // GET /admin/docs/:id/activity — per-document edit history with author names.
    if (method === "GET" && parts[1] === "docs" && parts[3] === "activity" && parts[2]) {
      const activity = await store.getActivity(parts[2]);
      if (!activity) return sendJson(res, 404, { error: "document not found" });
      return sendJson(res, 200, activity);
    }
    if (method === "GET" && parts[1] === "users" && parts.length === 2) {
      return sendJson(res, 200, { users: await store.listUsers() });
    }
    if (parts[1] === "webhooks") {
      if (method === "GET" && parts.length === 2) {
        return sendJson(res, 200, { webhooks: await store.listWebhooks() });
      }
      if (method === "POST" && parts.length === 2) {
        const b = await readJson<{ url?: string; secret?: string; events?: string[] }>(req);
        if (!b.url) return sendJson(res, 400, { error: "url required" });
        const secret = b.secret && b.secret.length > 0 ? b.secret : randomSecret();
        const events = b.events && b.events.length > 0 ? b.events : [SESSION_ENDED_EVENT];
        return sendJson(res, 201, await store.createWebhook({ url: b.url, secret, events }));
      }
      if (parts[2]) {
        const id = parts[2];
        if (method === "DELETE" && parts.length === 3) {
          await store.deleteWebhook(id);
          return sendJson(res, 200, { id });
        }
        if (method === "PATCH" && parts.length === 3) {
          const b = await readJson<{ active?: boolean }>(req);
          if (typeof b.active === "boolean") await store.setWebhookActive(id, b.active);
          return sendJson(res, 200, { id });
        }
        if (method === "GET" && parts[3] === "deliveries") {
          return sendJson(res, 200, { deliveries: await store.listDeliveries(id) });
        }
      }
    }

    // Integration tokens (API keys) for 3rd-party /upload.
    if (parts[1] === "tokens") {
      if (method === "GET" && parts.length === 2) {
        return sendJson(res, 200, { tokens: await store.listApiTokens() });
      }
      if (method === "POST" && parts.length === 2) {
        const b = await readJson<{ name?: string }>(req);
        const name = b.name?.trim() || "integration";
        const { token, tokenHash, prefix } = newApiToken();
        const rec = await store.createApiToken({ name, tokenHash, prefix });
        // The plaintext token is returned ONCE here and never stored.
        return sendJson(res, 201, { ...rec, token });
      }
      if (method === "DELETE" && parts[2] && parts.length === 3) {
        await store.revokeApiToken(parts[2]);
        return sendJson(res, 200, { id: parts[2] });
      }
    }
    return sendJson(res, 404, { error: "not found" });
  }

  // POST /upload — parse a .docx server-side, store it, return its docId so the
  // caller can redirect into WordCanvas({ docId }). Body = raw file bytes;
  // filename + optional user come from headers.
  if (method === "POST" && parts.length === 1 && parts[0] === "upload") {
    if (UPLOAD_REQUIRES_AUTH && !(await authorizeUpload(req, store))) {
      return sendJson(res, 401, { error: "unauthorized" });
    }
    const bytes = await readBody(req);
    const filename = headerStr(req.headers["x-filename"]) ?? "Untitled.docx";
    const userHeader = headerStr(req.headers["x-user"]);
    const user = userHeader ? safeParseUser(userHeader) : undefined;
    try {
      const result = await parseAndStore(new Uint8Array(bytes), filename, store, user);
      return sendJson(res, 201, result);
    } catch (e) {
      return sendJson(res, 400, { error: String(e instanceof Error ? e.message : e) });
    }
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

    // GET /docs/:id/review  — the review overlay (track changes + comments),
    // rehydrated to HEAD: the op log folded with anchors fast-forwarded through
    // the change log, so a joiner sees every suggestion/comment correctly placed.
    if (method === "GET" && parts[2] === "review") {
      const snap = await store.getSnapshot(docId);
      if (!snap) return sendJson(res, 404, { error: "document not found" });
      const [changes, reviewOps] = await Promise.all([
        store.getChanges(docId, snap.version),
        store.getReviewOps(docId),
      ]);
      const layer = rehydrateReview(docId, snap.snapshot, snap.version, changes, reviewOps);
      return sendJson(res, 200, { docId, version: snap.version + changes.length, layer });
    }

    // POST /docs/:id/review/ops  — append a review op (REST fallback for offline
    // clients; the live path is the WebSocket "review" message).
    if (method === "POST" && parts[2] === "review" && parts[3] === "ops") {
      const env = await readJson<ReviewOpEnvelope>(req);
      await store.appendReviewOp(docId, env);
      return sendJson(res, 200, { ok: true });
    }

    // GET /docs/:id/activity  — edit history with author names (attribution)
    if (method === "GET" && parts[2] === "activity") {
      const activity = await store.getActivity(docId);
      if (!activity) return sendJson(res, 404, { error: "document not found" });
      return sendJson(res, 200, activity);
    }

    // GET /docs/:id/export.docx | /docs/:id/export.pdf — server-side download.
    // These are the URLs the session-ended webhook hands to integrations.
    if (method === "GET" && (parts[2] === "export.docx" || parts[2] === "export.pdf")) {
      const format = parts[2] === "export.pdf" ? "pdf" : "docx";
      // ?tracked=accept folds suggestions as accepted; default rejects (baseline).
      const bakeMode = url.searchParams.get("tracked") === "accept" ? "accept" : "reject";
      const exported = await exportDoc(docId, format, store, bakeMode);
      if (!exported) return sendJson(res, 404, { error: "document not found" });
      const filename = `${sanitizeFilename(exported.title ?? docId)}.${format}`;
      const mime =
        format === "pdf"
          ? "application/pdf"
          : "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
      res.writeHead(200, {
        "content-type": mime,
        "content-disposition": `attachment; filename="${filename}"`,
        "access-control-allow-origin": "*",
      });
      res.end(Buffer.from(exported.bytes));
      return;
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
