// ChangeStore — the persistence layer for the collaboration backend.
//
// A document is stored as a base snapshot (version 0) plus an append-only,
// totally-ordered change log; the server assigns each accepted change its
// canonical seq. Reconstructing a version = newest snapshot + replay of the
// changes after it (see shared/replay). Media is content-addressed (sha256).
//
// The interface is narrow so the Postgres impl below can later be swapped for a
// different store without touching the server.

import type { Pool } from "pg";
import { SNAPSHOT_VERSION, transformOps, type Change, type Op, type ReviewOpEnvelope, type SerializedDocument, type UserInfo } from "@cw/shared";

export interface DocSnapshotRecord {
  docId: string;
  /** Document version this snapshot represents (0 = base, N = after N changes). */
  version: number;
  snapshot: SerializedDocument;
}

export interface MediaRecord {
  hash: string;
  mime: string;
  bytes: Uint8Array;
}

/** One row of a document's edit history (for the activity panel). */
export interface ActivityEntry {
  seq: number;
  userId: string | null;
  firstName: string;
  lastName: string;
  ts: number;
  origin: Change["origin"];
  opCount: number;
}

export interface DocumentActivity {
  docId: string;
  createdBy: string | null;
  creatorFirstName: string;
  creatorLastName: string;
  createdAt: number;
  entries: ActivityEntry[];
}

/** One row of the dashboard's document list. */
export interface DocumentSummary {
  docId: string;
  title: string | null;
  createdBy: string | null;
  creatorFirstName: string;
  creatorLastName: string;
  createdAt: number;
  headVersion: number;
  /** Distinct authors who have a recorded change (NULL authors excluded). */
  contributorCount: number;
  /** ts of the most recent change, or null if none. */
  lastEditedAt: number | null;
}

/** One row of the dashboard's user list. */
export interface UserSummary {
  id: string;
  firstName: string;
  lastName: string;
  docsCreated: number;
  changeCount: number;
  lastActive: number | null;
}

export interface WebhookRecord {
  id: string;
  url: string;
  secret: string;
  events: string[];
  active: boolean;
  createdAt: number;
}

export interface DeliveryRecord {
  id: string;
  webhookId: string;
  docId: string | null;
  event: string;
  status: "success" | "failed";
  responseCode: number | null;
  attempts: number;
  lastError: string | null;
  createdAt: number;
}

/** A 3rd-party integration token (API key). The hash is never returned. */
export interface ApiTokenRecord {
  id: string;
  name: string;
  prefix: string;
  active: boolean;
  createdAt: number;
  lastUsedAt: number | null;
}

export interface ChangeStore {
  /** Create a document from a base snapshot (stored as version 0). */
  createDocument(
    base: SerializedDocument,
    opts?: { docId?: string; createdBy?: string; title?: string },
  ): Promise<{ docId: string; version: number }>;
  /** Rename a document (dashboard / upload filename). */
  setDocumentTitle(docId: string, title: string): Promise<void>;
  /** A document's title (for export filenames), or null if unset/missing. */
  getDocumentTitle(docId: string): Promise<string | null>;
  /** Document list for the dashboard, newest first. */
  listDocuments(): Promise<DocumentSummary[]>;
  /** User list for the dashboard, most-recently-seen first. */
  listUsers(): Promise<UserSummary[]>;
  /** Registered webhook endpoints. */
  listWebhooks(): Promise<WebhookRecord[]>;
  createWebhook(input: { url: string; secret: string; events: string[] }): Promise<WebhookRecord>;
  deleteWebhook(id: string): Promise<void>;
  setWebhookActive(id: string, active: boolean): Promise<void>;
  /** Append a webhook delivery attempt to the log. */
  recordDelivery(rec: Omit<DeliveryRecord, "id" | "createdAt">): Promise<void>;
  /** Delivery log, newest first, optionally scoped to one webhook. */
  listDeliveries(webhookId?: string): Promise<DeliveryRecord[]>;
  /** Integration tokens (API keys) for machine-to-machine upload. */
  createApiToken(input: { name: string; tokenHash: string; prefix: string }): Promise<ApiTokenRecord>;
  listApiTokens(): Promise<ApiTokenRecord[]>;
  revokeApiToken(id: string): Promise<void>;
  /** True if the hash matches an active token; touches its last_used_at. */
  verifyApiToken(tokenHash: string): Promise<boolean>;
  /** Newest snapshot (base or checkpoint) + the version it represents. */
  getSnapshot(docId: string): Promise<DocSnapshotRecord | null>;
  /** Changes with seq >= sinceSeq, in seq order. */
  getChanges(docId: string, sinceSeq: number): Promise<Change[]>;
  /** Append a change: assign the next canonical seq, bump head_version, return
   *  the change with its seq. Idempotent on change.id (re-submits are no-ops). */
  appendChange(docId: string, change: Change): Promise<Change>;
  /** Latest version (number of accepted changes), or null if no such document. */
  getHead(docId: string): Promise<number | null>;
  /** Upsert a caller-supplied user (id + names) for attribution. */
  upsertUser(user: UserInfo): Promise<void>;
  /** Full edit history with author names, or null if the document is missing. */
  getActivity(docId: string): Promise<DocumentActivity | null>;
  putMedia(rec: MediaRecord): Promise<void>;
  getMedia(hash: string): Promise<MediaRecord | null>;
  // ---- review layer (track changes + comments) ----------------------------
  /** Append a review op to the document's review log (separate from changes). */
  appendReviewOp(docId: string, env: ReviewOpEnvelope): Promise<void>;
  /** The review op log in append order — fold through applyReviewOp to rebuild
   *  the overlay (a late joiner then fast-forwards anchors to head). */
  getReviewOps(docId: string): Promise<ReviewOpEnvelope[]>;
}

const CHANGE_COLS = "seq, change_id, base_version, site_id, user_id, origin, ts, ops, selection";

export class PgChangeStore implements ChangeStore {
  constructor(private readonly pool: Pool) {}

  async createDocument(
    base: SerializedDocument,
    opts: { docId?: string; createdBy?: string; title?: string } = {},
  ): Promise<{ docId: string; version: number }> {
    const { docId, createdBy = null, title = null } = opts;
    const client = await this.pool.connect();
    try {
      await client.query("BEGIN");
      const row = docId
        ? await client.query<{ id: string }>(
            "INSERT INTO documents (id, created_by, title) VALUES ($1, $2, $3) RETURNING id",
            [docId, createdBy, title],
          )
        : await client.query<{ id: string }>(
            "INSERT INTO documents (created_by, title) VALUES ($1, $2) RETURNING id",
            [createdBy, title],
          );
      const id = row.rows[0]!.id;
      await client.query("INSERT INTO snapshots (doc_id, version, doc) VALUES ($1, 0, $2)", [
        id,
        JSON.stringify(base.doc),
      ]);
      await client.query("COMMIT");
      return { docId: id, version: 0 };
    } catch (e) {
      await client.query("ROLLBACK");
      throw e;
    } finally {
      client.release();
    }
  }

  async getSnapshot(docId: string): Promise<DocSnapshotRecord | null> {
    const res = await this.pool.query<{ version: string; doc: unknown }>(
      "SELECT version, doc FROM snapshots WHERE doc_id = $1 ORDER BY version DESC LIMIT 1",
      [docId],
    );
    if ((res.rowCount ?? 0) === 0) return null;
    const r = res.rows[0]!;
    return {
      docId,
      version: Number(r.version),
      snapshot: { version: SNAPSHOT_VERSION, doc: r.doc as SerializedDocument["doc"] },
    };
  }

  async getChanges(docId: string, sinceSeq: number): Promise<Change[]> {
    const res = await this.pool.query(
      `SELECT ${CHANGE_COLS} FROM changes WHERE doc_id = $1 AND seq >= $2 ORDER BY seq ASC`,
      [docId, sinceSeq],
    );
    return res.rows.map((r) => this.rowToChange(docId, r));
  }

  async appendChange(docId: string, change: Change): Promise<Change> {
    const client = await this.pool.connect();
    try {
      await client.query("BEGIN");
      // Idempotency: a re-submitted change (same id) returns the stored one.
      const existing = await client.query(
        `SELECT ${CHANGE_COLS} FROM changes WHERE doc_id = $1 AND change_id = $2`,
        [docId, change.id],
      );
      if ((existing.rowCount ?? 0) > 0) {
        await client.query("COMMIT");
        return this.rowToChange(docId, existing.rows[0]);
      }
      // Lock the document row to serialize seq assignment.
      const doc = await client.query<{ head_version: string }>(
        "SELECT head_version FROM documents WHERE id = $1 FOR UPDATE",
        [docId],
      );
      if ((doc.rowCount ?? 0) === 0) throw new Error(`document ${docId} not found`);
      const seq = Number(doc.rows[0]!.head_version);

      // OT: if the change was generated against an older version, rebase its ops
      // onto current head by transforming against the changes committed since —
      // same transform + side the client used, so the two agree (TP1).
      let ops: Op[] = change.ops;
      if (change.baseVersion < seq) {
        const concurrent = await client.query<{ ops: Op[] }>(
          "SELECT ops FROM changes WHERE doc_id = $1 AND seq >= $2 AND seq < $3 ORDER BY seq ASC",
          [docId, change.baseVersion, seq],
        );
        const against = concurrent.rows.flatMap((r) => r.ops);
        ops = transformOps(change.ops, against, "right");
      }

      await client.query(
        `INSERT INTO changes (doc_id, ${CHANGE_COLS})
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10)`,
        [
          docId,
          seq,
          change.id,
          change.baseVersion,
          change.siteId,
          change.userId ?? null,
          change.origin,
          change.ts,
          JSON.stringify(ops),
          change.selectionAfter != null ? JSON.stringify(change.selectionAfter) : null,
        ],
      );
      await client.query("UPDATE documents SET head_version = $1 WHERE id = $2", [seq + 1, docId]);
      await client.query("COMMIT");
      return { ...change, ops, seq };
    } catch (e) {
      await client.query("ROLLBACK");
      throw e;
    } finally {
      client.release();
    }
  }

  async getHead(docId: string): Promise<number | null> {
    const res = await this.pool.query<{ head_version: string }>(
      "SELECT head_version FROM documents WHERE id = $1",
      [docId],
    );
    return (res.rowCount ?? 0) > 0 ? Number(res.rows[0]!.head_version) : null;
  }

  async putMedia(rec: MediaRecord): Promise<void> {
    await this.pool.query(
      "INSERT INTO media (hash, mime, bytes) VALUES ($1, $2, $3) ON CONFLICT (hash) DO NOTHING",
      [rec.hash, rec.mime, Buffer.from(rec.bytes)],
    );
  }

  async getMedia(hash: string): Promise<MediaRecord | null> {
    const res = await this.pool.query<{ mime: string; bytes: Buffer }>(
      "SELECT mime, bytes FROM media WHERE hash = $1",
      [hash],
    );
    if ((res.rowCount ?? 0) === 0) return null;
    const r = res.rows[0]!;
    return { hash, mime: r.mime, bytes: new Uint8Array(r.bytes) };
  }

  async appendReviewOp(docId: string, env: ReviewOpEnvelope): Promise<void> {
    await this.pool.query(
      "INSERT INTO review_ops (doc_id, op, depends_on_seq, author_id) VALUES ($1, $2, $3, $4)",
      [docId, JSON.stringify(env.op), env.dependsOnSeq, env.author?.id ?? null],
    );
  }

  async getReviewOps(docId: string): Promise<ReviewOpEnvelope[]> {
    const res = await this.pool.query<{ op: unknown; depends_on_seq: string; author_id: string | null }>(
      "SELECT op, depends_on_seq, author_id FROM review_ops WHERE doc_id = $1 ORDER BY id ASC",
      [docId],
    );
    return res.rows.map((r) => ({
      op: r.op as ReviewOpEnvelope["op"],
      dependsOnSeq: Number(r.depends_on_seq),
      ...(r.author_id ? { author: { id: r.author_id, firstName: "", lastName: "" } } : {}),
    }));
  }

  async upsertUser(user: UserInfo): Promise<void> {
    if (!user.id) return;
    await this.pool.query(
      `INSERT INTO users (id, first_name, last_name, updated_at) VALUES ($1, $2, $3, $4)
       ON CONFLICT (id) DO UPDATE SET first_name = EXCLUDED.first_name,
         last_name = EXCLUDED.last_name, updated_at = EXCLUDED.updated_at`,
      [user.id, user.firstName ?? "", user.lastName ?? "", Date.now()],
    );
  }

  async getActivity(docId: string): Promise<DocumentActivity | null> {
    const docRes = await this.pool.query<{ created_by: string | null; created_at: string }>(
      `SELECT d.created_by, (extract(epoch from d.created_at) * 1000)::bigint AS created_at
       FROM documents d WHERE d.id = $1`,
      [docId],
    );
    if ((docRes.rowCount ?? 0) === 0) return null;
    const d = docRes.rows[0]!;

    const creator = d.created_by
      ? await this.pool.query<{ first_name: string; last_name: string }>(
          "SELECT first_name, last_name FROM users WHERE id = $1",
          [d.created_by],
        )
      : null;

    const rows = await this.pool.query<{
      seq: string;
      user_id: string | null;
      first_name: string | null;
      last_name: string | null;
      ts: string;
      origin: string;
      op_count: string;
    }>(
      `SELECT c.seq, c.user_id, u.first_name, u.last_name, c.ts, c.origin,
              jsonb_array_length(c.ops) AS op_count
       FROM changes c LEFT JOIN users u ON u.id = c.user_id
       WHERE c.doc_id = $1 ORDER BY c.seq DESC`,
      [docId],
    );

    return {
      docId,
      createdBy: d.created_by,
      creatorFirstName: creator?.rows[0]?.first_name ?? "",
      creatorLastName: creator?.rows[0]?.last_name ?? "",
      createdAt: Number(d.created_at),
      entries: rows.rows.map((r) => ({
        seq: Number(r.seq),
        userId: r.user_id,
        firstName: r.first_name ?? "",
        lastName: r.last_name ?? "",
        ts: Number(r.ts),
        origin: r.origin as Change["origin"],
        opCount: Number(r.op_count),
      })),
    };
  }

  async setDocumentTitle(docId: string, title: string): Promise<void> {
    await this.pool.query("UPDATE documents SET title = $1 WHERE id = $2", [title, docId]);
  }

  async getDocumentTitle(docId: string): Promise<string | null> {
    const res = await this.pool.query<{ title: string | null }>(
      "SELECT title FROM documents WHERE id = $1",
      [docId],
    );
    return (res.rowCount ?? 0) > 0 ? (res.rows[0]!.title ?? null) : null;
  }

  async listDocuments(): Promise<DocumentSummary[]> {
    const res = await this.pool.query<{
      id: string;
      title: string | null;
      created_by: string | null;
      head_version: string;
      created_at: string;
      creator_first: string | null;
      creator_last: string | null;
      contributor_count: string;
      last_edited_at: string | null;
    }>(
      `SELECT d.id, d.title, d.created_by, d.head_version,
              (extract(epoch from d.created_at) * 1000)::bigint AS created_at,
              cu.first_name AS creator_first, cu.last_name AS creator_last,
              (SELECT count(DISTINCT c.user_id) FROM changes c WHERE c.doc_id = d.id) AS contributor_count,
              (SELECT max(c.ts) FROM changes c WHERE c.doc_id = d.id) AS last_edited_at
       FROM documents d
       LEFT JOIN users cu ON cu.id = d.created_by
       ORDER BY d.created_at DESC`,
    );
    return res.rows.map((r) => ({
      docId: r.id,
      title: r.title,
      createdBy: r.created_by,
      creatorFirstName: r.creator_first ?? "",
      creatorLastName: r.creator_last ?? "",
      createdAt: Number(r.created_at),
      headVersion: Number(r.head_version),
      contributorCount: Number(r.contributor_count),
      lastEditedAt: r.last_edited_at != null ? Number(r.last_edited_at) : null,
    }));
  }

  async listUsers(): Promise<UserSummary[]> {
    const res = await this.pool.query<{
      id: string;
      first_name: string;
      last_name: string;
      docs_created: string;
      change_count: string;
      last_active: string | null;
    }>(
      `SELECT u.id, u.first_name, u.last_name,
              (SELECT count(*) FROM documents d WHERE d.created_by = u.id) AS docs_created,
              (SELECT count(*) FROM changes c WHERE c.user_id = u.id) AS change_count,
              (SELECT max(c.ts) FROM changes c WHERE c.user_id = u.id) AS last_active
       FROM users u
       ORDER BY u.updated_at DESC`,
    );
    return res.rows.map((r) => ({
      id: r.id,
      firstName: r.first_name,
      lastName: r.last_name,
      docsCreated: Number(r.docs_created),
      changeCount: Number(r.change_count),
      lastActive: r.last_active != null ? Number(r.last_active) : null,
    }));
  }

  async listWebhooks(): Promise<WebhookRecord[]> {
    const res = await this.pool.query<WebhookRow>(
      `SELECT id, url, secret, events, active,
              (extract(epoch from created_at) * 1000)::bigint AS created_at
       FROM webhooks ORDER BY created_at DESC`,
    );
    return res.rows.map(rowToWebhook);
  }

  async createWebhook(input: { url: string; secret: string; events: string[] }): Promise<WebhookRecord> {
    const res = await this.pool.query<WebhookRow>(
      `INSERT INTO webhooks (url, secret, events) VALUES ($1, $2, $3)
       RETURNING id, url, secret, events, active,
                 (extract(epoch from created_at) * 1000)::bigint AS created_at`,
      [input.url, input.secret, input.events],
    );
    return rowToWebhook(res.rows[0]!);
  }

  async deleteWebhook(id: string): Promise<void> {
    await this.pool.query("DELETE FROM webhooks WHERE id = $1", [id]);
  }

  async setWebhookActive(id: string, active: boolean): Promise<void> {
    await this.pool.query("UPDATE webhooks SET active = $1 WHERE id = $2", [active, id]);
  }

  async recordDelivery(rec: Omit<DeliveryRecord, "id" | "createdAt">): Promise<void> {
    await this.pool.query(
      `INSERT INTO webhook_deliveries
         (webhook_id, doc_id, event, status, response_code, attempts, last_error)
       VALUES ($1, $2, $3, $4, $5, $6, $7)`,
      [rec.webhookId, rec.docId, rec.event, rec.status, rec.responseCode, rec.attempts, rec.lastError],
    );
  }

  async listDeliveries(webhookId?: string): Promise<DeliveryRecord[]> {
    const res = webhookId
      ? await this.pool.query<DeliveryRow>(
          `SELECT id, webhook_id, doc_id, event, status, response_code, attempts, last_error,
                  (extract(epoch from created_at) * 1000)::bigint AS created_at
           FROM webhook_deliveries WHERE webhook_id = $1 ORDER BY created_at DESC LIMIT 200`,
          [webhookId],
        )
      : await this.pool.query<DeliveryRow>(
          `SELECT id, webhook_id, doc_id, event, status, response_code, attempts, last_error,
                  (extract(epoch from created_at) * 1000)::bigint AS created_at
           FROM webhook_deliveries ORDER BY created_at DESC LIMIT 200`,
        );
    return res.rows.map(rowToDelivery);
  }

  async createApiToken(input: { name: string; tokenHash: string; prefix: string }): Promise<ApiTokenRecord> {
    const res = await this.pool.query<ApiTokenRow>(
      `INSERT INTO api_tokens (name, token_hash, prefix) VALUES ($1, $2, $3)
       RETURNING id, name, prefix, active,
                 (extract(epoch from created_at) * 1000)::bigint AS created_at,
                 (extract(epoch from last_used_at) * 1000)::bigint AS last_used_at`,
      [input.name, input.tokenHash, input.prefix],
    );
    return rowToApiToken(res.rows[0]!);
  }

  async listApiTokens(): Promise<ApiTokenRecord[]> {
    const res = await this.pool.query<ApiTokenRow>(
      `SELECT id, name, prefix, active,
              (extract(epoch from created_at) * 1000)::bigint AS created_at,
              (extract(epoch from last_used_at) * 1000)::bigint AS last_used_at
       FROM api_tokens ORDER BY created_at DESC`,
    );
    return res.rows.map(rowToApiToken);
  }

  async revokeApiToken(id: string): Promise<void> {
    await this.pool.query("DELETE FROM api_tokens WHERE id = $1", [id]);
  }

  async verifyApiToken(tokenHash: string): Promise<boolean> {
    const res = await this.pool.query(
      "UPDATE api_tokens SET last_used_at = now() WHERE token_hash = $1 AND active = true RETURNING id",
      [tokenHash],
    );
    return (res.rowCount ?? 0) > 0;
  }

  private rowToChange(docId: string, r: Record<string, unknown>): Change {
    return {
      id: r.change_id as string,
      docId,
      baseVersion: Number(r.base_version),
      seq: Number(r.seq),
      siteId: r.site_id as string,
      ...(r.user_id ? { userId: r.user_id as string } : {}),
      origin: r.origin as Change["origin"],
      ts: Number(r.ts),
      ops: r.ops as Change["ops"],
      selectionAfter: (r.selection as Change["selectionAfter"]) ?? null,
    };
  }
}

// --- webhook / delivery row mappers (module-level: no instance state) --------

interface WebhookRow {
  id: string;
  url: string;
  secret: string;
  events: string[];
  active: boolean;
  created_at: string;
}

interface DeliveryRow {
  id: string;
  webhook_id: string;
  doc_id: string | null;
  event: string;
  status: string;
  response_code: number | null;
  attempts: number;
  last_error: string | null;
  created_at: string;
}

function rowToWebhook(r: WebhookRow): WebhookRecord {
  return {
    id: r.id,
    url: r.url,
    secret: r.secret,
    events: r.events,
    active: r.active,
    createdAt: Number(r.created_at),
  };
}

function rowToDelivery(r: DeliveryRow): DeliveryRecord {
  return {
    id: r.id,
    webhookId: r.webhook_id,
    docId: r.doc_id,
    event: r.event,
    status: r.status === "success" ? "success" : "failed",
    responseCode: r.response_code,
    attempts: r.attempts,
    lastError: r.last_error,
    createdAt: Number(r.created_at),
  };
}

interface ApiTokenRow {
  id: string;
  name: string;
  prefix: string;
  active: boolean;
  created_at: string;
  last_used_at: string | null;
}

function rowToApiToken(r: ApiTokenRow): ApiTokenRecord {
  return {
    id: r.id,
    name: r.name,
    prefix: r.prefix,
    active: r.active,
    createdAt: Number(r.created_at),
    lastUsedAt: r.last_used_at != null ? Number(r.last_used_at) : null,
  };
}
