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
import { SNAPSHOT_VERSION, transformOps, type Change, type Op, type SerializedDocument, type UserInfo } from "@cw/shared";

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

export interface ChangeStore {
  /** Create a document from a base snapshot (stored as version 0). */
  createDocument(
    base: SerializedDocument,
    opts?: { docId?: string; createdBy?: string },
  ): Promise<{ docId: string; version: number }>;
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
}

const CHANGE_COLS = "seq, change_id, base_version, site_id, user_id, origin, ts, ops, selection";

export class PgChangeStore implements ChangeStore {
  constructor(private readonly pool: Pool) {}

  async createDocument(
    base: SerializedDocument,
    opts: { docId?: string; createdBy?: string } = {},
  ): Promise<{ docId: string; version: number }> {
    const { docId, createdBy = null } = opts;
    const client = await this.pool.connect();
    try {
      await client.query("BEGIN");
      const row = docId
        ? await client.query<{ id: string }>(
            "INSERT INTO documents (id, created_by) VALUES ($1, $2) RETURNING id",
            [docId, createdBy],
          )
        : await client.query<{ id: string }>(
            "INSERT INTO documents (created_by) VALUES ($1) RETURNING id",
            [createdBy],
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
