// In-memory ChangeStore — a dependency-free implementation of the full store
// interface. Used by tests (no Postgres required) and usable as a zero-config
// dev backend. Mirrors PgChangeStore's semantics, including OT rebase on append.

import { randomUUID } from "node:crypto";
import { transformOps, type Change, type SerializedDocument, type UserInfo } from "@cw/shared";
import type {
  ApiTokenRecord,
  ChangeStore,
  DeliveryRecord,
  DocSnapshotRecord,
  DocumentActivity,
  DocumentSummary,
  MediaRecord,
  UserSummary,
  WebhookRecord,
} from "./ChangeStore";

interface DocRow {
  id: string;
  title: string | null;
  createdBy: string | null;
  createdAt: number;
  head: number;
  snapshot: SerializedDocument;
}

export class MemoryChangeStore implements ChangeStore {
  private readonly docs = new Map<string, DocRow>();
  private readonly log = new Map<string, Change[]>();
  private readonly users = new Map<string, UserInfo & { updatedAt: number }>();
  private readonly media = new Map<string, MediaRecord>();
  private readonly webhooks = new Map<string, WebhookRecord>();
  private readonly deliveries: DeliveryRecord[] = [];
  private readonly apiTokens = new Map<string, ApiTokenRecord & { tokenHash: string }>();

  async createDocument(
    base: SerializedDocument,
    opts: { docId?: string; createdBy?: string; title?: string } = {},
  ): Promise<{ docId: string; version: number }> {
    const docId = opts.docId ?? randomUUID();
    this.docs.set(docId, {
      id: docId,
      title: opts.title ?? null,
      createdBy: opts.createdBy ?? null,
      createdAt: Date.now(),
      head: 0,
      snapshot: base,
    });
    this.log.set(docId, []);
    return { docId, version: 0 };
  }

  async setDocumentTitle(docId: string, title: string): Promise<void> {
    const d = this.docs.get(docId);
    if (d) d.title = title;
  }

  async getDocumentTitle(docId: string): Promise<string | null> {
    return this.docs.get(docId)?.title ?? null;
  }

  async getSnapshot(docId: string): Promise<DocSnapshotRecord | null> {
    const d = this.docs.get(docId);
    return d ? { docId, version: 0, snapshot: d.snapshot } : null;
  }

  async getChanges(docId: string, sinceSeq: number): Promise<Change[]> {
    return (this.log.get(docId) ?? []).filter((c) => (c.seq ?? 0) >= sinceSeq);
  }

  async appendChange(docId: string, change: Change): Promise<Change> {
    const d = this.docs.get(docId);
    if (!d) throw new Error(`document ${docId} not found`);
    const log = this.log.get(docId)!;
    const existing = log.find((c) => c.id === change.id);
    if (existing) return existing;
    const seq = d.head;
    let ops = change.ops;
    if (change.baseVersion < seq) {
      const against = log
        .filter((c) => (c.seq ?? 0) >= change.baseVersion && (c.seq ?? 0) < seq)
        .flatMap((c) => c.ops);
      ops = transformOps(change.ops, against, "right");
    }
    const stored: Change = { ...change, ops, seq };
    log.push(stored);
    d.head = seq + 1;
    return stored;
  }

  async getHead(docId: string): Promise<number | null> {
    return this.docs.has(docId) ? this.docs.get(docId)!.head : null;
  }

  async upsertUser(user: UserInfo): Promise<void> {
    if (user.id) this.users.set(user.id, { ...user, updatedAt: Date.now() });
  }

  async getActivity(docId: string): Promise<DocumentActivity | null> {
    const d = this.docs.get(docId);
    if (!d) return null;
    const creator = d.createdBy ? this.users.get(d.createdBy) : undefined;
    const entries = [...(this.log.get(docId) ?? [])]
      .sort((a, b) => (b.seq ?? 0) - (a.seq ?? 0))
      .map((c) => {
        const u = c.userId ? this.users.get(c.userId) : undefined;
        return {
          seq: c.seq ?? 0,
          userId: c.userId ?? null,
          firstName: u?.firstName ?? "",
          lastName: u?.lastName ?? "",
          ts: c.ts,
          origin: c.origin,
          opCount: c.ops.length,
        };
      });
    return {
      docId,
      createdBy: d.createdBy,
      creatorFirstName: creator?.firstName ?? "",
      creatorLastName: creator?.lastName ?? "",
      createdAt: d.createdAt,
      entries,
    };
  }

  async putMedia(rec: MediaRecord): Promise<void> {
    this.media.set(rec.hash, { ...rec, bytes: new Uint8Array(rec.bytes) });
  }

  async getMedia(hash: string): Promise<MediaRecord | null> {
    return this.media.get(hash) ?? null;
  }

  async listDocuments(): Promise<DocumentSummary[]> {
    return [...this.docs.values()]
      .sort((a, b) => b.createdAt - a.createdAt)
      .map((d) => {
        const log = this.log.get(d.id) ?? [];
        const creator = d.createdBy ? this.users.get(d.createdBy) : undefined;
        const contributors = new Set(log.map((c) => c.userId).filter((x): x is string => !!x));
        const lastEditedAt = log.length ? Math.max(...log.map((c) => c.ts)) : null;
        return {
          docId: d.id,
          title: d.title,
          createdBy: d.createdBy,
          creatorFirstName: creator?.firstName ?? "",
          creatorLastName: creator?.lastName ?? "",
          createdAt: d.createdAt,
          headVersion: d.head,
          contributorCount: contributors.size,
          lastEditedAt,
        };
      });
  }

  async listUsers(): Promise<UserSummary[]> {
    return [...this.users.values()].map((u) => {
      let docsCreated = 0;
      for (const d of this.docs.values()) if (d.createdBy === u.id) docsCreated++;
      let changeCount = 0;
      let lastActive = 0;
      for (const log of this.log.values()) {
        for (const c of log) {
          if (c.userId === u.id) {
            changeCount++;
            lastActive = Math.max(lastActive, c.ts);
          }
        }
      }
      return {
        id: u.id,
        firstName: u.firstName,
        lastName: u.lastName,
        docsCreated,
        changeCount,
        lastActive: changeCount ? lastActive : null,
      };
    });
  }

  async listWebhooks(): Promise<WebhookRecord[]> {
    return [...this.webhooks.values()].sort((a, b) => b.createdAt - a.createdAt);
  }

  async createWebhook(input: { url: string; secret: string; events: string[] }): Promise<WebhookRecord> {
    const w: WebhookRecord = {
      id: randomUUID(),
      url: input.url,
      secret: input.secret,
      events: input.events,
      active: true,
      createdAt: Date.now(),
    };
    this.webhooks.set(w.id, w);
    return w;
  }

  async deleteWebhook(id: string): Promise<void> {
    this.webhooks.delete(id);
  }

  async setWebhookActive(id: string, active: boolean): Promise<void> {
    const w = this.webhooks.get(id);
    if (w) w.active = active;
  }

  async recordDelivery(rec: Omit<DeliveryRecord, "id" | "createdAt">): Promise<void> {
    this.deliveries.push({ ...rec, id: randomUUID(), createdAt: Date.now() });
  }

  async listDeliveries(webhookId?: string): Promise<DeliveryRecord[]> {
    return this.deliveries
      .filter((d) => !webhookId || d.webhookId === webhookId)
      .sort((a, b) => b.createdAt - a.createdAt);
  }

  async createApiToken(input: { name: string; tokenHash: string; prefix: string }): Promise<ApiTokenRecord> {
    const rec: ApiTokenRecord = {
      id: randomUUID(),
      name: input.name,
      prefix: input.prefix,
      active: true,
      createdAt: Date.now(),
      lastUsedAt: null,
    };
    this.apiTokens.set(rec.id, { ...rec, tokenHash: input.tokenHash });
    return rec;
  }

  async listApiTokens(): Promise<ApiTokenRecord[]> {
    return [...this.apiTokens.values()]
      .map((t) => ({
        id: t.id,
        name: t.name,
        prefix: t.prefix,
        active: t.active,
        createdAt: t.createdAt,
        lastUsedAt: t.lastUsedAt,
      }))
      .sort((a, b) => b.createdAt - a.createdAt);
  }

  async revokeApiToken(id: string): Promise<void> {
    this.apiTokens.delete(id);
  }

  async verifyApiToken(tokenHash: string): Promise<boolean> {
    for (const t of this.apiTokens.values()) {
      if (t.active && t.tokenHash === tokenHash) {
        t.lastUsedAt = Date.now();
        return true;
      }
    }
    return false;
  }
}
