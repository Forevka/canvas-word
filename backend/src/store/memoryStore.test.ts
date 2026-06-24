import {
  serializeDocument,
  type Change,
  type Document,
  type Op,
  type Paragraph,
  type SectionProps,
} from "@cw/shared";
import { describe, expect, it } from "vitest";
import { MemoryChangeStore } from "./memoryStore";

const section = (): SectionProps => ({
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
});
const CHAR = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA = { align: "left" as const, lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const doc = (text: string): Document => ({
  section: section(),
  blocks: [{ kind: "paragraph", id: "p1", revision: 0, runs: [{ text, style: CHAR }], style: PARA } as Paragraph],
});
const base = () => serializeDocument(doc("hello"));
const insert = (offset: number, text: string, id: string, extra: Partial<Change> = {}): Change => ({
  id,
  docId: "",
  baseVersion: 0,
  siteId: "siteA",
  origin: "typing",
  ts: 1700000000000,
  ops: [{ type: "insertText", at: { blockId: "p1", offset }, text } as Op],
  ...extra,
});

describe("MemoryChangeStore documents", () => {
  it("creates a document at version 0 with a generated id", async () => {
    const store = new MemoryChangeStore();
    const { docId, version } = await store.createDocument(base());
    expect(version).toBe(0);
    expect(docId).toBeTruthy();
    expect(await store.getHead(docId)).toBe(0);
  });

  it("honors an explicit docId, title and createdBy", async () => {
    const store = new MemoryChangeStore();
    const { docId } = await store.createDocument(base(), { docId: "fixed", title: "T", createdBy: "u1" });
    expect(docId).toBe("fixed");
    expect(await store.getDocumentTitle("fixed")).toBe("T");
  });

  it("returns the stored snapshot at version 0", async () => {
    const store = new MemoryChangeStore();
    const { docId } = await store.createDocument(base());
    const snap = await store.getSnapshot(docId);
    expect(snap?.version).toBe(0);
    expect(snap?.snapshot).toEqual(base());
  });

  it("returns null snapshot / head for an unknown doc", async () => {
    const store = new MemoryChangeStore();
    expect(await store.getSnapshot("nope")).toBeNull();
    expect(await store.getHead("nope")).toBeNull();
    expect(await store.getDocumentTitle("nope")).toBeNull();
  });

  it("updates the title via setDocumentTitle", async () => {
    const store = new MemoryChangeStore();
    const { docId } = await store.createDocument(base());
    await store.setDocumentTitle(docId, "Renamed");
    expect(await store.getDocumentTitle(docId)).toBe("Renamed");
  });
});

describe("MemoryChangeStore changes", () => {
  it("appends changes, advancing head and assigning sequential seq", async () => {
    const store = new MemoryChangeStore();
    const { docId } = await store.createDocument(base());
    const a0 = await store.appendChange(docId, insert(5, " world", "ch0"));
    const a1 = await store.appendChange(docId, insert(0, ">> ", "ch1"));
    expect(a0.seq).toBe(0);
    expect(a1.seq).toBe(1);
    expect(await store.getHead(docId)).toBe(2);
  });

  it("is idempotent on change.id (returns the stored change, does not double-apply)", async () => {
    const store = new MemoryChangeStore();
    const { docId } = await store.createDocument(base());
    const first = await store.appendChange(docId, insert(5, "a", "dup"));
    const again = await store.appendChange(docId, insert(5, "DIFFERENT", "dup"));
    expect(again.seq).toBe(first.seq);
    expect(await store.getHead(docId)).toBe(1);
    expect(await store.getChanges(docId, 0)).toHaveLength(1);
  });

  it("getChanges filters by sinceSeq", async () => {
    const store = new MemoryChangeStore();
    const { docId } = await store.createDocument(base());
    await store.appendChange(docId, insert(5, "a", "c0"));
    await store.appendChange(docId, insert(6, "b", "c1"));
    expect(await store.getChanges(docId, 0)).toHaveLength(2);
    expect(await store.getChanges(docId, 1)).toHaveLength(1);
    expect((await store.getChanges(docId, 1))[0]!.id).toBe("c1");
  });

  it("throws when appending to a missing document", async () => {
    const store = new MemoryChangeStore();
    await expect(store.appendChange("ghost", insert(0, "x", "c0"))).rejects.toThrow(/not found/);
  });
});

describe("MemoryChangeStore fontsConfig", () => {
  it("defaults to null and round-trips a set value", async () => {
    const store = new MemoryChangeStore();
    const { docId } = await store.createDocument(base());
    expect(await store.getFontsConfig(docId)).toBeNull();
    const cfg = { disableBuiltin: ["Arial"], fonts: [] };
    await store.setFontsConfig(docId, cfg);
    expect(await store.getFontsConfig(docId)).toEqual(cfg);
  });

  it("accepts fontsConfig at creation time", async () => {
    const store = new MemoryChangeStore();
    const cfg = { disableBuiltin: [], fonts: [] };
    const { docId } = await store.createDocument(base(), { fontsConfig: cfg });
    expect(await store.getFontsConfig(docId)).toEqual(cfg);
  });
});

describe("MemoryChangeStore media", () => {
  it("stores and retrieves content-addressed media", async () => {
    const store = new MemoryChangeStore();
    const bytes = new Uint8Array([1, 2, 3, 250]);
    await store.putMedia({ hash: "h1", mime: "image/png", bytes });
    const got = await store.getMedia("h1");
    expect(got?.mime).toBe("image/png");
    expect([...(got?.bytes ?? [])]).toEqual([1, 2, 3, 250]);
  });

  it("returns null for unknown media", async () => {
    const store = new MemoryChangeStore();
    expect(await store.getMedia("missing")).toBeNull();
  });

  it("copies the bytes so later mutation of the source does not leak in", async () => {
    const store = new MemoryChangeStore();
    const bytes = new Uint8Array([9, 9]);
    await store.putMedia({ hash: "h2", mime: "image/png", bytes });
    bytes[0] = 0;
    expect([...((await store.getMedia("h2"))?.bytes ?? [])]).toEqual([9, 9]);
  });
});

describe("MemoryChangeStore users + activity", () => {
  it("upserts a user and reflects it in activity", async () => {
    const store = new MemoryChangeStore();
    await store.upsertUser({ id: "u-alice", firstName: "Alice", lastName: "Smith" });
    const { docId } = await store.createDocument(base(), { createdBy: "u-alice" });
    await store.appendChange(docId, insert(5, "!", "act0", { userId: "u-alice" }));

    const activity = await store.getActivity(docId);
    expect(activity?.createdBy).toBe("u-alice");
    expect(activity?.creatorFirstName).toBe("Alice");
    expect(activity?.entries).toHaveLength(1);
    expect(activity?.entries[0]).toMatchObject({
      userId: "u-alice",
      firstName: "Alice",
      lastName: "Smith",
      origin: "typing",
      opCount: 1,
    });
  });

  it("ignores an upsert with an empty id", async () => {
    const store = new MemoryChangeStore();
    await store.upsertUser({ id: "", firstName: "Nobody", lastName: "" });
    expect(await store.listUsers()).toHaveLength(0);
  });

  it("returns null activity for an unknown doc", async () => {
    const store = new MemoryChangeStore();
    expect(await store.getActivity("nope")).toBeNull();
  });

  it("listUsers aggregates docsCreated, changeCount and lastActive", async () => {
    const store = new MemoryChangeStore();
    await store.upsertUser({ id: "u1", firstName: "A", lastName: "B" });
    const { docId } = await store.createDocument(base(), { createdBy: "u1" });
    await store.appendChange(docId, insert(5, "x", "c0", { userId: "u1", ts: 111 }));
    await store.appendChange(docId, insert(6, "y", "c1", { userId: "u1", ts: 222 }));

    const [u] = await store.listUsers();
    expect(u).toMatchObject({ id: "u1", docsCreated: 1, changeCount: 2, lastActive: 222 });
  });

  it("reports lastActive null for a user with no changes", async () => {
    const store = new MemoryChangeStore();
    await store.upsertUser({ id: "u-idle", firstName: "I", lastName: "D" });
    const [u] = await store.listUsers();
    expect(u!.lastActive).toBeNull();
    expect(u!.changeCount).toBe(0);
  });
});

describe("MemoryChangeStore listDocuments", () => {
  it("summarizes documents newest-first with contributor and edit info", async () => {
    const store = new MemoryChangeStore();
    await store.upsertUser({ id: "u1", firstName: "A", lastName: "B" });
    const { docId } = await store.createDocument(base(), { docId: "d1", title: "Doc", createdBy: "u1" });
    await store.appendChange(docId, insert(5, "x", "c0", { userId: "u1", ts: 5000 }));

    const docs = await store.listDocuments();
    expect(docs).toHaveLength(1);
    expect(docs[0]).toMatchObject({
      docId: "d1",
      title: "Doc",
      createdBy: "u1",
      creatorFirstName: "A",
      headVersion: 1,
      contributorCount: 1,
      lastEditedAt: 5000,
    });
  });

  it("reports null lastEditedAt for a doc with no changes", async () => {
    const store = new MemoryChangeStore();
    await store.createDocument(base(), { docId: "empty" });
    const [d] = await store.listDocuments();
    expect(d!.lastEditedAt).toBeNull();
    expect(d!.contributorCount).toBe(0);
  });
});

describe("MemoryChangeStore review ops", () => {
  it("appends and returns review ops in order", async () => {
    const store = new MemoryChangeStore();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const e1 = { op: { type: "addThread" }, dependsOnSeq: 0 } as any;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const e2 = { op: { type: "resolveThread" }, dependsOnSeq: 1 } as any;
    await store.appendReviewOp("d", e1);
    await store.appendReviewOp("d", e2);
    expect(await store.getReviewOps("d")).toEqual([e1, e2]);
  });

  it("returns an empty array for a doc with no review ops", async () => {
    const store = new MemoryChangeStore();
    expect(await store.getReviewOps("none")).toEqual([]);
  });
});

describe("MemoryChangeStore webhooks", () => {
  it("creates, lists, toggles and deletes webhooks", async () => {
    const store = new MemoryChangeStore();
    const w = await store.createWebhook({ url: "https://x/hook", secret: "s", events: ["document.session_ended"] });
    expect(w.active).toBe(true);
    expect(await store.listWebhooks()).toHaveLength(1);

    await store.setWebhookActive(w.id, false);
    expect((await store.listWebhooks())[0]!.active).toBe(false);

    await store.deleteWebhook(w.id);
    expect(await store.listWebhooks()).toHaveLength(0);
  });

  it("records and lists deliveries, filtered by webhook", async () => {
    const store = new MemoryChangeStore();
    await store.recordDelivery({ webhookId: "wh1", docId: "d1", event: "e", status: "success", responseCode: 200, attempts: 1, lastError: null });
    await store.recordDelivery({ webhookId: "wh2", docId: "d1", event: "e", status: "failed", responseCode: 500, attempts: 2, lastError: "boom" });
    expect(await store.listDeliveries()).toHaveLength(2);
    expect(await store.listDeliveries("wh1")).toHaveLength(1);
    expect((await store.listDeliveries("wh1"))[0]!.webhookId).toBe("wh1");
  });
});

describe("MemoryChangeStore api tokens", () => {
  it("creates a token record without exposing the hash, and lists it", async () => {
    const store = new MemoryChangeStore();
    const rec = await store.createApiToken({ name: "ci", tokenHash: "abc123", prefix: "cwk_0001" });
    expect(rec).toMatchObject({ name: "ci", prefix: "cwk_0001", active: true, lastUsedAt: null });
    expect("tokenHash" in rec).toBe(false);
    expect((await store.listApiTokens())[0]!.id).toBe(rec.id);
  });

  it("verifies an active token by its hash and records lastUsedAt", async () => {
    const store = new MemoryChangeStore();
    await store.createApiToken({ name: "ci", tokenHash: "hash-1", prefix: "cwk_x" });
    expect(await store.verifyApiToken("hash-1")).toBe(true);
    expect(await store.verifyApiToken("wrong-hash")).toBe(false);
    expect((await store.listApiTokens())[0]!.lastUsedAt).not.toBeNull();
  });

  it("does not verify a revoked token", async () => {
    const store = new MemoryChangeStore();
    const rec = await store.createApiToken({ name: "ci", tokenHash: "hash-2", prefix: "cwk_y" });
    await store.revokeApiToken(rec.id);
    expect(await store.verifyApiToken("hash-2")).toBe(false);
    expect(await store.listApiTokens()).toHaveLength(0);
  });
});
