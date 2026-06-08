// Integration test — requires the Postgres from `docker compose up -d`.
// Skips automatically if the DB is unreachable so `npm test` stays green offline.

import { afterAll, describe, expect, it } from "vitest";
import {
  reconstruct,
  serializeDocument,
  textOfRuns,
  type Change,
  type Document,
  type Op,
  type Paragraph,
  type SectionProps,
} from "@cw/shared";
import { createPool } from "../db";
import { PgChangeStore } from "./ChangeStore";

const pool = createPool();
const store = new PgChangeStore(pool);
const docIds: string[] = [];

// Honest skip: probe connectivity up-front. A reachable DB runs the suite (it is
// NOT optional then); an unreachable one is reported as skipped, never as a
// vacuous pass. Set the env to force-skip even when a DB is up.
let dbUp = false;
if (!process.env.SKIP_DB_TESTS) {
  try {
    await pool.query("SELECT 1");
    dbUp = true;
  } catch {
    await pool.end();
  }
}

afterAll(async () => {
  if (docIds.length) await pool.query("DELETE FROM documents WHERE id = ANY($1)", [docIds]);
  await pool.end();
});

const section = (): SectionProps => ({ pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } });
const CHAR = { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" };
const PARA = { align: "left" as const, lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 };
const doc = (text: string): Document => ({ section: section(), blocks: [{ kind: "paragraph", id: "p1", revision: 0, runs: [{ text, style: CHAR }], style: PARA } as Paragraph] });
const insert = (offset: number, text: string, id: string): Change => ({
  id, docId: "", baseVersion: 0, siteId: "siteA", origin: "typing", ts: 1700000000000,
  ops: [{ type: "insertText", at: { blockId: "p1", offset }, text } as Op],
});
const bodyText = (d: Document): string => textOfRuns((d.blocks[0] as Paragraph).runs);

describe.runIf(dbUp)("PgChangeStore (integration)", () => {
  it("create → append → getChanges → reconstruct round-trips", async () => {
    const base = serializeDocument(doc("hello"));
    const created = await store.createDocument(base);
    docIds.push(created.docId);
    expect(created.version).toBe(0);

    const a0 = await store.appendChange(created.docId, insert(5, " world", "ch0"));
    const a1 = await store.appendChange(created.docId, insert(0, ">> ", "ch1"));
    expect(a0.seq).toBe(0);
    expect(a1.seq).toBe(1);
    expect(await store.getHead(created.docId)).toBe(2);

    const snap = await store.getSnapshot(created.docId);
    const changes = await store.getChanges(created.docId, snap!.version);
    expect(changes).toHaveLength(2);
    expect(changes[0]!.ops).toEqual(a0.ops); // jsonb round-trips ops faithfully
    expect(changes[0]!.origin).toBe("typing");

    expect(bodyText(reconstruct(snap!.snapshot, changes))).toBe(">> hello world");
  });

  it("appendChange is idempotent on change.id", async () => {
    const base = serializeDocument(doc("x"));
    const { docId } = await store.createDocument(base);
    docIds.push(docId);
    const first = await store.appendChange(docId, insert(1, "y", "dup"));
    const again = await store.appendChange(docId, insert(1, "DIFFERENT", "dup")); // same id
    expect(again.seq).toBe(first.seq); // returned the stored one
    expect(await store.getHead(docId)).toBe(1); // not double-applied
    const changes = await store.getChanges(docId, 0);
    expect(changes).toHaveLength(1);
  });

  it("stores and retrieves content-addressed media", async () => {
    const bytes = new Uint8Array([1, 2, 3, 4, 250]);
    await store.putMedia({ hash: "testhash-abc", mime: "image/png", bytes });
    await store.putMedia({ hash: "testhash-abc", mime: "image/png", bytes }); // idempotent
    const got = await store.getMedia("testhash-abc");
    expect(got?.mime).toBe("image/png");
    expect([...(got?.bytes ?? [])]).toEqual([1, 2, 3, 4, 250]);
    await pool.query("DELETE FROM media WHERE hash = $1", ["testhash-abc"]);
  });
});
