// /upload authorization via a 3rd-party integration token (no DB): a valid
// X-API-Key is accepted and returns a docId; missing/invalid keys are rejected.

import { runExport } from "@forevka/wordcanvas/export";
import { installMeasureHost } from "@forevka/wordcanvas/export/measure";
import { type Document, type Paragraph, type SectionProps } from "@cw/shared";
import { type Server } from "node:http";
import { type AddressInfo } from "node:net";
import { afterEach, beforeAll, describe, expect, it } from "vitest";
import { newApiToken } from "./auth/apiToken";
import { MemoryChangeStore } from "./store/memoryStore";

let createApp: typeof import("./server").createApp;
beforeAll(async () => {
  ({ createApp } = await import("./server"));
  await installMeasureHost();
});

const section = (): SectionProps => ({
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
});
const doc = (): Document => ({
  section: section(),
  blocks: [
    {
      kind: "paragraph",
      id: "p1",
      revision: 0,
      runs: [{ text: "uploaded via api token", style: { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" } }],
      style: { align: "left", lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 },
    } as Paragraph,
  ],
});

const listen = (server: Server): Promise<number> =>
  new Promise((resolve) => server.listen(0, () => resolve((server.address() as AddressInfo).port)));

const servers: Server[] = [];
afterEach(() => {
  for (const s of servers.splice(0)) s.close();
});

describe("/upload integration-token auth", () => {
  it("accepts a valid token, rejects missing/invalid ones", async () => {
    const store = new MemoryChangeStore();
    const { token, tokenHash, prefix } = newApiToken();
    await store.createApiToken({ name: "partner-crm", tokenHash, prefix });

    const { server } = createApp(store);
    servers.push(server);
    const port = await listen(server);
    const url = `http://127.0.0.1:${port}/upload`;
    const exported = (await runExport(doc(), "docx", {})).bytes;
    // Copy into a plain ArrayBuffer-backed array so it's a valid BodyInit.
    const docx = new Uint8Array(exported.byteLength);
    docx.set(exported);
    const body = (): Blob => new Blob([docx]); // fresh BodyInit per request

    // No credentials → 401.
    const noauth = await fetch(url, { method: "POST", headers: { "x-filename": "r.docx" }, body: body() });
    expect(noauth.status).toBe(401);

    // Wrong token → 401.
    const bad = await fetch(url, {
      method: "POST",
      headers: { "x-api-key": "cwk_not_a_real_token", "x-filename": "r.docx" },
      body: body(),
    });
    expect(bad.status).toBe(401);

    // Valid token (X-API-Key) → 201 + docId, and the doc is stored.
    const ok = await fetch(url, {
      method: "POST",
      headers: { "x-api-key": token, "x-filename": "report.docx" },
      body: body(),
    });
    expect(ok.status).toBe(201);
    const result = (await ok.json()) as { docId: string };
    expect(result.docId).toBeTruthy();
    const docs = await store.listDocuments();
    expect(docs.find((d) => d.docId === result.docId)?.title).toBe("report");

    // The same token also works as a Bearer credential.
    const viaBearer = await fetch(url, {
      method: "POST",
      headers: { authorization: `Bearer ${token}`, "x-filename": "second.docx" },
      body: body(),
    });
    expect(viaBearer.status).toBe(201);

    // Last-used was touched by verification.
    const tokens = await store.listApiTokens();
    expect(tokens[0]!.lastUsedAt).not.toBeNull();
  }, 15000);
});
