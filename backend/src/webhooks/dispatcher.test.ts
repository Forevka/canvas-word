// End-to-end session-end webhook test (no DB): two WebSocket clients join a doc,
// one submits a change, both disconnect — the room empties, the debounce fires,
// and the registered endpoint receives one HMAC-signed POST. Asserts the
// signature verifies and a delivery row is logged.

import { createHmac } from "node:crypto";
import { createServer, type Server } from "node:http";
import { type AddressInfo } from "node:net";
import { serializeDocument, type Change, type Document, type Op, type Paragraph, type SectionProps } from "@cw/shared";
import { afterEach, beforeAll, describe, expect, it } from "vitest";
import { WebSocket } from "ws";
import { MemoryChangeStore } from "../store/memoryStore";

// Short debounce so the test is fast — must be set before createApp builds the
// Broadcaster (its default reads this env at construction).
process.env.SESSION_END_DEBOUNCE_MS = "40";

let createApp: typeof import("../server").createApp;
beforeAll(async () => {
  ({ createApp } = await import("../server"));
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
      runs: [{ text: "hi", style: { fontFamily: "Georgia", fontSizePx: 16, bold: false, italic: false, underline: false, strikethrough: false, color: "#000" } }],
      style: { align: "left", lineHeight: 1.2, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 },
    } as Paragraph,
  ],
});
const change = (docId: string): Change => ({
  id: "c1",
  docId,
  baseVersion: 0,
  siteId: "siteA",
  origin: "typing",
  ts: 1700000000000,
  userId: "u1",
  ops: [{ type: "insertText", at: { blockId: "p1", offset: 2 }, text: "!" } as Op],
});

const listen = (server: Server): Promise<number> =>
  new Promise((resolve) => server.listen(0, () => resolve((server.address() as AddressInfo).port)));

const servers: Server[] = [];
afterEach(() => {
  for (const s of servers.splice(0)) s.close();
});

describe("session-end webhook", () => {
  it("delivers a signed notification when the last editor disconnects", async () => {
    // 1) Catcher endpoint that records the inbound webhook.
    const received: { body: string; sig: string | undefined; event: string | undefined }[] = [];
    let resolveHit: () => void;
    const hit = new Promise<void>((r) => (resolveHit = r));
    const catcher = createServer((req, res) => {
      const chunks: Buffer[] = [];
      req.on("data", (c) => chunks.push(c as Buffer));
      req.on("end", () => {
        received.push({
          body: Buffer.concat(chunks).toString("utf8"),
          sig: req.headers["x-cw-signature"] as string | undefined,
          event: req.headers["x-cw-event"] as string | undefined,
        });
        res.writeHead(200).end("ok");
        resolveHit();
      });
    });
    servers.push(catcher);
    const catcherPort = await listen(catcher);

    // 2) App + store with a document and a registered webhook.
    const store = new MemoryChangeStore();
    const { docId } = await store.createDocument(serializeDocument(doc()), { createdBy: "u1" });
    await store.upsertUser({ id: "u1", firstName: "Ed", lastName: "Itor" });
    const secret = "test-secret";
    await store.createWebhook({ url: `http://127.0.0.1:${catcherPort}/hook`, secret, events: ["document.session_ended"] });

    const { server } = createApp(store);
    servers.push(server);
    const port = await listen(server);

    // 3) Two collaborators join; one identifies + edits; both leave.
    const wsUrl = `ws://127.0.0.1:${port}/ws?doc=${docId}`;
    const a = new WebSocket(wsUrl);
    const b = new WebSocket(wsUrl);
    await Promise.all([once(a, "open"), once(b, "open")]);
    a.send(JSON.stringify({ type: "hello", siteId: "siteA", user: { id: "u1", firstName: "Ed", lastName: "Itor" } }));
    a.send(JSON.stringify({ type: "submit", change: change(docId) }));
    await delay(60); // let the change append
    a.close();
    b.close();

    // 4) Await the delivery (debounce 40ms + dispatch).
    await Promise.race([hit, delay(5000)]);
    expect(received).toHaveLength(1);

    const got = received[0]!;
    expect(got.event).toBe("document.session_ended");
    const expectedSig = "sha256=" + createHmac("sha256", secret).update(got.body).digest("hex");
    expect(got.sig).toBe(expectedSig);

    const payload = JSON.parse(got.body) as {
      event: string;
      docId: string;
      editors: { id: string }[];
      downloadUrls: { docx: string; pdf: string };
    };
    expect(payload.docId).toBe(docId);
    expect(payload.editors.map((e) => e.id)).toContain("u1");
    expect(payload.downloadUrls.docx).toContain(`/docs/${docId}/export.docx`);

    // 5) The attempt is logged (recordDelivery runs just after the catcher
    //    responds, so poll briefly rather than racing it).
    let deliveries = await store.listDeliveries();
    for (let i = 0; i < 50 && deliveries.length === 0; i++) {
      await delay(20);
      deliveries = await store.listDeliveries();
    }
    expect(deliveries).toHaveLength(1);
    expect(deliveries[0]!.status).toBe("success");
  }, 15000);
});

function once(ws: WebSocket, ev: "open"): Promise<void> {
  return new Promise((resolve, reject) => {
    ws.once(ev, () => resolve());
    ws.once("error", reject);
  });
}
const delay = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));
