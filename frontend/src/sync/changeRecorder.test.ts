import { describe, expect, it } from "vitest";
import { configureIds, type Change, type Op } from "@cw/shared";
import { ChangeRecorder } from "./changeRecorder";

const op = (text: string): Op => ({ type: "insertText", at: { blockId: "p1", offset: 0 }, text });

describe("ChangeRecorder", () => {
  it("stamps docId/siteId/seq and forwards to the sink", () => {
    configureIds("siteX");
    const seen: Change[] = [];
    const rec = new ChangeRecorder("doc-42", (c) => seen.push(c));

    const c0 = rec.record([op("a")], "typing", null, 1000)!;
    const c1 = rec.record([op("b")], "command", null, 1001)!;

    expect(c0.docId).toBe("doc-42");
    expect(c0.siteId).toBe("siteX");
    expect(c0.seq).toBe(0);
    expect(c0.baseVersion).toBe(0);
    expect(c1.seq).toBe(1);
    expect(c1.baseVersion).toBe(1);
    expect(c0.id).not.toBe(c1.id); // unique idempotency keys
    expect(seen).toEqual([c0, c1]); // sink got every change in order
    expect(rec.head()).toBe(2);
    expect(rec.changes()).toEqual([c0, c1]);
  });

  it("ignores empty op groups", () => {
    const rec = new ChangeRecorder("d");
    expect(rec.record([], "typing", null, 0)).toBeNull();
    expect(rec.head()).toBe(0);
  });

  it("returns a copy of the log (callers can't mutate internal state)", () => {
    const rec = new ChangeRecorder("d");
    rec.record([op("x")], "typing", null, 0);
    rec.changes().push({} as Change);
    expect(rec.changes()).toHaveLength(1);
  });
});
