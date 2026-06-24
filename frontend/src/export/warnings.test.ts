// WarningSink dedupes by code (one entry per code; repeats bump count) and keeps
// the FIRST detail seen for a code. Mirrors the importer's WarningSink behavior.

import { describe, expect, it } from "vitest";
import { WarningSink } from "./warnings";

describe("WarningSink", () => {
  it("starts empty", () => {
    expect(new WarningSink().list()).toEqual([]);
  });

  it("records a single warning with count 1", () => {
    const sink = new WarningSink();
    sink.warn("image-format-unsupported", "tiff");
    expect(sink.list()).toEqual([
      { code: "image-format-unsupported", detail: "tiff", count: 1 },
    ]);
  });

  it("records a warning without a detail (no detail key)", () => {
    const sink = new WarningSink();
    sink.warn("some-code");
    const [w] = sink.list();
    expect(w).toEqual({ code: "some-code", count: 1 });
    expect("detail" in w!).toBe(false);
  });

  it("bumps count for a repeated code instead of adding a new entry", () => {
    const sink = new WarningSink();
    sink.warn("dup");
    sink.warn("dup");
    sink.warn("dup");
    expect(sink.list()).toEqual([{ code: "dup", count: 3 }]);
  });

  it("keeps the FIRST detail when later repeats supply a different one", () => {
    const sink = new WarningSink();
    sink.warn("img", "png");
    sink.warn("img", "gif"); // ignored — only count bumps
    expect(sink.list()).toEqual([{ code: "img", detail: "png", count: 2 }]);
  });

  it("does NOT add a detail when the first call had none but a repeat does", () => {
    const sink = new WarningSink();
    sink.warn("x"); // no detail
    sink.warn("x", "late-detail");
    expect(sink.list()).toEqual([{ code: "x", count: 2 }]);
  });

  it("keeps distinct codes as separate entries in insertion order", () => {
    const sink = new WarningSink();
    sink.warn("b", "second");
    sink.warn("a", "first");
    sink.warn("b"); // bumps b
    expect(sink.list()).toEqual([
      { code: "b", detail: "second", count: 2 },
      { code: "a", detail: "first", count: 1 },
    ]);
  });
});
