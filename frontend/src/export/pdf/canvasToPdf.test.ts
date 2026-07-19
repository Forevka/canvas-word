import { describe, it, expect } from "vitest";
import { renderCustomBlockPdf, CanvasToPdfContext } from "./canvasToPdf";
import type { CustomBlockType } from "../../blockRegistry";

// A minimal chainable pdfkit-doc mock that records the calls the shim makes.
function mockDoc() {
  const calls: { m: string; args: unknown[] }[] = [];
  const rec = (m: string) => (...args: unknown[]) => { calls.push({ m, args }); return doc; };
  const doc: Record<string, unknown> = {};
  for (const m of [
    "save","restore","translate","scale","rotate","transform","rect","roundedRect","circle",
    "moveTo","lineTo","bezierCurveTo","quadraticCurveTo","closePath","fill","stroke","clip",
    "fillColor","strokeColor","fillOpacity","strokeOpacity","lineWidth","lineCap","lineJoin",
    "dash","undash","font","fontSize","text",
  ]) doc[m] = rec(m);
  doc["widthOfString"] = (s: string) => s.length * 6; // deterministic metric
  doc["linearGradient"] = (...a: unknown[]) => { calls.push({ m: "linearGradient", args: a }); return { stop: rec("stop") }; };
  doc["radialGradient"] = (...a: unknown[]) => { calls.push({ m: "radialGradient", args: a }); return { stop: rec("stop") }; };
  return { doc: doc as unknown as PDFKit.PDFDocument, calls, names: () => calls.map((c) => c.m) };
}

const resolveFont = () => "Body";
const warnLog: string[] = [];
const warn = (code: string) => { warnLog.push(code); };

const render = (type: CustomBlockType, data: unknown = {}) => {
  const mk = mockDoc();
  const ok = renderCustomBlockPdf(mk.doc, 100, 200, { width: 300, height: 80 }, type, data, resolveFont, warn);
  return { ok, ...mk };
};

describe("renderCustomBlockPdf (Canvas2D → pdfkit shim)", () => {
  it("positions + clips the block box, then runs paint", () => {
    const { ok, names } = render({ type: "t", measure: () => ({ height: 80 }), paint: () => {} });
    expect(ok).toBe(true);
    // save → translate(100,200) → rect(0,0,300,80) → clip → … → restore
    expect(names().slice(0, 4)).toEqual(["save", "translate", "rect", "clip"]);
    expect(names()[names().length - 1]).toBe("restore");
  });

  it("forwards fillRect as rect(...).fill(color)", () => {
    const { calls } = render({
      type: "t", measure: () => ({ height: 80 }),
      paint: (ctx) => { ctx.fillStyle = "#ff0000"; ctx.fillRect(1, 2, 3, 4); },
    });
    const rectCall = calls.find((c) => c.m === "rect" && c.args[2] === 3);
    expect(rectCall).toBeTruthy();
    expect(calls.find((c) => c.m === "fill" && c.args[0] === "#ff0000")).toBeTruthy();
  });

  it("draws text via font/fontSize/text and measures via widthOfString", () => {
    let measured = -1;
    const { calls } = render({
      type: "t", measure: () => ({ height: 80 }),
      paint: (ctx) => {
        ctx.font = "bold 14px Arial";
        measured = ctx.measureText("hello").width;
        ctx.fillText("hello", 5, 20);
      },
    });
    expect(measured).toBe(30); // 5 chars * 6
    expect(calls.find((c) => c.m === "fontSize" && c.args[0] === 14)).toBeTruthy();
    expect(calls.find((c) => c.m === "text" && c.args[0] === "hello")).toBeTruthy();
  });

  it("builds a linear gradient fill", () => {
    const { calls } = render({
      type: "t", measure: () => ({ height: 80 }),
      paint: (ctx) => {
        const g = ctx.createLinearGradient(0, 0, 10, 0);
        g.addColorStop(0, "#000"); g.addColorStop(1, "#fff");
        ctx.fillStyle = g; ctx.fillRect(0, 0, 10, 10);
      },
    });
    expect(calls.find((c) => c.m === "linearGradient")).toBeTruthy();
    expect(calls.filter((c) => c.m === "stop")).toHaveLength(2);
  });

  it("a bare beginPath()+arc()+fill() starts the subpath with moveTo (no stray line)", () => {
    const { calls } = render({
      type: "t", measure: () => ({ height: 80 }),
      paint: (ctx) => { ctx.beginPath(); ctx.arc(30, 30, 10, 0, Math.PI * 2); ctx.fill(); },
    });
    // The replayed path must OPEN with moveTo, and a full circle carries no lineTo.
    const firstPathOp = calls.find((c) => c.m === "moveTo" || c.m === "lineTo");
    expect(firstPathOp?.m).toBe("moveTo");
    expect(calls.find((c) => c.m === "lineTo")).toBeFalsy();
  });

  it("fillRect after an unpainted open path is isolated (no stray segments filled)", () => {
    const { calls } = render({
      type: "t", measure: () => ({ height: 80 }),
      paint: (ctx) => { ctx.beginPath(); ctx.moveTo(0, 0); ctx.lineTo(50, 50); ctx.fillRect(1, 2, 3, 4); },
    });
    // The open path was never filled/stroked, so it's never replayed into pdfkit.
    expect(calls.find((c) => c.m === "lineTo")).toBeFalsy();
    expect(calls.find((c) => c.m === "rect" && c.args[2] === 3)).toBeTruthy();
  });

  it("the same buffered path can be BOTH filled and stroked (canvas keeps the path)", () => {
    const { calls } = render({
      type: "t", measure: () => ({ height: 80 }),
      paint: (ctx) => { ctx.beginPath(); ctx.rect(0, 0, 10, 10); ctx.fill(); ctx.stroke(); },
    });
    expect(calls.filter((c) => c.m === "fill")).toHaveLength(1);
    expect(calls.filter((c) => c.m === "stroke")).toHaveLength(1);
    // The block's 10×10 rect is replayed for BOTH paints (the 300-wide box-clip
    // rect from the render wrapper is separate).
    expect(calls.filter((c) => c.m === "rect" && c.args[2] === 10)).toHaveLength(2);
  });

  it("unwinds a save() the block left open (no state leak)", () => {
    const { calls } = render({
      type: "t", measure: () => ({ height: 80 }),
      paint: (ctx) => { ctx.save(); /* never restores */ },
    });
    // The outer save + the block's leaked save must both be restored.
    const saves = calls.filter((c) => c.m === "save").length;
    const restores = calls.filter((c) => c.m === "restore").length;
    expect(restores).toBe(saves);
  });

  it("a throwing paint returns false, warns, and stays balanced", () => {
    warnLog.length = 0;
    const { ok, calls } = render({
      type: "boom", measure: () => ({ height: 80 }),
      paint: (ctx) => { ctx.save(); throw new Error("bad paint"); },
    });
    expect(ok).toBe(false);
    expect(warnLog).toContain("custom-block-paint-failed");
    expect(calls.filter((c) => c.m === "restore").length).toBe(calls.filter((c) => c.m === "save").length);
  });

  it("exposes a subset context that satisfies the CanvasRenderingContext2D shape", () => {
    // The paint contract is typed CanvasRenderingContext2D; the shim is passed via a
    // cast. Sanity-check a few members exist so plugin code doesn't hit undefined.
    const mk = mockDoc();
    const ctx = new CanvasToPdfContext(mk.doc, resolveFont);
    for (const m of ["beginPath", "moveTo", "lineTo", "arc", "fill", "stroke", "fillText", "measureText", "save", "restore"]) {
      expect(typeof (ctx as unknown as Record<string, unknown>)[m]).toBe("function");
    }
  });
});
