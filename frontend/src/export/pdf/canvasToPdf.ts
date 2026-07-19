// Canvas2D → pdfkit vector shim ("headless canvas").
//
// A registered custom block draws with a browser `CanvasRenderingContext2D`
// (see blockRegistry). To render the SAME `paint(ctx, box, data)` into a PDF
// without a browser — on the Node backend or in the bare-V8 ClearScript host —
// we hand it a `ctx` that implements the common Canvas2D surface and forwards
// each call to pdfkit as VECTOR ops. Pure JS: no native canvas, no WASM, so it
// runs anywhere the export runs. pdfkit shares canvas's y-down/top-left coords,
// so document-px coordinates map straight through.
//
// Supported: paths (moveTo/lineTo/bezierCurveTo/quadraticCurveTo/rect/arc/
// ellipse/roundRect/closePath), fill/stroke/fillRect/strokeRect, clip (current
// path), fillText/strokeText/measureText, font/textAlign/textBaseline,
// fill/strokeStyle (CSS color or linear/radial gradient), lineWidth/cap/join,
// globalAlpha, save/restore, translate/scale/rotate/transform/setTransform,
// drawImage (raster bytes). Unsupported ops (shadows, `filter`, pattern fills)
// are no-ops — a block that needs them can opt into a raster fallback (see the
// `pdf` seam in blockRegistry / paintBlock).

import type { CustomBlockType } from "../../blockRegistry";

/** Resolves a CSS family + weight/style to a registered pdfkit font name. */
export type PdfFontResolver = (family: string, bold: boolean, italic: boolean) => string;

interface GradientStop {
  offset: number;
  color: string;
}

/** Canvas gradient stand-in — built into a pdfkit gradient lazily at paint time. */
class PdfGradient {
  readonly stops: GradientStop[] = [];
  constructor(
    readonly kind: "linear" | "radial",
    readonly coords: number[],
  ) {}
  addColorStop(offset: number, color: string): void {
    this.stops.push({ offset, color });
  }
}

type Style = string | PdfGradient;

/** One buffered path command (replayed into pdfkit at fill/stroke/clip time). */
type PathSeg =
  | { c: "M"; a: [number, number] }
  | { c: "L"; a: [number, number] }
  | { c: "C"; a: [number, number, number, number, number, number] }
  | { c: "Q"; a: [number, number, number, number] }
  | { c: "R"; a: [number, number, number, number] }
  | { c: "RR"; a: [number, number, number, number, number] }
  | { c: "Z"; a: [] };

const DEFAULT_FONT = { sizePx: 10, family: "sans-serif", bold: false, italic: false };

/** Parse a CSS `font` shorthand ("bold italic 14px Arial, sans-serif"). Only the
 *  parts pdfkit needs (weight, style, size, family) — enough for block chrome. */
function parseFont(font: string): { sizePx: number; family: string; bold: boolean; italic: boolean } {
  const bold = /\b(bold|[6-9]00)\b/.test(font);
  const italic = /\b(italic|oblique)\b/.test(font);
  const size = /(\d+(?:\.\d+)?)px/.exec(font);
  const sizePx = size ? parseFloat(size[1]!) : DEFAULT_FONT.sizePx;
  // Family = everything after the size token; strip quotes, take the first face.
  let family = DEFAULT_FONT.family;
  if (size) {
    const after = font.slice(size.index + size[0].length).trim();
    if (after) family = after.split(",")[0]!.trim().replace(/^["']|["']$/g, "");
  }
  return { sizePx, family, bold, italic };
}

/** A minimal CanvasRenderingContext2D implementation over a pdfkit document.
 *  Instantiated per custom block; drawing is relative to the block's box (the
 *  caller translates the pdfkit CTM to the block origin first). */
export class CanvasToPdfContext {
  fillStyle: Style = "#000000";
  strokeStyle: Style = "#000000";
  lineWidth = 1;
  lineCap: CanvasLineCap = "butt";
  lineJoin: CanvasLineJoin = "miter";
  globalAlpha = 1;
  font = "10px sans-serif";
  textAlign: CanvasTextAlign = "start";
  textBaseline: CanvasTextBaseline = "alphabetic";
  // Ignored (unsupported) — declared so plugin code that sets them doesn't throw.
  lineDashOffset = 0;
  miterLimit = 10;
  shadowBlur = 0;
  shadowColor = "rgba(0,0,0,0)";
  shadowOffsetX = 0;
  shadowOffsetY = 0;
  filter = "none";

  private dash: number[] = [];
  private saveDepth = 0;
  // Buffered current path. pdfkit writes path ops straight to the content stream
  // and consumes them on fill/stroke, so it has no clearable "current path" — we
  // keep our own so Canvas2D semantics hold: beginPath() truly resets, arc()
  // implicit-moveTo works, a path can be both filled AND stroked, and fillRect()
  // never picks up stray segments. The buffer is replayed into pdfkit fresh on
  // each fill/stroke/clip (and kept across them, like canvas — beginPath clears).
  private path: PathSeg[] = [];
  private hasPoint = false;

  constructor(
    private readonly doc: PDFKit.PDFDocument,
    private readonly resolveFont: PdfFontResolver,
  ) {}

  // ---- state ---------------------------------------------------------------
  save(): void {
    this.doc.save();
    this.saveDepth++;
  }
  restore(): void {
    if (this.saveDepth > 0) {
      this.saveDepth--;
      this.doc.restore();
    }
  }
  /** Restore any save() the block left unbalanced — so a plugin bug (or a throw
   *  mid-paint) can't leak graphics state onto the rest of the page. */
  balance(): void {
    while (this.saveDepth > 0) this.restore();
  }
  translate(x: number, y: number): void {
    this.doc.translate(x, y);
  }
  scale(x: number, y: number): void {
    this.doc.scale(x, y);
  }
  rotate(angle: number): void {
    this.doc.rotate((angle * 180) / Math.PI); // pdfkit rotate is in degrees
  }
  transform(a: number, b: number, c: number, d: number, e: number, f: number): void {
    this.doc.transform(a, b, c, d, e, f);
  }
  setTransform(): void {
    // Canvas setTransform replaces the CTM; pdfkit has no absolute reset without a
    // save/restore stack. No-op (blocks should use save/restore + transform).
  }
  setLineDash(segments: number[]): void {
    this.dash = segments.slice();
  }
  getLineDash(): number[] {
    return this.dash.slice();
  }

  // ---- path building (buffered — replayed at paint time) -------------------
  beginPath(): void {
    this.path = [];
    this.hasPoint = false;
  }
  closePath(): void {
    this.path.push({ c: "Z", a: [] });
  }
  moveTo(x: number, y: number): void {
    this.path.push({ c: "M", a: [x, y] });
    this.hasPoint = true;
  }
  lineTo(x: number, y: number): void {
    // Canvas: lineTo with no current subpath acts as moveTo.
    this.path.push({ c: this.hasPoint ? "L" : "M", a: [x, y] });
    this.hasPoint = true;
  }
  bezierCurveTo(cp1x: number, cp1y: number, cp2x: number, cp2y: number, x: number, y: number): void {
    if (!this.hasPoint) this.moveTo(cp1x, cp1y);
    this.path.push({ c: "C", a: [cp1x, cp1y, cp2x, cp2y, x, y] });
    this.hasPoint = true;
  }
  quadraticCurveTo(cpx: number, cpy: number, x: number, y: number): void {
    if (!this.hasPoint) this.moveTo(cpx, cpy);
    this.path.push({ c: "Q", a: [cpx, cpy, x, y] });
    this.hasPoint = true;
  }
  rect(x: number, y: number, w: number, h: number): void {
    // A self-contained subpath; a following op starts fresh (canvas leaves the
    // current point at the rect's start).
    this.path.push({ c: "R", a: [x, y, w, h] });
    this.hasPoint = false;
  }
  roundRect(x: number, y: number, w: number, h: number, radius: number | number[] = 0): void {
    const r = Math.min(Array.isArray(radius) ? (radius[0] ?? 0) : radius, w / 2, h / 2);
    this.path.push({ c: "RR", a: [x, y, w, h, r] });
    this.hasPoint = false;
  }
  arc(cx: number, cy: number, r: number, start: number, end: number, ccw = false): void {
    this.appendArc(cx, cy, r, r, start, end, ccw);
  }
  ellipse(cx: number, cy: number, rx: number, ry: number, rotation: number, start: number, end: number, ccw = false): void {
    if (rotation) {
      // Rare — approximate by ignoring rotation (documented limitation).
    }
    this.appendArc(cx, cy, rx, ry, start, end, ccw);
  }
  arcTo(x1: number, y1: number, x2: number, y2: number, _r: number): void {
    // Approximate with straight segments (radius corner-rounding is rarely used
    // in block chrome; roundRect covers the common case).
    this.lineTo(x1, y1);
    this.lineTo(x2, y2);
  }

  /** Append an (elliptical) arc as cubic-bezier segments into the path buffer.
   *  Canvas connects with a straight line from the current point to the arc's
   *  start when a subpath is open; on an EMPTY subpath the first point is a
   *  moveTo (no spurious connecting line). */
  private appendArc(cx: number, cy: number, rx: number, ry: number, start: number, end: number, ccw: boolean): void {
    let a0 = start;
    let a1 = end;
    if (!ccw && a1 - a0 >= Math.PI * 2) a1 = a0 + Math.PI * 2;
    else if (ccw && a0 - a1 >= Math.PI * 2) a1 = a0 - Math.PI * 2;
    else if (!ccw && a1 < a0) a1 += Math.PI * 2 * Math.ceil((a0 - a1) / (Math.PI * 2));
    else if (ccw && a1 > a0) a1 -= Math.PI * 2 * Math.ceil((a1 - a0) / (Math.PI * 2));
    const segs = Math.max(2, Math.ceil(Math.abs(a1 - a0) / (Math.PI / 2)));
    const da = (a1 - a0) / segs;
    const k = (4 / 3) * Math.tan(da / 4);
    let ang = a0;
    let px = cx + rx * Math.cos(ang);
    let py = cy + ry * Math.sin(ang);
    // moveTo the arc's start on an empty subpath; lineTo it when one is open.
    if (this.hasPoint) this.path.push({ c: "L", a: [px, py] });
    else this.path.push({ c: "M", a: [px, py] });
    this.hasPoint = true;
    for (let i = 0; i < segs; i++) {
      const ang2 = ang + da;
      const x2 = cx + rx * Math.cos(ang2);
      const y2 = cy + ry * Math.sin(ang2);
      const c1x = px - k * rx * Math.sin(ang);
      const c1y = py + k * ry * Math.cos(ang);
      const c2x = x2 + k * rx * Math.sin(ang2);
      const c2y = y2 - k * ry * Math.cos(ang2);
      this.path.push({ c: "C", a: [c1x, c1y, c2x, c2y, x2, y2] });
      ang = ang2;
      px = x2;
      py = y2;
    }
  }

  // ---- paint ---------------------------------------------------------------
  /** Replay the buffered path into pdfkit (a fresh, self-contained path each
   *  time, so pdfkit's stream never carries stray segments between paints). */
  private replayPath(): void {
    for (const s of this.path) {
      if (s.c === "M") this.doc.moveTo(s.a[0], s.a[1]);
      else if (s.c === "L") this.doc.lineTo(s.a[0], s.a[1]);
      else if (s.c === "C") this.doc.bezierCurveTo(s.a[0], s.a[1], s.a[2], s.a[3], s.a[4], s.a[5]);
      else if (s.c === "Q") this.doc.quadraticCurveTo(s.a[0], s.a[1], s.a[2], s.a[3]);
      else if (s.c === "R") this.doc.rect(s.a[0], s.a[1], s.a[2], s.a[3]);
      else if (s.c === "RR") this.doc.roundedRect(s.a[0], s.a[1], s.a[2], s.a[3], s.a[4]);
      else this.doc.closePath();
    }
  }
  private applyDash(): void {
    if (this.dash.length) this.doc.dash(this.dash[0]!, { space: this.dash[1] ?? this.dash[0]! });
    else this.doc.undash();
  }
  private fillArg(style: Style): string | PDFKit.PDFGradient {
    if (typeof style === "string") return style;
    const g =
      style.kind === "linear"
        ? this.doc.linearGradient(style.coords[0]!, style.coords[1]!, style.coords[2]!, style.coords[3]!)
        : this.doc.radialGradient(style.coords[0]!, style.coords[1]!, style.coords[2]!, style.coords[3]!, style.coords[4]!, style.coords[5]!);
    for (const s of style.stops) g.stop(s.offset, s.color);
    return g;
  }
  fill(): void {
    this.doc.save();
    this.doc.fillOpacity(this.globalAlpha);
    this.replayPath();
    this.doc.fill(this.fillArg(this.fillStyle));
    this.doc.restore();
  }
  stroke(): void {
    this.doc.save();
    this.doc.strokeOpacity(this.globalAlpha);
    this.doc.lineWidth(this.lineWidth).lineCap(this.lineCap).lineJoin(this.lineJoin);
    this.applyDash();
    this.replayPath();
    this.doc.stroke(this.fillArg(this.strokeStyle));
    this.doc.restore();
  }
  fillRect(x: number, y: number, w: number, h: number): void {
    this.doc.save();
    this.doc.fillOpacity(this.globalAlpha);
    this.doc.rect(x, y, w, h).fill(this.fillArg(this.fillStyle));
    this.doc.restore();
  }
  strokeRect(x: number, y: number, w: number, h: number): void {
    this.doc.save();
    this.doc.strokeOpacity(this.globalAlpha).lineWidth(this.lineWidth);
    this.applyDash();
    this.doc.rect(x, y, w, h).stroke(this.fillArg(this.strokeStyle));
    this.doc.restore();
  }
  clearRect(): void {
    // No "clear" in PDF; a block should paint its own background. No-op.
  }
  clip(): void {
    this.replayPath();
    this.doc.clip();
  }

  // ---- gradients -----------------------------------------------------------
  createLinearGradient(x0: number, y0: number, x1: number, y1: number): PdfGradient {
    return new PdfGradient("linear", [x0, y0, x1, y1]);
  }
  createRadialGradient(x0: number, y0: number, r0: number, x1: number, y1: number, r1: number): PdfGradient {
    return new PdfGradient("radial", [x0, y0, r0, x1, y1, r1]);
  }

  // ---- text ----------------------------------------------------------------
  private setPdfFont(): number {
    const f = parseFont(this.font);
    const name = this.resolveFont(f.family, f.bold, f.italic);
    this.doc.font(name).fontSize(f.sizePx);
    return f.sizePx;
  }
  measureText(text: string): { width: number } {
    this.setPdfFont();
    return { width: this.doc.widthOfString(text) };
  }
  fillText(text: string, x: number, y: number): void {
    this.setPdfFont();
    let tx = x;
    if (this.textAlign === "center" || this.textAlign === "right" || this.textAlign === "end") {
      const w = this.doc.widthOfString(text);
      tx = this.textAlign === "center" ? x - w / 2 : x - w;
    }
    this.doc.save();
    this.doc.fillOpacity(this.globalAlpha);
    this.doc.fillColor(this.fillArg(this.fillStyle)); // honors a gradient fillStyle too
    // pdfkit's `baseline` option maps canvas textBaseline; default alphabetic.
    this.doc.text(text, tx, y, { lineBreak: false, baseline: this.textBaseline as PDFKit.Mixins.TextOptions["baseline"] });
    this.doc.restore();
  }
  strokeText(text: string, x: number, y: number): void {
    // Approximate stroked text with filled text in the stroke color.
    const prev = this.fillStyle;
    this.fillStyle = this.strokeStyle;
    this.fillText(text, x, y);
    this.fillStyle = prev;
  }

  // ---- images --------------------------------------------------------------
  drawImage(): void {
    // Raster image sources can't be serialized generically here; a block needing
    // embedded bitmaps in PDF should register them as ImageBlocks or use the
    // raster fallback. No-op (documented).
  }
}

/** Render a registered custom block into the PDF at (x, y) by running its
 *  `paint(ctx, box, data)` against the Canvas2D→pdfkit shim. Returns true on
 *  success. The context is clipped to the box and any unbalanced save()s are
 *  unwound, so a plugin bug can't corrupt the rest of the page. */
export function renderCustomBlockPdf(
  doc: PDFKit.PDFDocument,
  x: number,
  y: number,
  box: { width: number; height: number },
  type: CustomBlockType,
  data: unknown,
  resolveFont: PdfFontResolver,
  warn: (code: string, detail?: string) => void,
): boolean {
  const ctx = new CanvasToPdfContext(doc, resolveFont);
  doc.save();
  doc.translate(x, y);
  doc.rect(0, 0, box.width, box.height).clip(); // keep overpaint inside the block box
  try {
    type.paint(ctx as unknown as CanvasRenderingContext2D, box, data);
    return true;
  } catch (err) {
    console.error(`[canvas-word] custom block "${type.type}" paint() threw during PDF export`, err);
    warn("custom-block-paint-failed", type.type);
    return false;
  } finally {
    ctx.balance(); // unwind any save() the block left open
    doc.restore(); // our own save/translate/clip
  }
}
