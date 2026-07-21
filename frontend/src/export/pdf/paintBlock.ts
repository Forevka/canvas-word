// PlacedBlock -> pdfkit, a constant-for-constant inverse of src/paint/renderer.ts.
// Same baseline formula (block.y + line.y + line.ascent), same sub/super shifts,
// underline/strike offsets, default grid color, footnote-rule width and leader
// dash patterns — so a page in the PDF matches the canvas pixel-for-pixel (modulo
// metric-clone glyph shapes). pdfkit shares y-down, top-left coords with canvas.

import type { CellBorder } from "@cw/shared";
import { DEFAULT_CHAR_STYLE } from "@cw/shared";
import type { LineBox, Page, PlacedBlock, PlacedShape, PlacedTableCell } from "../../layout/layoutTree";
import type { MathBox } from "../../layout/math/mathBox";
import { MATH_FONT_FAMILY } from "../../fonts/clones";
import {
  cellBorderDash,
  cellBorderWidth,
  decorationThickness,
  DEFAULT_GRID_COLOR,
  doubleBorderGap,
  IMAGE_PLACEHOLDER_COLOR,
  leaderDash,
  leaderWidth,
  normalizeLinkBlue,
  resolveShapePaint,
  runPaint,
  strikeOffset,
  SUB_SUPER_SCALE,
  TOC_LEADER_COLOR,
  TOC_LEADER_DASH,
  TOC_LEADER_GAP_PX,
  UNDERLINE_OFFSET_PX,
  underlinePlan,
  doubleUnderlineGap,
  underlineWavePoints,
  runVerticalShift,
  widthScale,
  doubleStrikeOffsets,
  bulletShapeFor,
  paraDecorBox,
} from "../../paint/paintStyle";
import type { RunPaint, BulletShape } from "../../paint/paintStyle";
import { getBlockType } from "../../blockRegistry";
import { renderCustomBlockPdf } from "./canvasToPdf";

/** Stroke a run's underline into the PDF. "single" keeps the historical filled
 *  rect; richer styles stroke per the shared plan (dashed/dotted approximated with
 *  pdfkit's single dash pattern, double = two lines, wave = a sampled polyline). */
function paintUnderlinePdf(
  doc: PDFKit.PDFDocument,
  x: number,
  yTop: number,
  width: number,
  fontSizePx: number,
  rp: RunPaint,
): void {
  if (rp.underlineStyle === "single") {
    doc.rect(x, yTop, width, decorationThickness(fontSizePx)).fill(rp.underlineColor);
    return;
  }
  const plan = underlinePlan(rp.underlineStyle, fontSizePx);
  const yCenter = yTop + plan.thickness / 2;
  doc.save();
  doc.lineWidth(plan.thickness).strokeColor(rp.underlineColor);
  // pdfkit's .dash() only models a single on/off pair, so multi-segment patterns
  // (dotDash/dotDotDash) go through the raw PDF dash array to keep canvas parity.
  if (plan.dash.length === 2) doc.dash(plan.dash[0]!, { space: plan.dash[1]! });
  else if (plan.dash.length > 2) doc.addContent(`[${plan.dash.map((d) => Number(d.toFixed(2))).join(" ")}] 0 d`);
  if (plan.wave) {
    const pts = underlineWavePoints(x, yCenter, width, fontSizePx);
    doc.moveTo(pts[0]!.x, pts[0]!.y);
    for (let i = 1; i < pts.length; i++) doc.lineTo(pts[i]!.x, pts[i]!.y);
    doc.stroke();
  } else {
    doc.moveTo(x, yCenter).lineTo(x + width, yCenter).stroke();
    if (plan.double) {
      const dy = doubleUnderlineGap(plan.thickness);
      doc.moveTo(x, yCenter + dy).lineTo(x + width, yCenter + dy).stroke();
    }
  }
  doc.undash();
  doc.restore();
}

export interface PaintCtx {
  doc: PDFKit.PDFDocument;
  /** Register (idempotently) the bundled face and return its pdfkit name. */
  font: (family: string, bold: boolean, italic: boolean) => string;
  /**
   * Split `text` into face-homogeneous segments, registering every face used
   * (including fallbacks) with pdfkit so they are subset-embedded. Returns an
   * array of `{ text, pdfkitName, font }` tuples; paintLine paints each tuple
   * in sequence at the right x position.
   *
   * Optional so tests that only exercise non-text paint paths (markers, bullets)
   * can supply a minimal PaintCtx. When absent, paintLine falls back to a
   * single-segment paint using `ctx.font()`.
   */
  textSegments?: (
    text: string,
    family: string,
    bold: boolean,
    italic: boolean,
  ) => Array<{ text: string; pdfkitName: string; font: import("fontkit").Font }>;
  /** Image bytes for a src, or undefined if unresolved/unsupported. */
  image: (src: string) => Uint8Array | undefined;
  warn: (code: string, detail?: string) => void;
  /** Resolve a TOC target blockId or an in-document "#anchor" to a registered
   *  named-destination string, or undefined when the target doesn't exist. The
   *  caller emits a GoTo annotation only when a name comes back. */
  destName?: (ref: { blockId?: string; anchor?: string }) => string | undefined;
}

const firstFamily = (stack: string): string => (stack.split(",")[0] ?? "sans-serif").trim();

/** Paint a page's margin line numbers (w:lnNumType). Each label is already
 *  pre-measured and right-aligned (`x` is its left edge, `baseline` the text
 *  baseline), so this just selects the font and draws — the PDF counterpart of
 *  the canvas renderer's line-number pass. */
export function paintLineNumbers(ctx: PaintCtx, lineNumbers: NonNullable<Page["lineNumbers"]>): void {
  const { doc } = ctx;
  for (const ln of lineNumbers) {
    const s = ln.style;
    const name = ctx.font(firstFamily(s.fontFamily), !!s.bold, !!s.italic);
    doc
      .font(name)
      .fontSize(s.fontSizePx)
      .fillColor(s.color as string)
      .text(ln.text, ln.x, ln.baseline, { lineBreak: false, baseline: "alphabetic" });
  }
}

export function paintBlock(ctx: PaintCtx, block: PlacedBlock): void {
  const { doc } = ctx;

  if (block.image) {
    const bytes = ctx.image(block.image.src);
    const { width, height, clip, crop } = block.image;
    if (clip) {
      doc.save();
      doc.rect(clip.x, clip.y, clip.width, clip.height).clip();
    }
    if (bytes) {
      try {
        if (crop) {
          // pdfkit can't source-crop, so clip to the box and draw the full image
          // scaled up so its visible [left,1-right]×[top,1-bottom] window fills it.
          const fracW = Math.max(0.0001, 1 - crop.left - crop.right);
          const fracH = Math.max(0.0001, 1 - crop.top - crop.bottom);
          const fullW = width / fracW;
          const fullH = height / fracH;
          doc.save();
          doc.rect(block.x, block.y, width, height).clip();
          try {
            drawImage(doc, toBuffer(bytes), block.x - crop.left * fullW, block.y - crop.top * fullH, fullW, fullH);
          } finally {
            // Always pop the clip, even if drawImage throws on an undecodable
            // image: otherwise the clip leaks and blanks the rest of the page
            // (body, header, footer). The outer catch then draws the placeholder
            // unclipped. See the white-on-white page regression from SVG images.
            doc.restore();
          }
        } else {
          drawImage(doc, toBuffer(bytes), block.x, block.y, width, height);
        }
      } catch {
        ctx.warn("image-format-unsupported", block.image.src);
        placeholderBox(ctx, block.x, block.y, width, height);
      }
    } else {
      placeholderBox(ctx, block.x, block.y, width, height);
    }
    if (clip) doc.restore();
    return;
  }

  if (block.shape) {
    paintPlacedShapePdf(ctx, block.shape, block.x, block.y);
    return;
  }

  if (block.equation) {
    drawMathBoxPdf(ctx, block.equation.box, block.x, block.y + block.equation.baseline, DEFAULT_CHAR_STYLE.color);
    return;
  }

  if (block.custom) {
    // Render the registered block's paint() into the PDF via the Canvas2D→pdfkit
    // vector shim — the SAME paint used on screen. Requires the type to be
    // registered in THIS export context (true on the Node backend / ClearScript
    // host, and in the browser inline-export path). A `pdf: "raster"` type opts
    // out (a future CanvasKit fallback); until that exists it reserves + warns.
    const type = getBlockType(block.custom.customType);
    if (type && type.pdf !== "raster") {
      renderCustomBlockPdf(
        ctx.doc,
        block.x,
        block.y,
        { width: block.custom.width, height: block.custom.height },
        type,
        block.custom.data,
        (family, bold, italic) => ctx.font(family, bold, italic),
        ctx.warn,
      );
      return;
    }
    // Unregistered here (e.g. a browser worker with no registry) or raster-only:
    // reserve the height and warn so the blank gap isn't a silent surprise.
    ctx.warn("custom-block-not-rendered", block.custom.customType);
    return;
  }

  if (block.table) {
    const rows = block.table.rows;
    // 1) fills, 2) contents, 3) borders — so a neighbour's fill never clips a
    // shared edge that was already drawn.
    for (const row of rows) {
      for (const cell of row.cells) {
        if (cell.shading) doc.rect(cell.x, cell.y, cell.width, cell.height).fill(cell.shading);
      }
    }
    for (const row of rows) {
      for (const cell of row.cells) {
        const clip = cell.contentClip;
        if (clip) {
          doc.save();
          doc.rect(clip.x, clip.y, clip.width, clip.height).clip();
        }
        // Vertical text: rotate the cell's local content frame into place so the PDF
        // matches the canvas. Mirror the canvas order exactly — translate to the
        // origin, then rotate (pdfkit shares canvas' y-down space, so the same angle
        // rotates the same way). pdfkit rotates in DEGREES.
        const rot = cell.rotation;
        if (rot) {
          doc.save();
          doc.translate(rot.originX, rot.originY);
          doc.rotate((rot.angle * 180) / Math.PI);
        }
        for (const cb of cell.blocks) paintBlock(ctx, cb);
        if (rot) doc.restore();
        if (clip) doc.restore();
      }
    }
    for (const row of rows) {
      for (const cell of row.cells) paintCellBorders(ctx, cell);
    }
    return;
  }

  // Paragraph shading (w:shd) — a fill behind the text box, beneath the runs.
  if (block.paraDecor?.shading) {
    const box = paraDecorBox(block.x, block.y, block.paraDecor);
    doc.rect(box.x, box.y, box.width, box.height).fill(block.paraDecor.shading);
  }

  // List marker — paint-only, first line's baseline in the hanging indent.
  const firstLine = block.lines[0];
  if (block.marker && firstLine) {
    const s = block.marker.style;
    const markerBaseline = block.y + firstLine.y + firstLine.ascent;
    // The three default bullets render as vector shapes so ◦/▪ (absent from the
    // bundled Latin font subset) never draw as .notdef/tofu; everything else
    // (numbers, custom glyphs) paints as text. Mirrors the canvas renderer.
    const shape = bulletShapeFor(block.marker.text, s.fontSizePx, block.marker.x, markerBaseline);
    if (shape) {
      drawBulletShape(doc, shape, s.color as string);
    } else {
      const name = ctx.font(firstFamily(s.fontFamily), !!s.bold, !!s.italic);
      doc
        .font(name)
        .fontSize(s.fontSizePx)
        .fillColor(s.color as string)
        .text(block.marker.text, block.marker.x, markerBaseline, {
          lineBreak: false,
          baseline: "alphabetic",
        });
    }
  }

  // TOC decoration — page number + dot leader on its line.
  if (block.toc) {
    const line = block.lines[block.toc.lineIndex];
    if (line) {
      const baseline = block.y + line.y + line.ascent;
      const s = block.toc.style;
      const color = normalizeLinkBlue(s.color);
      const name = ctx.font(firstFamily(s.fontFamily), !!s.bold, !!s.italic);
      doc
        .font(name)
        .fontSize(s.fontSizePx)
        .fillColor(color)
        .text(block.toc.numText, block.toc.numX, baseline, { lineBreak: false, baseline: "alphabetic" });
      const lastFrag = line.fragments[line.fragments.length - 1];
      const fromX = block.x + (lastFrag ? lastFrag.x + lastFrag.width : 0) + TOC_LEADER_GAP_PX;
      const toX = block.toc.numX - TOC_LEADER_GAP_PX;
      if (toX > fromX) {
        doc.save();
        doc.lineWidth(1).strokeColor(TOC_LEADER_COLOR).dash(TOC_LEADER_DASH[0]!, { space: TOC_LEADER_DASH[1]! });
        doc.moveTo(fromX, baseline).lineTo(toX, baseline).stroke();
        doc.undash().restore();
      }
    }
    // Whole entry is a clickable GoTo to its heading (Ctrl+click in the editor).
    const name = ctx.destName?.({ blockId: block.toc.targetId });
    if (name) {
      for (const ln of block.lines) {
        const last = ln.fragments[ln.fragments.length - 1];
        const right = Math.max(last ? block.x + last.x + last.width : block.x, block.toc.numX);
        if (right > block.x) doc.goTo(block.x, block.y + ln.y, right - block.x, ln.height, name);
      }
    }
  }

  for (const line of block.lines) {
    const baselineY = block.y + line.y + line.ascent;
    if (line.leaders) {
      for (const ld of line.leaders) {
        doc.save();
        doc.lineWidth(leaderWidth(ld.fontSizePx)).strokeColor(normalizeLinkBlue(ld.color));
        const dash = leaderDash(ld.kind);
        if (dash.length) doc.dash(dash[0]!, { space: dash[1]! });
        doc.moveTo(block.x + ld.x1, baselineY).lineTo(block.x + ld.x2, baselineY).stroke();
        doc.undash().restore();
      }
    }
    paintLine(ctx, block, line, baselineY);
  }

  // Paragraph border box (w:pBdr) — over the shading and text, like cell edges.
  // `between` is not drawn for a standalone paragraph (matches the canvas renderer).
  if (block.paraDecor?.borders) {
    const d = block.paraDecor;
    const b = d.borders!;
    // Box expanded outward by the border-to-text padding (paraDecorBox), so the
    // rules sit outside the glyphs — identical to the canvas renderer.
    const box = paraDecorBox(block.x, block.y, d);
    const x = box.x;
    const yT = box.y;
    const xR = box.x + box.width;
    const yB = box.y + box.height;
    strokeCellEdge(ctx, b.top, x, yT, xR, yT, 0, 1);
    strokeCellEdge(ctx, b.bottom, x, yB, xR, yB, 0, -1);
    strokeCellEdge(ctx, b.left, x, yT, x, yB, 1, 0);
    strokeCellEdge(ctx, b.right, xR, yT, xR, yB, -1, 0);
  }
}

/** Paint a MathBox to pdfkit at an absolute (originX, baselineY). Shared by block
 *  equations and inline equation fragments. */
function drawMathBoxPdf(ctx: PaintCtx, box: MathBox, originX: number, baselineY: number, color: string): void {
  const { doc } = ctx;
  const family = firstFamily(MATH_FONT_FAMILY);
  for (const r of box.rules) doc.rect(originX + r.x, baselineY + r.y, r.w, r.h).fill(color);
  for (const g of box.glyphs) {
    const name = ctx.font(family, g.bold, g.italic);
    doc.font(name).fontSize(g.sizePx).fillColor(color).text(g.text, originX + g.x, baselineY + g.y, { lineBreak: false, baseline: "alphabetic" });
  }
}

function paintLine(ctx: PaintCtx, block: PlacedBlock, line: LineBox, baselineY: number): void {
  const { doc } = ctx;
  for (const frag of line.fragments) {
    // Equation fragments paint the math box regardless of their (sentinel) text —
    // check before the empty-text skip. Mirror the run path so a styled or
    // hyperlinked inline equation keeps its highlight, color, rules, and link.
    if (frag.equation) {
      const s = frag.style;
      const x = block.x + frag.x;
      if (s.highlightColor) {
        doc.rect(x, block.y + line.y, frag.width, line.height).fill(s.highlightColor as string);
      }
      const rp = runPaint(s);
      drawMathBoxPdf(ctx, frag.equation, x, baselineY, rp.color);
      const th = decorationThickness(s.fontSizePx);
      if (rp.underline) paintUnderlinePdf(doc, x, baselineY + UNDERLINE_OFFSET_PX, frag.width, s.fontSizePx, rp);
      if (rp.strike) doc.rect(x, baselineY + strikeOffset(s.fontSizePx), frag.width, th).fill(rp.color);
      if (rp.externalLink && s.link) {
        doc.link(x, block.y + line.y, frag.width, line.height, s.link);
      } else if (s.link && s.link.startsWith("#")) {
        const name = ctx.destName?.({ anchor: s.link.slice(1) });
        if (name) doc.goTo(x, block.y + line.y, frag.width, line.height, name);
      }
      continue;
    }
    if (frag.text.length === 0) continue;
    const s = frag.style;
    const x = block.x + frag.x;
    // sub/super shift + explicit w:position raise/lower, folded together.
    const vShift = runVerticalShift(s);
    // Sub/super are measured (and so must paint) at the scaled size.
    const sizePx = s.verticalAlign ? Math.round(s.fontSizePx * SUB_SUPER_SCALE) : s.fontSizePx;

    if (s.highlightColor) {
      doc.rect(x, block.y + line.y, frag.width, line.height).fill(s.highlightColor as string);
    }

    const rp = runPaint(s);
    // Character width scaling (w:w): the fragment box already reserves the scaled
    // advance, so stretch the glyphs horizontally about the fragment's left edge.
    const wScale = widthScale(s);

    // Per-glyph fallback: split the fragment text into face-homogeneous segments
    // so characters the primary face lacks are rendered from a fallback face
    // (StixTwoMath covers ✓ U+2713, ☒ U+2612, etc.). Both measure (fontkitContext)
    // and paint use segmentByFace so reserved widths match painted positions.
    if (ctx.textSegments) {
      const segs = ctx.textSegments(frag.text, firstFamily(s.fontFamily), !!s.bold, !!s.italic);
      // Apply wScale once around all segments — the scale origin is the fragment's
      // left edge (x) so every segment's logical advance maps to the right visual x.
      if (wScale !== 1) doc.save().scale(wScale, 1, { origin: [x, baselineY + vShift] });
      const wordSpacingPx = frag.wordSpacingPx ?? 0;
      let curX = x;
      for (const seg of segs) {
        doc
          .font(seg.pdfkitName)
          .fontSize(sizePx)
          .fillColor(rp.color)
          .text(seg.text, curX, baselineY + vShift, {
            lineBreak: false,
            baseline: "alphabetic",
            wordSpacing: wordSpacingPx,
          });
        // Advance by the glyph run width PLUS pdfkit's per-space wordSpacing, which
        // it adds to every U+0020 — otherwise a segment containing spaces leaves the
        // next (fallback) segment shifted too far left.
        const spaceCount = wordSpacingPx ? (seg.text.match(/ /g)?.length ?? 0) : 0;
        curX +=
          seg.font.layout(seg.text).advanceWidth * (sizePx / seg.font.unitsPerEm) +
          spaceCount * wordSpacingPx;
      }
      if (wScale !== 1) doc.restore();
    } else {
      // Fallback for minimal PaintCtx (tests that don't supply textSegments).
      const name = ctx.font(firstFamily(s.fontFamily), !!s.bold, !!s.italic);
      if (wScale !== 1) doc.save().scale(wScale, 1, { origin: [x, baselineY + vShift] });
      doc
        .font(name)
        .fontSize(sizePx)
        .fillColor(rp.color)
        .text(frag.text, x, baselineY + vShift, {
          lineBreak: false,
          baseline: "alphabetic",
          wordSpacing: frag.wordSpacingPx ?? 0,
        });
      if (wScale !== 1) doc.restore();
    }

    const th = decorationThickness(s.fontSizePx);
    if (rp.underline) {
      paintUnderlinePdf(doc, x, baselineY + vShift + UNDERLINE_OFFSET_PX, frag.width, s.fontSizePx, rp);
    }
    if (rp.strike) {
      doc.rect(x, baselineY + vShift + strikeOffset(s.fontSizePx), frag.width, th).fill(rp.color);
    }
    // Double strikethrough (w:dstrike): two rules straddling the single-strike line.
    if (s.doubleStrikethrough) {
      for (const off of doubleStrikeOffsets(s.fontSizePx)) {
        doc.rect(x, baselineY + vShift + off, frag.width, th).fill(rp.color);
      }
    }
    if (rp.externalLink && s.link) {
      // Clickable annotation over the run's line box.
      doc.link(x, block.y + line.y, frag.width, line.height, s.link);
    } else if (s.link && s.link.startsWith("#")) {
      // In-document anchor (cross-reference / hyperlinked TOC entry): GoTo the
      // resolved heading instead of an external URI.
      const name = ctx.destName?.({ anchor: s.link.slice(1) });
      if (name) doc.goTo(x, block.y + line.y, frag.width, line.height, name);
    }
  }
}

function placeholderBox(ctx: PaintCtx, x: number, y: number, w: number, h: number): void {
  ctx.doc.rect(x, y, w, h).fill(IMAGE_PLACEHOLDER_COLOR);
}

/** Trace a preset geometry into the pdfkit path, inside the box (x,y,w,h) — the PDF
 *  mirror of the canvas traceShapePath (same proportions, so screen and export
 *  match). pdfkit consumes the path on fill/stroke, so it is retraced for each. */
function shapePathPdf(doc: PDFKit.PDFDocument, preset: PlacedShape["preset"], x: number, y: number, w: number, h: number): void {
  const poly = (pts: [number, number][]): void => {
    doc.moveTo(pts[0]![0], pts[0]![1]);
    for (let i = 1; i < pts.length; i++) doc.lineTo(pts[i]![0], pts[i]![1]);
    doc.closePath();
  };
  switch (preset) {
    case "ellipse": doc.ellipse(x + w / 2, y + h / 2, w / 2, h / 2); return;
    case "line": doc.moveTo(x, y).lineTo(x + w, y + h); return;
    case "roundRect": doc.roundedRect(x, y, w, h, Math.min(w, h) * 0.18); return;
    case "triangle": poly([[x + w / 2, y], [x + w, y + h], [x, y + h]]); return;
    case "diamond": poly([[x + w / 2, y], [x + w, y + h / 2], [x + w / 2, y + h], [x, y + h / 2]]); return;
    case "rightArrow": {
      const head = Math.min(0.4 * w, w);
      const t = y + h * 0.25, b = y + h * 0.75;
      poly([[x, t], [x + w - head, t], [x + w - head, y], [x + w, y + h / 2], [x + w - head, y + h], [x + w - head, b], [x, b]]);
      return;
    }
    case "leftArrow": {
      const head = Math.min(0.4 * w, w);
      const t = y + h * 0.25, b = y + h * 0.75;
      poly([[x + w, t], [x + head, t], [x + head, y], [x, y + h / 2], [x + head, y + h], [x + head, b], [x + w, b]]);
      return;
    }
    default: doc.rect(x, y, w, h);
  }
}

/** Trace a freeform custom geometry (ShapePath) into the pdfkit path, scaling the
 *  normalized (0–1) segment coords into the box (x,y,w,h) — the PDF mirror of the
 *  canvas traceCustomPath (PR 7, issue #220). Retraced per fill/stroke like the
 *  preset path helper. */
function customPathPdf(doc: PDFKit.PDFDocument, path: PlacedShape["custom"] & object, x: number, y: number, w: number, h: number): void {
  for (const seg of path.segments) {
    switch (seg.type) {
      case "moveTo": doc.moveTo(x + seg.x * w, y + seg.y * h); break;
      case "lineTo": doc.lineTo(x + seg.x * w, y + seg.y * h); break;
      case "cubicBezierTo": doc.bezierCurveTo(x + seg.x1 * w, y + seg.y1 * h, x + seg.x2 * w, y + seg.y2 * h, x + seg.x * w, y + seg.y * h); break;
      case "close": doc.closePath(); break;
    }
  }
}

/** Draw a placed drawing shape's geometry into its box at (x,y) — the PDF mirror of
 *  the canvas paintShapeCanvas: fill, then the (optionally dashed) stroke, about the
 *  box center when rotated. A custom freeform path is traced when present; otherwise
 *  the preset (line is the box diagonal, OOXML prst="line"). */
function paintShapePdf(doc: PDFKit.PDFDocument, shape: PlacedShape, x: number, y: number): void {
  const { fill, stroke } = resolveShapePaint(shape);
  const w = shape.width;
  const h = shape.height;
  const trace = (): void => {
    if (shape.custom) customPathPdf(doc, shape.custom, x, y, w, h);
    else shapePathPdf(doc, shape.preset, x, y, w, h);
  };
  doc.save();
  if (shape.rotation) doc.rotate(shape.rotation, { origin: [x + w / 2, y + h / 2] });
  if (fill && (shape.custom || shape.preset !== "line")) {
    trace();
    doc.fill(fill);
  }
  if (stroke) {
    trace();
    doc.lineWidth(stroke.widthPx);
    if (stroke.dash.length === 2) doc.dash(stroke.dash[0]!, { space: stroke.dash[1]! });
    else if (stroke.dash.length > 2) doc.addContent(`[${stroke.dash.map((d) => Number(d.toFixed(2))).join(" ")}] 0 d`);
    doc.stroke(stroke.color);
    doc.undash();
  }
  doc.restore();
}

/** Paint a placed drawing shape at (x,y): a GROUP container (PlacedShape with
 *  `children`) recurses into its members (each already offset into the box); a leaf
 *  paints its preset geometry and its read-only text box body (clipped + translated
 *  into the local text frame). Recursive so nested groups compose. Mirrors the
 *  canvas painter (paint/renderer.ts). */
function paintPlacedShapePdf(ctx: PaintCtx, shape: PlacedShape, x: number, y: number): void {
  const { doc } = ctx;
  if (shape.children) {
    for (const c of shape.children) paintPlacedShapePdf(ctx, c.shape, x + c.x, y + c.y);
    return;
  }
  paintShapePdf(doc, shape, x, y);
  // Read-only text box body (mirrors canvas). KNOWN LIMITATION (same as canvas): the
  // text is painted axis-aligned even when the shape is rotated — only the geometry rotates.
  const text = shape.text;
  if (text && text.blocks.length > 0) {
    doc.save();
    doc.rect(x, y, shape.width, shape.height).clip();
    doc.translate(x + text.offsetX, y + text.offsetY);
    for (const cb of text.blocks) paintBlock(ctx, cb);
    doc.restore();
  }
}

/** Draw a default-bullet marker as a vector shape (disc / ring / square), the PDF
 *  counterpart of the canvas renderer's paintBulletShape — both consume the same
 *  {@link BulletShape} so a bullet looks identical on screen and in the export. */
function drawBulletShape(doc: PDFKit.PDFDocument, shape: BulletShape, color: string): void {
  if (shape.kind === "square") {
    doc.rect(shape.x, shape.y, shape.size, shape.size).fill(color);
  } else if (shape.kind === "ring") {
    doc.save().lineWidth(shape.lineWidth).circle(shape.cx, shape.cy, shape.r).stroke(color).restore();
  } else {
    doc.circle(shape.cx, shape.cy, shape.r).fill(color);
  }
}

interface OpenableImageDoc {
  openImage(src: Buffer): { colorSpace?: string; width: number; height: number };
}

/** Draw an image, correcting pdfkit's JPEG colorspace bug: its parser reads the
 *  first component's *id* (1) instead of the component *count* (3/4), tagging
 *  every standard color JPEG as DeviceGray (renders grayed/garbled). We parse the
 *  real count and override colorSpace on the opened image before placing it. PNGs
 *  go straight through pdfkit. */
function drawImage(doc: PDFKit.PDFDocument, buf: Buffer, x: number, y: number, w: number, h: number): void {
  const cs = jpegColorSpace(buf);
  if (cs) {
    const img = (doc as unknown as OpenableImageDoc).openImage(buf);
    img.colorSpace = cs;
    doc.image(img as unknown as Buffer, x, y, { width: w, height: h });
  } else {
    doc.image(buf, x, y, { width: w, height: h });
  }
}

/** Real JPEG colorspace from the SOF component count (DeviceGray/RGB/CMYK), or
 *  undefined if not a JPEG (let pdfkit handle PNG). */
function jpegColorSpace(b: Uint8Array): string | undefined {
  if (b.length < 4 || b[0] !== 0xff || b[1] !== 0xd8) return undefined;
  let pos = 2;
  while (pos < b.length - 8) {
    if (b[pos] !== 0xff) {
      pos++;
      continue;
    }
    const low = b[pos + 1]!;
    pos += 2;
    // SOF0..SOF15 (start of frame) except DHT(C4)/JPG(C8)/DAC(CC).
    if (low >= 0xc0 && low <= 0xcf && low !== 0xc4 && low !== 0xc8 && low !== 0xcc) {
      const nf = b[pos + 7]!; // Lf(2)+P(1)+Y(2)+X(2) → Nf
      return nf === 1 ? "DeviceGray" : nf === 4 ? "DeviceCMYK" : "DeviceRGB";
    }
    pos += (b[pos]! << 8) | b[pos + 1]!; // skip segment by length
  }
  return undefined;
}

function paintCellBorders(ctx: PaintCtx, cell: PlacedTableCell): void {
  const { doc } = ctx;
  const { x, y, width: w, height: h } = cell;
  if (cell.borders === undefined) {
    doc.undash().lineWidth(1).strokeColor(DEFAULT_GRID_COLOR);
    doc.rect(x + 0.5, y + 0.5, w, h).stroke();
    return;
  }
  const b = cell.borders;
  strokeCellEdge(ctx, b.top, x, y, x + w, y, 0, 1);
  strokeCellEdge(ctx, b.bottom, x, y + h, x + w, y + h, 0, -1);
  strokeCellEdge(ctx, b.left, x, y, x, y + h, 1, 0);
  strokeCellEdge(ctx, b.right, x + w, y, x + w, y + h, -1, 0);
}

function strokeCellEdge(
  ctx: PaintCtx,
  spec: CellBorder | undefined,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  ix: number,
  iy: number,
): void {
  if (!spec) return;
  const { doc } = ctx;
  const w = cellBorderWidth(spec.widthPx);
  doc.lineWidth(w).strokeColor(spec.color as string);
  const dash = cellBorderDash(spec.style, w);
  if (dash.length) doc.dash(dash[0]!, { space: dash[1]! });
  else doc.undash();
  const horizontal = y1 === y2;
  const off = Math.round(w) % 2 ? 0.5 : 0;
  const ox = horizontal ? 0 : off;
  const oy = horizontal ? off : 0;
  const line = (dx: number, dy: number): void => {
    doc.moveTo(x1 + ox + dx, y1 + oy + dy).lineTo(x2 + ox + dx, y2 + oy + dy).stroke();
  };
  line(0, 0);
  if (spec.style === "double") {
    const g = doubleBorderGap(w);
    line(ix * g, iy * g);
  }
  doc.undash();
}

function toBuffer(bytes: Uint8Array): Buffer {
  return Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}
