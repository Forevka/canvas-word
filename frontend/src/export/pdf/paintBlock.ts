// PlacedBlock -> pdfkit, a constant-for-constant inverse of src/paint/renderer.ts.
// Same baseline formula (block.y + line.y + line.ascent), same sub/super shifts,
// underline/strike offsets, default grid color, footnote-rule width and leader
// dash patterns — so a page in the PDF matches the canvas pixel-for-pixel (modulo
// metric-clone glyph shapes). pdfkit shares y-down, top-left coords with canvas.

import type { CellBorder } from "@cw/shared";
import type { LineBox, PlacedBlock, PlacedTableCell } from "../../layout/layoutTree";
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
  runPaint,
  strikeOffset,
  SUB_SUPER_SCALE,
  TOC_LEADER_COLOR,
  TOC_LEADER_DASH,
  TOC_LEADER_GAP_PX,
  UNDERLINE_OFFSET_PX,
  verticalShift,
} from "../../paint/paintStyle";

export interface PaintCtx {
  doc: PDFKit.PDFDocument;
  /** Register (idempotently) the bundled face and return its pdfkit name. */
  font: (family: string, bold: boolean, italic: boolean) => string;
  /** Image bytes for a src, or undefined if unresolved/unsupported. */
  image: (src: string) => Uint8Array | undefined;
  warn: (code: string, detail?: string) => void;
  /** Resolve a TOC target blockId or an in-document "#anchor" to a registered
   *  named-destination string, or undefined when the target doesn't exist. The
   *  caller emits a GoTo annotation only when a name comes back. */
  destName?: (ref: { blockId?: string; anchor?: string }) => string | undefined;
}

const firstFamily = (stack: string): string => (stack.split(",")[0] ?? "sans-serif").trim();

export function paintBlock(ctx: PaintCtx, block: PlacedBlock): void {
  const { doc } = ctx;

  if (block.image) {
    const bytes = ctx.image(block.image.src);
    const { width, height, clip } = block.image;
    if (clip) {
      doc.save();
      doc.rect(clip.x, clip.y, clip.width, clip.height).clip();
    }
    if (bytes) {
      try {
        drawImage(doc, toBuffer(bytes), block.x, block.y, width, height);
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
        for (const cb of cell.blocks) paintBlock(ctx, cb);
      }
    }
    for (const row of rows) {
      for (const cell of row.cells) paintCellBorders(ctx, cell);
    }
    return;
  }

  // List marker — paint-only, first line's baseline in the hanging indent.
  const firstLine = block.lines[0];
  if (block.marker && firstLine) {
    const s = block.marker.style;
    const name = ctx.font(firstFamily(s.fontFamily), !!s.bold, !!s.italic);
    doc
      .font(name)
      .fontSize(s.fontSizePx)
      .fillColor(s.color as string)
      .text(block.marker.text, block.marker.x, block.y + firstLine.y + firstLine.ascent, {
        lineBreak: false,
        baseline: "alphabetic",
      });
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
}

function paintLine(ctx: PaintCtx, block: PlacedBlock, line: LineBox, baselineY: number): void {
  const { doc } = ctx;
  for (const frag of line.fragments) {
    if (frag.text.length === 0) continue;
    const s = frag.style;
    const x = block.x + frag.x;
    const vShift = verticalShift(s.verticalAlign, s.fontSizePx);
    // Sub/super are measured (and so must paint) at the scaled size.
    const sizePx = s.verticalAlign ? Math.round(s.fontSizePx * SUB_SUPER_SCALE) : s.fontSizePx;

    if (s.highlightColor) {
      doc.rect(x, block.y + line.y, frag.width, line.height).fill(s.highlightColor as string);
    }

    const rp = runPaint(s);
    const name = ctx.font(firstFamily(s.fontFamily), !!s.bold, !!s.italic);
    doc
      .font(name)
      .fontSize(sizePx)
      .fillColor(rp.color)
      .text(frag.text, x, baselineY + vShift, {
        lineBreak: false,
        baseline: "alphabetic",
        wordSpacing: frag.wordSpacingPx ?? 0,
      });

    const th = decorationThickness(s.fontSizePx);
    if (rp.underline) {
      doc.rect(x, baselineY + vShift + UNDERLINE_OFFSET_PX, frag.width, th).fill(rp.color);
    }
    if (rp.strike) {
      doc.rect(x, baselineY + vShift + strikeOffset(s.fontSizePx), frag.width, th).fill(rp.color);
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
