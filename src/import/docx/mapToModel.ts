// IR → canvas-word model. This is the lossy stage: everything the model
// can't hold is decided HERE (with a warning), never silently in the parser.
//
// Policies (see IMPORT.md "Lossy mappings"):
//   w:br soft break  → paragraph split (split halves hug: zero space between)
//   w:tab            → fixed spaces
//   gridSpan/vMerge  → padded uniform cell grid
//   inline images    → block-level ImageBlock (paragraph splits around it)
//   images in cells  → skipped (TableCell holds paragraphs only)

import type {
  Block,
  CharStyle,
  Document,
  ImageBlock,
  ParaStyle,
  Paragraph,
  Run,
  SectionProps,
  TableBlock,
  TableCell,
} from "../../model/document";
import { normalizeRuns } from "../../model/ops";
import type { MediaStore } from "./media";
import type { StyleResolver } from "./styles";
import type { IRBlock, IRInline, IRParaProps, IRParagraph, IRRunProps, IRSection, IRTable } from "./types";
import { WarningSink } from "./types";
import { emuToPx, halfPointsToPx, round2, twipsToPx } from "./units";

// Fallbacks for properties NOTHING specifies (no docDefaults, no style, no
// direct formatting). These mirror what Word itself renders for a bare
// document — single line spacing, no space after, Times New Roman 12pt —
// NOT the editor's native defaults: an imported page must look like it did
// in Word, even where the file is silent.
const DEFAULT_CHAR: CharStyle = {
  fontFamily: "Times New Roman, serif",
  fontSizePx: 16, // 12pt
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#000000",
};

const DEFAULT_PARA: ParaStyle = {
  align: "left",
  lineHeight: 1,
  spaceBeforePx: 0,
  spaceAfterPx: 0,
  indentFirstLinePx: 0,
  indentLeftPx: 0,
};

/** US Letter at 96dpi, 1in margins — matches sampleDoc. */
const DEFAULT_SECTION: SectionProps = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
};

const TAB_AS_SPACES = "    ";

export interface Mapper {
  /** Map a block list (document body, header story, footer story). */
  mapBlocks(blocks: IRBlock[], media: MediaStore): Block[];
  mapSection(section: IRSection | null): SectionProps;
  /** The editor assumes at least one block (the caret needs a home). */
  emptyParagraph(): Paragraph;
}

export function createMapper(warnings: WarningSink, resolver: StyleResolver): Mapper {
  // Import-distinct id prefix: commands.ts mints `n…`, sampleDoc `b…`.
  let nextId = 0;
  const id = (): string => `i${nextId++}`;

  // What a run with no formatting at all resolves to in THIS document
  // (docDefaults + default paragraph style) — empty paragraphs and padded
  // table cells must carry the document's defaults, not the editor's.
  const documentChar = mapCharStyle(resolver.run(undefined, {}));
  const documentPara = mapParaStyle(resolver.para({}));

  const emptyParagraph = (): Paragraph => ({
    kind: "paragraph",
    id: id(),
    revision: 0,
    runs: [{ text: "", style: { ...documentChar } }],
    style: { ...documentPara },
  });

  const mapBlocks = (blocks: IRBlock[], media: MediaStore): Block[] => {
    const out: Block[] = [];
    // A paragraph carrying w:sectPr ends a section; unless the break is
    // "continuous", whatever follows starts a new page.
    let sectionPageBreak = false;
    for (const irBlock of blocks) {
      const mapped = irBlock.kind === "paragraph" ? mapParagraph(irBlock, media) : [mapTable(irBlock, media)];
      if (sectionPageBreak && mapped.length > 0) {
        const first = mapped[0]!;
        if (first.kind === "paragraph") first.style.pageBreakBefore = true;
        else mapped.unshift({ ...emptyParagraph(), style: { ...documentPara, pageBreakBefore: true } });
        sectionPageBreak = false;
      }
      out.push(...mapped);
      if (irBlock.kind === "paragraph" && irBlock.props.sectionBreak === "page") sectionPageBreak = true;
    }
    return out;
  };

  // -------------------------------------------------------------------------
  // Paragraphs

  /** One IR paragraph maps to 1..N model blocks: soft breaks split it (the
   *  model has no intra-paragraph line break), and inline images surface as
   *  block-level ImageBlocks between the text fragments. */
  function mapParagraph(ir: IRParagraph, media: MediaStore): Block[] {
    const style = mapParaStyle(resolver.para(ir.props));
    // Empty paragraphs take the paragraph MARK's formatting (w:pPr/w:rPr over
    // the style cascade) — that's what sizes the empty line in Word.
    const markChar = mapCharStyle(resolver.run(ir.props.styleId, ir.props.markRunProps ?? {}));
    const blocks: Block[] = [];
    let runs: Run[] = [];
    let trailingBreak = false;
    // Set by a page break; consumed by the NEXT emitted paragraph — what
    // follows the break starts a new page.
    let pendingPageBreak = false;

    const paraOf = (paraRuns: Run[]): Paragraph => {
      const paraStyle = { ...style };
      if (pendingPageBreak) {
        paraStyle.pageBreakBefore = true;
        pendingPageBreak = false;
      }
      return {
        kind: "paragraph",
        id: id(),
        revision: 0,
        runs: normalizeRuns(paraRuns, markChar),
        style: paraStyle,
      };
    };
    const flushPara = (): void => {
      if (runs.length === 0) return;
      blocks.push(paraOf(runs));
      runs = [];
    };

    for (const inline of ir.inlines) {
      switch (inline.kind) {
        case "break":
          if (!inline.page) {
            warnings.add("soft-breaks", "Soft line breaks (Shift+Enter) became paragraph breaks.");
          }
          // A break with nothing pending is an empty visual line ("a\n\nb").
          if (runs.length > 0) flushPara();
          else blocks.push(paraOf([]));
          if (inline.page) pendingPageBreak = true;
          trailingBreak = true;
          break;
        case "image": {
          const image = mapImage(inline, media, style.align);
          if (image) {
            flushPara();
            // An image can't carry pageBreakBefore — give the break a carrier.
            if (pendingPageBreak) blocks.push(paraOf([]));
            blocks.push(image);
            trailingBreak = false;
          }
          break;
        }
        case "run": {
          const effective = resolver.run(ir.props.styleId, inline.props);
          if (effective.vanish) {
            // Hidden text is invisible in Word's normal view — keeping it would
            // surface text the author deliberately suppressed. (vanish can
            // arrive via the style cascade too, hence the post-resolution check.)
            warnings.add("hidden-text", "Hidden text (w:vanish) was dropped.");
            break;
          }
          runs.push(mapRun(inline.text, effective));
          trailingBreak = false;
          break;
        }
      }
    }
    flushPara();

    // Trailing soft break = an empty line after it. A paragraph with no
    // surviving content still occupies vertical space — but an image-only
    // paragraph shouldn't add a stray empty line.
    if (trailingBreak || blocks.length === 0) blocks.push(paraOf([]));

    // Spacing belongs to the ORIGINAL paragraph: first split keeps the space
    // before, last keeps the space after, seams between splits hug.
    const paragraphs = blocks.filter((b): b is Paragraph => b.kind === "paragraph");
    paragraphs.forEach((p, i) => {
      if (i > 0) p.style.spaceBeforePx = 0;
      if (i < paragraphs.length - 1) p.style.spaceAfterPx = 0;
    });
    return blocks;
  }

  function mapImage(
    inline: Extract<IRInline, { kind: "image" }>,
    media: MediaStore,
    paraAlign: ParaStyle["align"],
  ): ImageBlock | undefined {
    const src = media.resolve(inline.relId);
    if (!src) return undefined; // media store already warned
    if (inline.widthEmu === undefined || inline.heightEmu === undefined) {
      warnings.add("images-unsized", "An image without explicit dimensions was skipped.");
      return undefined;
    }
    const image: ImageBlock = {
      kind: "image",
      id: id(),
      revision: 0,
      src,
      widthPx: round2(emuToPx(inline.widthEmu)),
      heightPx: round2(emuToPx(inline.heightEmu)),
      align: inline.anchorAlign ?? (paraAlign === "justify" ? "left" : paraAlign),
    };
    // Anchored with square/tight wrap → an honest float; the model flows
    // following text around it.
    if (inline.anchored && inline.anchorWrap === "square") image.wrap = "square";
    return image;
  }

  function mapRun(rawText: string, effective: IRRunProps): Run {
    let text = rawText;
    if (text.includes("\t")) {
      warnings.add("tabs", "Tab stops became fixed spaces (no tab-stop layout).");
      text = text.replaceAll("\t", TAB_AS_SPACES);
    }
    return { text, style: mapCharStyle(effective) };
  }

  // -------------------------------------------------------------------------
  // Tables — cells are full block stories; gridSpan maps to colSpan.

  function mapTable(ir: IRTable, media: MediaStore): TableBlock {
    const emptyCell = (): TableCell => ({ id: id(), blocks: [emptyParagraph()] });
    const spanSum = (cells: TableCell[]): number => cells.reduce((s, c) => s + (c.colSpan ?? 1), 0);

    const rows = ir.rows.map((irRow) => {
      const cells: TableCell[] = [];
      for (const irCell of irRow.cells) {
        if (irCell.vMergeContinue) {
          // No vertical spans in the model — the continuation stays as its own
          // (typically empty) cell.
          warnings.add("cell-vmerge", "Vertically merged cells were split (no row spans).");
        }
        const blocks: Block[] = [];
        for (const b of irCell.blocks) {
          if (b.kind === "paragraph") blocks.push(...mapParagraph(b, media));
          else blocks.push(mapTable(b, media)); // nested table — model renders one level
        }
        const cell: TableCell = { id: id(), blocks: blocks.length > 0 ? blocks : [emptyParagraph()] };
        if (irCell.gridSpan > 1) cell.colSpan = irCell.gridSpan;
        cells.push(cell);
      }
      return { cells };
    });

    // Keep every row's span total equal (ragged rows exist in real files).
    const width = Math.max(1, ...rows.map((r) => spanSum(r.cells)));
    for (const row of rows) {
      for (let w = spanSum(row.cells); w < width; w++) row.cells.push(emptyCell());
    }

    const table: TableBlock = { kind: "table", id: id(), revision: 0, rows };
    // w:tblGrid column widths → fractions of content width (when consistent).
    if (ir.colWidthsTwips && ir.colWidthsTwips.length === width) {
      const total = ir.colWidthsTwips.reduce((s, w) => s + w, 0);
      if (total > 0) {
        const fractions = ir.colWidthsTwips.map((w) => Math.round((w / total) * 10000) / 10000);
        fractions[fractions.length - 1] =
          Math.round((1 - fractions.slice(0, -1).reduce((s, f) => s + f, 0)) * 10000) / 10000;
        table.colFractions = fractions;
      }
    }
    return table;
  }

  // -------------------------------------------------------------------------
  // Section

  function mapSection(ir: IRSection | null): SectionProps {
    if (!ir) return { ...DEFAULT_SECTION, marginPx: { ...DEFAULT_SECTION.marginPx } };
    return {
      pageWidthPx:
        ir.pageWidthTwips !== undefined ? round2(twipsToPx(ir.pageWidthTwips)) : DEFAULT_SECTION.pageWidthPx,
      pageHeightPx:
        ir.pageHeightTwips !== undefined ? round2(twipsToPx(ir.pageHeightTwips)) : DEFAULT_SECTION.pageHeightPx,
      marginPx: ir.marginTwips
        ? {
            top: round2(twipsToPx(ir.marginTwips.top)),
            right: round2(twipsToPx(ir.marginTwips.right)),
            bottom: round2(twipsToPx(ir.marginTwips.bottom)),
            left: round2(twipsToPx(ir.marginTwips.left)),
          }
        : { ...DEFAULT_SECTION.marginPx },
    };
  }

  return { mapBlocks, mapSection, emptyParagraph };
}

// ---------------------------------------------------------------------------
// Style mapping (pure)

function mapCharStyle(props: IRRunProps): CharStyle {
  const style: CharStyle = { ...DEFAULT_CHAR };
  if (props.fontAscii) style.fontFamily = `${props.fontAscii}, serif`;
  if (props.sizeHalfPoints !== undefined) style.fontSizePx = round2(halfPointsToPx(props.sizeHalfPoints));
  if (props.bold !== undefined) style.bold = props.bold;
  if (props.italic !== undefined) style.italic = props.italic;
  if (props.underline !== undefined) style.underline = props.underline;
  if (props.strikethrough !== undefined) style.strikethrough = props.strikethrough;
  if (props.color && props.color !== "auto") style.color = `#${props.color.toLowerCase()}`;
  return style;
}

function mapParaStyle(props: IRParaProps): ParaStyle {
  const style: ParaStyle = { ...DEFAULT_PARA };
  if (props.align) style.align = props.align;
  if (props.lineHeight !== undefined) style.lineHeight = round2(props.lineHeight);
  if (props.spaceBeforeTwips !== undefined) style.spaceBeforePx = round2(twipsToPx(props.spaceBeforeTwips));
  if (props.spaceAfterTwips !== undefined) style.spaceAfterPx = round2(twipsToPx(props.spaceAfterTwips));
  if (props.indentLeftTwips !== undefined) style.indentLeftPx = round2(twipsToPx(props.indentLeftTwips));
  if (props.indentFirstLineTwips !== undefined)
    style.indentFirstLinePx = round2(twipsToPx(props.indentFirstLineTwips));
  if (props.keepWithNext) style.keepWithNext = true;
  if (props.pageBreakBefore) style.pageBreakBefore = true;
  if (props.styleId) style.namedStyle = props.styleId;
  return style;
}
