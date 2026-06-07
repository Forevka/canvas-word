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
  SdtProps,
  SectionProps,
  TableBlock,
  TableCell,
} from "../../model/document";
import type { NamedStyle, Stylesheet } from "../../model/stylesheet";
import { normalizeRuns } from "../../model/ops";
import type { MediaStore } from "./media";
import type { StyleResolver, StylesData } from "./styles";
import type { IRBlock, IRInline, IRParaProps, IRParagraph, IRRunProps, IRSdtProps, IRSection, IRTable } from "./types";
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

export function createMapper(
  warnings: WarningSink,
  resolver: StyleResolver,
  sdts: Record<string, IRSdtProps> = {},
): Mapper {
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
          runs.push(mapRun(inline.text, effective, inline.sdtId));
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

  function mapRun(rawText: string, effective: IRRunProps, sdtId?: string): Run {
    let text = rawText;
    if (text.includes("\t")) {
      warnings.add("tabs", "Tab stops became fixed spaces (no tab-stop layout).");
      text = text.replaceAll("\t", TAB_AS_SPACES);
    }
    const style = mapCharStyle(effective);
    if (sdtId) {
      style.sdtId = sdtId;
      const sdt = sdts[sdtId];
      // Checkbox glyphs arrive in symbol fonts (MS Gothic / Wingdings private
      // chars) — normalize to the Unicode glyphs the editor toggles between.
      if (sdt?.type === "checkbox" && text.length > 0) {
        text = sdt.checked ? "☒" : "☐";
        style.fontFamily = DEFAULT_CHAR.fontFamily;
      }
    }
    return { text, style };
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

/** IR content controls → model SdtProps (shapes match; copy defined fields). */
export function mapSdts(ir: Record<string, IRSdtProps>): Record<string, SdtProps> {
  const out: Record<string, SdtProps> = {};
  for (const [id, p] of Object.entries(ir)) {
    const props: SdtProps = { type: p.type };
    if (p.alias !== undefined) props.alias = p.alias;
    if (p.tag !== undefined) props.tag = p.tag;
    if (p.placeholder) props.placeholder = true;
    if (p.listItems !== undefined) props.listItems = p.listItems;
    if (p.dateFormat !== undefined) props.dateFormat = p.dateFormat;
    if (p.checked !== undefined) props.checked = p.checked;
    if (p.lockContent) props.lockContent = true;
    if (p.lockControl) props.lockControl = true;
    out[id] = props;
  }
  return out;
}

/** Paragraph styleIds referenced anywhere in the IR (incl. table cells). */
export function collectUsedStyleIds(blocks: IRBlock[], into: Set<string> = new Set()): Set<string> {
  for (const b of blocks) {
    if (b.kind === "paragraph") {
      if (b.props.styleId) into.add(b.props.styleId);
    } else {
      for (const row of b.rows) for (const cell of row.cells) collectUsedStyleIds(cell.blocks, into);
    }
  }
  return into;
}

/** Patch-only style mappers: a NamedStyle carries each style's OWN deltas —
 *  the editor resolves basedOn chains itself. Theme-indirected fields are
 *  skipped (the resolver already baked them into the runs). */
function mapCharPatch(props: IRRunProps): Partial<CharStyle> {
  const out: Partial<CharStyle> = {};
  if (props.fontAscii) out.fontFamily = `${props.fontAscii}, serif`;
  if (props.sizeHalfPoints !== undefined) out.fontSizePx = round2(halfPointsToPx(props.sizeHalfPoints));
  if (props.bold !== undefined) out.bold = props.bold;
  if (props.italic !== undefined) out.italic = props.italic;
  if (props.underline !== undefined) out.underline = props.underline;
  if (props.strikethrough !== undefined) out.strikethrough = props.strikethrough;
  if (props.color && props.color !== "auto") out.color = `#${props.color.toLowerCase()}`;
  return out;
}

function mapParaPatch(props: IRParaProps): Partial<ParaStyle> {
  const out: Partial<ParaStyle> = {};
  if (props.align) out.align = props.align;
  if (props.lineHeight !== undefined) out.lineHeight = round2(props.lineHeight);
  if (props.spaceBeforeTwips !== undefined) out.spaceBeforePx = round2(twipsToPx(props.spaceBeforeTwips));
  if (props.spaceAfterTwips !== undefined) out.spaceAfterPx = round2(twipsToPx(props.spaceAfterTwips));
  if (props.indentLeftTwips !== undefined) out.indentLeftPx = round2(twipsToPx(props.indentLeftTwips));
  if (props.indentFirstLineTwips !== undefined)
    out.indentFirstLinePx = round2(twipsToPx(props.indentFirstLineTwips));
  if (props.keepWithNext) out.keepWithNext = true;
  return out;
}

/** styles.xml → editor Stylesheet, restricted to paragraph styles the document
 *  actually USES (plus their basedOn closure and the default style) — a
 *  generated report carries hundreds of unused styles. Display names come from
 *  w:name; generated documents use opaque numeric styleIds. */
export function buildStylesheet(data: StylesData, usedIds: Set<string>): Stylesheet | undefined {
  if (data.styles.size === 0) return undefined;
  const include = new Set<string>();
  const addWithBases = (id: string): void => {
    for (let cur = data.styles.get(id); cur && !include.has(cur.id); cur = cur.basedOnId ? data.styles.get(cur.basedOnId) : undefined) {
      if (cur.type === "paragraph") include.add(cur.id);
    }
  };
  for (const id of usedIds) addWithBases(id);
  if (data.defaultParaStyleId) addWithBases(data.defaultParaStyleId);
  if (include.size === 0) return undefined;

  const styles: NamedStyle[] = [];
  for (const id of include) {
    const def = data.styles.get(id)!;
    const style: NamedStyle = {
      id,
      name: def.name ?? id,
      char: mapCharPatch(def.rPr),
      para: mapParaPatch(def.pPr),
    };
    if (def.basedOnId && include.has(def.basedOnId)) style.basedOn = def.basedOnId;
    styles.push(style);
  }
  // Default style first — the gallery reads top-to-bottom.
  styles.sort((a, b) =>
    a.id === data.defaultParaStyleId ? -1 : b.id === data.defaultParaStyleId ? 1 : a.name.localeCompare(b.name),
  );
  return { styles, defaultStyleId: data.defaultParaStyleId ?? styles[0]!.id };
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
