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
import type { ListDefinition, ListLevel, ListNumberFormat } from "../../model/lists";
import { normalizeRuns } from "../../model/ops";
import type { MediaStore } from "./media";
import type { NumberingData } from "./numbering";
import type { StyleResolver, StylesData } from "./styles";
import type {
  IRBlock,
  IRInline,
  IRListDefinition,
  IRParaProps,
  IRParagraph,
  IRRunProps,
  IRSdtProps,
  IRSection,
  IRTable,
} from "./types";
import { WarningSink } from "./types";
import { emuToPx, halfPointsToPx, round2, twipsToPx } from "./units";

/** Resolves a hyperlink relationship id to a URL (the part's external rels). */
export type LinkResolver = (relId: string) => string | undefined;
const NO_LINKS: LinkResolver = () => undefined;

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
  /** Map a block list (document body, header story, footer story). resolveLink
   *  resolves that part's hyperlink rels (body and bands have separate rels). */
  mapBlocks(blocks: IRBlock[], media: MediaStore, resolveLink?: LinkResolver): Block[];
  mapSection(section: IRSection | null): SectionProps;
  /** The editor assumes at least one block (the caret needs a home). */
  emptyParagraph(): Paragraph;
  /** Model list definitions for the lists actually referenced — for Document.lists. */
  lists(): Record<string, ListDefinition>;
  /** Bookmark name → model block id, collected across all mapped stories. */
  bookmarks(): Record<string, string>;
  /** Footnotes referenced (in document order) — the pipeline maps their bodies
   *  into Document.footnotes. `noteId` is the model key, `docxId` the source id. */
  footnoteRefs(): { docxId: string; noteId: string }[];
}

export function createMapper(
  warnings: WarningSink,
  resolver: StyleResolver,
  sdts: Record<string, IRSdtProps> = {},
  numbering: NumberingData = new Map(),
): Mapper {
  // Import-distinct id prefix: commands.ts mints `n…`, sampleDoc `b…`.
  let nextId = 0;
  const id = (): string => `i${nextId++}`;

  // Bookmark name → the id of the first model paragraph it anchors.
  const bookmarkMap: Record<string, string> = {};

  // Footnote markers numbered sequentially in document order (docx id → number),
  // matching the editor's renumber convention. The pipeline maps the referenced
  // bodies into Document.footnotes keyed by `fn<docxId>`.
  const footnoteNum = new Map<string, string>();
  const footnoteOrder: string[] = [];
  const footnoteNumber = (docxId: string): string => {
    let num = footnoteNum.get(docxId);
    if (num === undefined) {
      num = String(footnoteNum.size + 1);
      footnoteNum.set(docxId, num);
      footnoteOrder.push(docxId);
    }
    return num;
  };

  // Model list definitions, built lazily for referenced numIds only (a report
  // carries far more definitions than it uses).
  const usedLists = new Map<string, ListDefinition>();
  const listDefFor = (numId: string): ListDefinition | undefined => {
    const cached = usedLists.get(numId);
    if (cached) return cached;
    const ir = numbering.get(numId);
    if (!ir) return undefined;
    const def = buildListDefinition(ir);
    usedLists.set(numId, def);
    return def;
  };

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

  /** Apply list membership to a resolved paragraph style. The engine ADDS the
   *  list level's indent to the paragraph's own (engine.ts ~L446), so subtract
   *  the level indent here to avoid double-counting; the marker hang is handled
   *  by the level's hangingPx, so zero the paragraph's first-line indent.
   *  Returns the level indent (px) so soft-break continuation lines — which
   *  drop list membership — can restore it and stay aligned under the item. */
  const applyListMembership = (style: ParaStyle, ref: IRParaProps["list"]): number => {
    if (!ref) return 0; // undefined (no list) or null (explicitly removed)
    const def = listDefFor(ref.numId);
    if (!def) {
      warnings.add("list-missing", "A list reference had no matching definition — markers were dropped.");
      return 0;
    }
    style.list = { listId: ref.numId, level: ref.level };
    const lvl = def.levels[Math.min(ref.level, def.levels.length - 1)];
    const levelIndentPx = lvl ? lvl.indentLeftPx : 0;
    style.indentLeftPx = Math.max(0, round2(style.indentLeftPx - levelIndentPx));
    style.indentFirstLinePx = 0;
    return levelIndentPx;
  };

  const mapBlocks = (blocks: IRBlock[], media: MediaStore, resolveLink: LinkResolver = NO_LINKS): Block[] => {
    const out: Block[] = [];
    // A paragraph carrying w:sectPr ends a section; unless the break is
    // "continuous", whatever follows starts a new page.
    let sectionPageBreak = false;
    for (const irBlock of blocks) {
      const mapped =
        irBlock.kind === "paragraph"
          ? mapParagraph(irBlock, media, resolveLink)
          : [mapTable(irBlock, media, resolveLink)];
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
  function mapParagraph(ir: IRParagraph, media: MediaStore, resolveLink: LinkResolver): Block[] {
    const effPara = resolver.para(ir.props);
    const style = mapParaStyle(effPara);
    const listLevelIndentPx = applyListMembership(style, effPara.list);
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
          runs.push(mapRun(inline.text, effective, resolveLink, inline.sdtId));
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
      // A soft break inside a list item is the SAME item: only the first split
      // is numbered. Continuation lines drop list membership (no marker, no
      // counter bump) and restore the level indent so they align under the text.
      if (i > 0 && p.style.list) {
        delete p.style.list;
        p.style.indentLeftPx = round2(p.style.indentLeftPx + listLevelIndentPx);
      }
    });
    // Bookmarks anchored in this paragraph point at its first emitted block.
    if (ir.bookmarks && ir.bookmarks.length > 0 && blocks[0]) {
      for (const name of ir.bookmarks) bookmarkMap[name] = blocks[0]!.id;
    }
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

  function mapRun(rawText: string, effective: IRRunProps, resolveLink: LinkResolver, sdtId?: string): Run {
    let text = rawText;
    if (text.includes("\t")) {
      warnings.add("tabs", "Tab stops became fixed spaces (no tab-stop layout).");
      text = text.replaceAll("\t", TAB_AS_SPACES);
    }
    const style = mapCharStyle(effective);
    if (effective.linkRelId) {
      const url = resolveLink(effective.linkRelId);
      if (url) style.link = url;
      else warnings.add("links-unresolved", "A hyperlink target could not be resolved and was dropped.");
    } else if (effective.linkAnchor) {
      style.link = `#${effective.linkAnchor}`; // in-document bookmark
    }
    if (effective.footnoteId !== undefined) {
      // The ref run's TEXT is its number (the engine paints it at page bottom);
      // numbers are sequential in document order, matching the editor's convention.
      text = footnoteNumber(effective.footnoteId);
      style.footnoteRef = `fn${effective.footnoteId}`;
      style.verticalAlign = "super";
    }
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

  function mapTable(ir: IRTable, media: MediaStore, resolveLink: LinkResolver): TableBlock {
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
          if (b.kind === "paragraph") blocks.push(...mapParagraph(b, media, resolveLink));
          else blocks.push(mapTable(b, media, resolveLink)); // nested table — model renders one level
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
    const section: SectionProps = {
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
    if (ir.columns) {
      // OOXML default column gap is 720 twips (0.5") when w:space is absent.
      section.columns = { count: ir.columns.count, gapPx: round2(twipsToPx(ir.columns.spaceTwips ?? 720)) };
    }
    if (ir.pageNumberStart !== undefined) section.pageNumberStart = ir.pageNumberStart;
    return section;
  }

  return {
    mapBlocks,
    mapSection,
    emptyParagraph,
    lists: () => Object.fromEntries(usedLists),
    bookmarks: () => bookmarkMap,
    footnoteRefs: () => footnoteOrder.map((docxId) => ({ docxId, noteId: `fn${docxId}` })),
  };
}

// ---------------------------------------------------------------------------
// Lists (numbering.xml IR → model ListDefinition)

const NUM_FORMAT_MAP: Record<string, ListNumberFormat> = {
  decimal: "decimal",
  decimalZero: "decimal",
  lowerLetter: "lowerLetter",
  upperLetter: "upperLetter",
  lowerRoman: "lowerRoman",
  upperRoman: "upperRoman",
  bullet: "bullet",
};

/** Symbol/Wingdings private-use code points Word uses for bullets → Unicode. */
function bulletGlyph(glyph: string): string {
  if (glyph.length !== 1) return glyph; // already a real char, or multi-char marker
  switch (glyph.charCodeAt(0)) {
    case 0xf0b7: // Symbol  (filled round bullet)
    case 0x00b7: // middle dot
      return '•'; // •
    case 0xf0a7: // Wingdings  (filled square)
    case 0x00a7:
      return '▪'; // ▪
    case 0xf06f: // Wingdings  (open square)
      return '▫'; // ▫
    case 0xf0d8: // Wingdings  (arrowhead)
      return '‣'; // ‣
    case 0xf0fc: // Wingdings  (check)
      return '✔'; // ✔
    default:
      return glyph;
  }
}

function buildListDefinition(ir: IRListDefinition): ListDefinition {
  const levels: ListLevel[] = [];
  for (let i = 0; i < 9; i++) {
    levels.push(buildListLevel(ir.levels[i], i));
  }
  return { id: ir.id, levels };
}

function buildListLevel(ir: IRListDefinition["levels"][number] | undefined, i: number): ListLevel {
  if (!ir) {
    // Hole in the definition — synthesize a sane decimal level so the engine's
    // levels[level] lookup never hits undefined.
    return { format: "decimal", text: `%${i + 1}.`, indentLeftPx: 24 + i * 24, hangingPx: 18, start: 1 };
  }
  const format = NUM_FORMAT_MAP[ir.format] ?? (ir.format === "none" ? "bullet" : "decimal");
  const level: ListLevel = {
    format,
    text: format === "bullet" ? "" : ir.lvlText,
    indentLeftPx: ir.indentLeftTwips !== undefined ? round2(twipsToPx(ir.indentLeftTwips)) : 24 + i * 24,
    hangingPx: ir.hangingTwips !== undefined ? round2(twipsToPx(ir.hangingTwips)) : 18,
    start: ir.start,
  };
  if (format === "bullet") {
    level.bulletChar = ir.format === "none" ? "" : bulletGlyph(ir.lvlText);
  }
  if (ir.markerRunProps) {
    const marker = mapCharPatch(ir.markerRunProps);
    if (Object.keys(marker).length > 0) level.markerStyle = marker;
  }
  return level;
}

// ---------------------------------------------------------------------------
// Style mapping (pure)

/** Word's 16 named highlight colors → hex. */
const HIGHLIGHT_HEX: Record<string, string> = {
  yellow: "#ffff00", green: "#00ff00", cyan: "#00ffff", magenta: "#ff00ff",
  blue: "#0000ff", red: "#ff0000", darkBlue: "#000080", darkCyan: "#008080",
  darkGreen: "#008000", darkMagenta: "#800080", darkRed: "#800000", darkYellow: "#808000",
  darkGray: "#808080", lightGray: "#c0c0c0", black: "#000000", white: "#ffffff",
};

function mapCharStyle(props: IRRunProps): CharStyle {
  const style: CharStyle = { ...DEFAULT_CHAR };
  applyRunProps(style, props);
  return style;
}

/** Shared run-property mapping for full CharStyle and partial style-gallery patches. */
function applyRunProps(style: Partial<CharStyle>, props: IRRunProps): void {
  if (props.fontAscii) style.fontFamily = `${props.fontAscii}, serif`;
  if (props.sizeHalfPoints !== undefined) style.fontSizePx = round2(halfPointsToPx(props.sizeHalfPoints));
  if (props.bold !== undefined) style.bold = props.bold;
  if (props.italic !== undefined) style.italic = props.italic;
  if (props.underline !== undefined) style.underline = props.underline;
  if (props.strikethrough !== undefined) style.strikethrough = props.strikethrough;
  if (props.color && props.color !== "auto") style.color = `#${props.color.toLowerCase()}`;
  if (props.highlight) {
    const hex = HIGHLIGHT_HEX[props.highlight];
    if (hex) style.highlightColor = hex;
  }
  if (props.vertAlign === "superscript") style.verticalAlign = "super";
  else if (props.vertAlign === "subscript") style.verticalAlign = "sub";
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
  applyRunProps(out, props);
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
  if (props.keepLinesTogether) out.keepLinesTogether = true;
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
  if (props.keepLinesTogether) style.keepLinesTogether = true;
  if (props.pageBreakBefore) style.pageBreakBefore = true;
  if (props.styleId) style.namedStyle = props.styleId;
  return style;
}
