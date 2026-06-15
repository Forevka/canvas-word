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
  BookmarkRange,
  CharStyle,
  DocPosition,
  Document,
  ImageBlock,
  ParaStyle,
  Paragraph,
  Run,
  SdtProps,
  SectionPatch,
  SectionProps,
  TabAlign,
  TableBlock,
  TableCell,
  TabLeader,
  TabStop,
} from "@cw/shared";
import type { NamedStyle, Stylesheet } from "@cw/shared";
import type { ListDefinition, ListLevel, ListNumberFormat } from "@cw/shared";
import { normalizeRuns } from "@cw/shared";
import { resolveCellBorders, type BorderSources, type CellPosition } from "./borders";
import type { MediaStore } from "./media";
import type { NumberingData } from "./numbering";
import type { ResolvedTableStyle, StyleResolver, StylesData } from "./styles";
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
  IRTableCell,
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

/** Word's default cell margins (TableNormal): 0 top/bottom, 108 twips (~7.2px)
 *  left/right. The base each table's w:tblCellMar and each cell's w:tcMar
 *  override per side. Vertical is intentionally 0 — a single-line row is only as
 *  tall as its line, which is why a symmetric pad makes Word tables too tall. */
const WORD_CELL_MARGIN_TWIPS = { top: 0, right: 108, bottom: 0, left: 108 } as const;

/** US Letter at 96dpi, 1in margins — matches sampleDoc. */
const DEFAULT_SECTION: SectionProps = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
};

export interface Mapper {
  /** Map a block list (document body, header story, footer story). resolveLink
   *  resolves that part's hyperlink rels (body and bands have separate rels). */
  mapBlocks(blocks: IRBlock[], media: MediaStore, resolveLink?: LinkResolver): Block[];
  mapSection(section: IRSection | null): SectionProps;
  /** The editor assumes at least one block (the caret needs a home). */
  emptyParagraph(): Paragraph;
  /** Model list definitions for the lists actually referenced — for Document.lists. */
  lists(): Record<string, ListDefinition>;
  /** Bookmark name → resolved model range, collected across all mapped stories. */
  bookmarks(): Record<string, BookmarkRange>;
  /** Footnotes referenced (in document order) — the pipeline maps their bodies
   *  into Document.footnotes. `noteId` is the model key, `docxId` the source id. */
  footnoteRefs(): { docxId: string; noteId: string }[];
}

export function createMapper(
  warnings: WarningSink,
  resolver: StyleResolver,
  sdts: Record<string, IRSdtProps> = {},
  numbering: NumberingData = new Map(),
  tableStyle: (styleId: string) => ResolvedTableStyle = () => ({}),
  /** The document (body) section's page size in twips — section breaks that
   *  match it are geometry-preserving and flow rather than forcing a page. */
  refPgSize?: { w: number; h: number },
): Mapper {
  // Import-distinct id prefix: commands.ts mints `n…`, sampleDoc `b…`.
  let nextId = 0;
  const id = (): string => `i${nextId++}`;

  // Bookmark markers resolved to model positions, keyed by w:id; paired into
  // ranges in bookmarks(). Start carries the name; end may live in a later block.
  const bookmarkStarts = new Map<string, { name: string; pos: DocPosition }>();
  const bookmarkEnds = new Map<string, DocPosition>();

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

  const sectionChangesGeometry = (props: IRParaProps): boolean => {
    const sz = props.sectionPgSize;
    return !!sz && !!refPgSize && (sz.w !== refPgSize.w || sz.h !== refPgSize.h);
  };

  /** A section worth its own page+geometry: different page size, newspaper
   *  columns, or a page-number restart. (A bare footer/header switch with the
   *  same geometry flows instead — see mapBlocks.) */
  const sectionIsDistinct = (props: IRParaProps): boolean =>
    sectionChangesGeometry(props) || !!props.sectionColumns || props.sectionPageNumberStart !== undefined;

  /** Build a SectionPatch from a section-ending paragraph's geometry/columns. */
  const buildSectionPatch = (props: IRParaProps): SectionPatch => {
    const patch: SectionPatch = {};
    if (props.sectionPgSize) {
      patch.pageWidthPx = round2(twipsToPx(props.sectionPgSize.w));
      patch.pageHeightPx = round2(twipsToPx(props.sectionPgSize.h));
    }
    if (props.sectionMarginTwips) {
      const m = props.sectionMarginTwips;
      patch.marginPx = {
        top: round2(twipsToPx(m.top)),
        right: round2(twipsToPx(m.right)),
        bottom: round2(twipsToPx(m.bottom)),
        left: round2(twipsToPx(m.left)),
      };
    }
    if (props.sectionColumns) {
      patch.columns = { count: props.sectionColumns.count, gapPx: round2(twipsToPx(props.sectionColumns.spaceTwips ?? 720)) };
    }
    if (props.sectionPageNumberStart !== undefined) patch.pageNumberStart = props.sectionPageNumberStart;
    return patch;
  };

  const lastParagraphOf = (mapped: Block[]): Paragraph | undefined => {
    for (let i = mapped.length - 1; i >= 0; i--) if (mapped[i]!.kind === "paragraph") return mapped[i] as Paragraph;
    return undefined;
  };

  const mapBlocks = (blocks: IRBlock[], media: MediaStore, resolveLink: LinkResolver = NO_LINKS): Block[] => {
    const out: Block[] = [];
    // A paragraph carrying w:sectPr ends a section (OOXML: its props describe the
    // section ending there). A *distinct* section (different geometry/columns/page
    // numbering) becomes a real ParaStyle.sectionBreak on that paragraph — the
    // engine pages and applies the geometry. But these generated reports emit a
    // Next Page break wherever only the footer changes (dozens, identical
    // geometry); a spec-literal page break would strand half of every page, so
    // those FLOW — unless the next section opens with a heading (a real chapter
    // boundary), which page-breaks. "continuous" never breaks.
    let pending = false;
    // The first page-type section break ends the title/cover page, which always
    // stands alone — force it even though its new section (the Letter of
    // Transmittal) opens with an address rather than a heading. Gated on the cover
    // actually holding content (its title image/text), so a leading bare section
    // break with nothing before it still flows.
    let pendingForce = false;
    let seenPageSection = false;
    let coverHasContent = false;
    const hasRealContent = (bs: Block[]): boolean =>
      bs.some((b) => b.kind !== "paragraph" || b.runs.some((r) => r.text.trim().length > 0 && !r.style.hidden));

    // Looking PAST blank/hidden lead-in paragraphs, does the section beginning
    // after `afterIndex` open with a real section title? These generated reports
    // format titles directly (bold/centered) rather than with a Heading style, so
    // isHeading misses them — but Word always bookmarks a section title for its
    // TOC (_Toc<n> targets, or the TOC's own _Part_TOC anchor). A titled section
    // starts a new page (Word's Next Page break); plain continuing body text
    // after a footer-only break flows. Hidden-anchor paragraphs (w:vanish text,
    // dropped on import) are skipped here via the IR's vanish flag.
    const isHeadingLedSection = (afterIndex: number): boolean => {
      for (let i = afterIndex + 1; i < blocks.length; i++) {
        const b = blocks[i]!;
        if (b.kind !== "paragraph") return false; // opens with a table/image, not a title
        const visible = b.inlines.some(
          (inl) => inl.kind === "run" && inl.text.trim().length > 0 && !inl.props.vanish,
        );
        if (!visible) continue; // blank line or hidden anchor — keep looking
        if (resolver.isHeading(b.props.styleId)) return true;
        return (b.bookmarks ?? []).some((n) => /toc/i.test(n));
      }
      return false;
    };

    for (let bi = 0; bi < blocks.length; bi++) {
      const irBlock = blocks[bi]!;
      const mapped =
        irBlock.kind === "paragraph"
          ? mapParagraph(irBlock, media, resolveLink)
          : [mapTable(irBlock, media, resolveLink)];

      // Custom-field membership: an IR block in a field's result region stamps its
      // fieldId onto every model block it produced (a paragraph may split).
      if (irBlock.fieldId) for (const b of mapped) b.fieldId = irBlock.fieldId;

      // Resolve a pending (same-geometry) section break against this block.
      // Blank/hidden lead-in paragraphs are skipped so the break lands on the
      // section's first VISIBLE line (the title), not an empty anchor above it.
      if (pending && mapped.length > 0) {
        const first = mapped[0]!;
        // Hidden runs count as blank — a w:vanish anchor paragraph is invisible.
        const empty = first.kind === "paragraph" && first.runs.every((r) => r.text.length === 0 || r.style.hidden);
        if (!empty) {
          const heading = irBlock.kind === "paragraph" && resolver.isHeading(irBlock.props.styleId);
          if (heading || pendingForce) {
            if (first.kind === "paragraph") first.style.pageBreakBefore = true;
            else mapped.unshift({ ...emptyParagraph(), style: { ...documentPara, pageBreakBefore: true } });
          }
          pending = false;
          pendingForce = false;
        }
        // empty paragraph: keep pending and look past it for the real section start
      }

      // Does THIS paragraph end a (page-type) section?
      if (irBlock.kind === "paragraph" && irBlock.props.sectionBreak === "page") {
        const props = irBlock.props;
        if (sectionIsDistinct(props)) {
          // Apply the section's geometry: sectionBreak sits on the paragraph that
          // ENDS the section (engine pages at the next block).
          let carrier = lastParagraphOf(mapped);
          if (!carrier) {
            carrier = emptyParagraph();
            mapped.push(carrier);
          }
          carrier.style.sectionBreak = { type: "nextPage", props: buildSectionPatch(props) };
          pending = false;
        } else {
          if (props.sectionHasBands) {
            warnings.add(
              "section-bands-flattened",
              "Per-section headers/footers on geometry-preserving section breaks were not applied (the document section's are used).",
            );
          }
          pending = true;
          // Break out of the section when it's the cover→body boundary OR the new
          // section opens with a real title (heading style / TOC-bookmarked) —
          // Word's Next Page behavior for titled sections (TOC, Flood Map, …).
          pendingForce = (!seenPageSection && coverHasContent) || isHeadingLedSection(bi);
        }
        seenPageSection = true;
      }
      out.push(...mapped);
      if (!seenPageSection) coverHasContent ||= hasRealContent(mapped);
    }
    keepHeadingWithFollowingFigure(out);
    return out;
  };

  /** Don't strand a section heading above its figure: when a heading is followed
   *  — through only blank lines and a short caption — by an image or table, mark
   *  the heading (and that caption) "keep with next" so the engine's keep-chain
   *  rides them to the figure's page. These reports leave the heading's keepNext
   *  off, so a Heading1 like FLOOD MAP otherwise sits at a page bottom while its
   *  map flows to the next page. Bounded to a tight group so a heading above real
   *  body text (with a figure further down) is NOT pulled. */
  const keepHeadingWithFollowingFigure = (blocks: Block[]): void => {
    const isFigure = (b: Block): boolean => b.kind === "image" || b.kind === "table";
    const isBlank = (b: Block): boolean =>
      b.kind === "paragraph" && b.runs.every((r) => r.text.trim().length === 0 || r.style.hidden);
    const textLen = (b: Block): number => (b.kind === "paragraph" ? b.runs.reduce((s, r) => s + r.text.length, 0) : 0);
    for (let i = 0; i < blocks.length; i++) {
      const h = blocks[i]!;
      if (h.kind !== "paragraph" || !h.style.namedStyle || !resolver.isHeading(h.style.namedStyle)) continue;
      const between: Paragraph[] = [];
      let captions = 0;
      let figureAt = -1;
      for (let j = i + 1; j < blocks.length && j <= i + 6; j++) {
        const b = blocks[j]!;
        if (isFigure(b)) {
          figureAt = j;
          break;
        }
        if (b.kind !== "paragraph") break; // anything else ends the tight group
        if (isBlank(b)) continue; // blank spacer — the engine bridges it
        if (captions >= 2 || textLen(b) > 300) break; // real body content, not a caption
        captions++;
        between.push(b);
      }
      if (figureAt < 0) continue;
      h.style.keepWithNext = true;
      for (const cap of between) cap.style.keepWithNext = true;
    }
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
    // Set by a page/column break; consumed by the NEXT emitted paragraph — what
    // follows the break starts a new page (or newspaper column).
    let pendingPageBreak = false;
    let pendingColumnBreak = false;

    const paraOf = (paraRuns: Run[]): Paragraph => {
      const paraStyle = { ...style };
      if (pendingPageBreak) {
        paraStyle.pageBreakBefore = true;
        pendingPageBreak = false;
      }
      if (pendingColumnBreak) {
        paraStyle.columnBreakBefore = true;
        pendingColumnBreak = false;
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
          if (!inline.page && !inline.column) {
            warnings.add("soft-breaks", "Soft line breaks (Shift+Enter) became paragraph breaks.");
          }
          // A break with nothing pending is an empty visual line ("a\n\nb").
          if (runs.length > 0) flushPara();
          else blocks.push(paraOf([]));
          if (inline.page) pendingPageBreak = true;
          if (inline.column) pendingColumnBreak = true;
          trailingBreak = true;
          break;
        case "image": {
          const image = mapImage(inline, media, style.align);
          if (image) {
            flushPara();
            // An image can't carry a page/column break — give the break a carrier.
            if (pendingPageBreak || pendingColumnBreak) blocks.push(paraOf([]));
            blocks.push(image);
            trailingBreak = false;
          }
          break;
        }
        case "run": {
          const effective = resolver.run(ir.props.styleId, inline.props);
          // Hidden text (w:vanish) is PRESERVED with style.hidden (set via
          // applyRunProps) — never displayed, never caret-reachable, protected
          // from deletion — so a round-trip re-emits it instead of losing it.
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
    // Resolve this paragraph's bookmark markers to model positions. The home is
    // the first emitted block; offsets clamp to its text (exact for the common
    // single-block paragraph; a soft-break split anchors into the first block).
    if (ir.bookmarkMarkers && ir.bookmarkMarkers.length > 0 && blocks[0]) {
      const home = blocks[0]!;
      const homeLen = home.kind === "paragraph" ? home.runs.reduce((s, r) => s + r.text.length, 0) : 0;
      for (const m of ir.bookmarkMarkers) {
        const pos: DocPosition = { blockId: home.id, offset: Math.min(m.offset, homeLen) };
        if (m.kind === "start" && m.name) bookmarkStarts.set(m.id, { name: m.name, pos });
        else if (m.kind === "end") bookmarkEnds.set(m.id, pos);
      }
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
    // "\t" is preserved — the layout engine positions it at the paragraph's tab
    // stops (or the default interval); no flattening to spaces.
    let text = rawText;
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

  /** A model cell with its source IR (null for padding) and grid position —
   *  the border/shading pass needs the position; row spans make it non-trivial. */
  interface PlacedCell {
    cell: TableCell;
    ir: IRTableCell | null;
    startRow: number;
    startCol: number;
    colSpan: number;
  }

  /** Collapse the OOXML border/shading cascade onto one placed cell. Borders are
   *  set only when the table carries border info at SOME layer (style, table, or
   *  any cell) — otherwise left absent so the renderer draws its default grid. */
  function styleCell(
    p: PlacedCell,
    ir: IRTable,
    styled: ResolvedTableStyle,
    hasBorderInfo: boolean,
    width: number,
    rowCount: number,
  ): void {
    const rowSpan = p.cell.rowSpan ?? 1;
    const pos: CellPosition = {
      top: p.startRow === 0,
      bottom: p.startRow + rowSpan - 1 === rowCount - 1,
      left: p.startCol === 0,
      right: p.startCol + p.colSpan - 1 === width - 1,
    };
    const shd = p.ir?.shd ?? ir.shd ?? styled.shd; // cell over table over style
    if (shd) p.cell.shading = shd;
    if (hasBorderInfo) {
      const src: BorderSources = { cell: p.ir?.borders, table: ir.borders, style: styled.borders };
      p.cell.borders = resolveCellBorders(src, pos);
    }
  }

  function mapTable(ir: IRTable, media: MediaStore, resolveLink: LinkResolver): TableBlock {
    const buildCell = (irCell: IRTableCell): TableCell => {
      const blocks: Block[] = [];
      for (const b of irCell.blocks) {
        if (b.kind === "paragraph") blocks.push(...mapParagraph(b, media, resolveLink));
        else blocks.push(mapTable(b, media, resolveLink)); // nested table — model renders one level
      }
      const cell: TableCell = { id: id(), blocks: blocks.length > 0 ? blocks : [emptyParagraph()] };
      if (irCell.gridSpan > 1) cell.colSpan = irCell.gridSpan;
      // Cell-margin cascade: this cell's w:tcMar over the table's w:tblCellMar
      // over Word's defaults, resolved per side (a cell may set only top/bottom).
      const side = (s: "top" | "right" | "bottom" | "left"): number =>
        irCell.marginTwips?.[s] ?? ir.cellMarginTwips?.[s] ?? WORD_CELL_MARGIN_TWIPS[s];
      cell.margin = {
        top: round2(twipsToPx(side("top"))),
        right: round2(twipsToPx(side("right"))),
        bottom: round2(twipsToPx(side("bottom"))),
        left: round2(twipsToPx(side("left"))),
      };
      return cell;
    };

    // Column count: the widest row by grid columns. Merge-continuation cells are
    // still present in their rows (each covered row carries a w:vMerge cell), so
    // every row sums to the full width before we drop them.
    const width = Math.max(1, ...ir.rows.map((r) => r.cells.reduce((s, c) => s + Math.max(1, c.gridSpan), 0)));

    // owner[col] = the cell currently spanning down into column `col` (HTML
    // rowspan). A w:vMerge="continue" cell is dropped and bumps its owner's
    // rowSpan; the row simply omits a cell for those columns.
    const owner = new Array<TableCell | undefined>(width);
    const placedRows: PlacedCell[][] = ir.rows.map((irRow, ri) => {
      const placed: PlacedCell[] = [];
      let col = 0;
      for (const irCell of irRow.cells) {
        if (col >= width) break; // overflow beyond the declared grid — drop
        const span = Math.max(1, Math.min(irCell.gridSpan, width - col));
        if (irCell.vMergeContinue) {
          const o = owner[col];
          if (o) o.rowSpan = (o.rowSpan ?? 1) + 1;
          else warnings.add("cell-vmerge-orphan", "A merged-cell continuation had no cell above to extend.");
          col += span; // owner[col..] persists so deeper continuations find it
          continue;
        }
        const cell = buildCell(irCell);
        placed.push({ cell, ir: irCell, startRow: ri, startCol: col, colSpan: span });
        for (let k = 0; k < span && col + k < width; k++) owner[col + k] = cell;
        col += span;
      }
      // Genuinely short row (not rowspan holes): pad the trailing columns.
      while (col < width) {
        const cell: TableCell = { id: id(), blocks: [emptyParagraph()] };
        placed.push({ cell, ir: null, startRow: ri, startCol: col, colSpan: 1 });
        owner[col] = cell;
        col += 1;
      }
      return placed;
    });

    // Borders/shading after all rows so each owner's rowSpan is final.
    const styled = ir.styleId ? tableStyle(ir.styleId) : {};
    // An explicitly-declared (even empty) w:tblBorders/w:tcBorders counts as
    // border info: the author chose "no borders", which must override the
    // renderer's default grid rather than fall through to it.
    const hasBorderInfo =
      !!styled.borders ||
      !!ir.borders ||
      !!ir.bordersSpecified ||
      ir.rows.some((r) => r.cells.some((c) => c.borders || c.bordersSpecified));
    for (const row of placedRows) for (const p of row) styleCell(p, ir, styled, hasBorderInfo, width, ir.rows.length);

    const rows = placedRows.map((row) => ({ cells: row.map((p) => p.cell) }));
    const table: TableBlock = { kind: "table", id: id(), revision: 0, rows };
    // Always emit colFractions of the true width: dropped continuation cells mean
    // a row's cell count no longer equals the column count, so the layout engine
    // can't infer the column count from cells.length anymore.
    table.colFractions = columnFractions(ir.colWidthsTwips, width);
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
    if (ir.headerDistTwips !== undefined) section.headerDistancePx = round2(twipsToPx(ir.headerDistTwips));
    if (ir.footerDistTwips !== undefined) section.footerDistancePx = round2(twipsToPx(ir.footerDistTwips));
    return section;
  }

  return {
    mapBlocks,
    mapSection,
    emptyParagraph,
    lists: () => Object.fromEntries(usedLists),
    bookmarks: () => {
      const out: Record<string, BookmarkRange> = {};
      for (const [id, s] of bookmarkStarts) {
        out[s.name] = { start: s.pos, end: bookmarkEnds.get(id) ?? s.pos }; // point bookmark if no end
      }
      return out;
    },
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

const TAB_ALIGN: Record<string, TabAlign> = {
  left: "left",
  start: "left",
  num: "left",
  center: "center",
  right: "right",
  end: "right",
  decimal: "decimal",
};
const TAB_LEADER: Record<string, TabLeader> = {
  dot: "dot",
  middleDot: "dot",
  hyphen: "dash",
  underscore: "underscore",
};

/** Raw w:tabs → model TabStop[] (twips → px), sorted by position. */
function mapTabStops(raw: NonNullable<IRParaProps["tabStops"]>): TabStop[] {
  const stops = raw.map((t) => {
    const stop: TabStop = { posPx: round2(twipsToPx(t.posTwips)) };
    const align = t.val ? TAB_ALIGN[t.val] : undefined;
    if (align && align !== "left") stop.align = align;
    const leader = t.leader ? TAB_LEADER[t.leader] : undefined;
    if (leader) stop.leader = leader;
    return stop;
  });
  stops.sort((a, b) => a.posPx - b.posPx);
  return stops;
}

/** Column fractions of length `width` (sum ≈ 1). From w:tblGrid when it matches
 *  the column count, else equal columns. The last cell absorbs rounding. */
function columnFractions(widthsTwips: number[] | undefined, width: number): number[] {
  const fractions =
    widthsTwips && widthsTwips.length === width && widthsTwips.reduce((s, w) => s + w, 0) > 0
      ? widthsTwips.map((w) => Math.round((w / widthsTwips.reduce((s, x) => s + x, 0)) * 10000) / 10000)
      : Array.from({ length: width }, () => Math.round((1 / width) * 10000) / 10000);
  fractions[fractions.length - 1] =
    Math.round((1 - fractions.slice(0, -1).reduce((s, f) => s + f, 0)) * 10000) / 10000;
  return fractions;
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
  if (props.vanish) style.hidden = true; // preserved, never displayed (see CharStyle.hidden)
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
  if (props.indentRightTwips !== undefined) out.indentRightPx = round2(twipsToPx(props.indentRightTwips));
  if (props.indentFirstLineTwips !== undefined)
    out.indentFirstLinePx = round2(twipsToPx(props.indentFirstLineTwips));
  if (props.keepWithNext) out.keepWithNext = true;
  if (props.keepLinesTogether) out.keepLinesTogether = true;
  if (props.tabStops) out.tabStops = mapTabStops(props.tabStops);
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
  if (props.indentRightTwips !== undefined) style.indentRightPx = round2(twipsToPx(props.indentRightTwips));
  if (props.indentFirstLineTwips !== undefined)
    style.indentFirstLinePx = round2(twipsToPx(props.indentFirstLineTwips));
  if (props.keepWithNext) style.keepWithNext = true;
  if (props.keepLinesTogether) style.keepLinesTogether = true;
  if (props.tabStops) style.tabStops = mapTabStops(props.tabStops);
  if (props.pageBreakBefore) style.pageBreakBefore = true;
  if (props.outlineLevel !== undefined) style.outlineLevel = props.outlineLevel;
  if (props.styleId) style.namedStyle = props.styleId;
  return style;
}
