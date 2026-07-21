// IR → canvas-word model. This is the lossy stage: everything the model
// can't hold is decided HERE (with a warning), never silently in the parser.
//
// Policies (see IMPORT.md "Lossy mappings"):
//   w:br soft break  → paragraph split (split halves hug: zero space between)
//   w:tab            → fixed spaces
//   gridSpan/vMerge  → padded uniform cell grid
//   inline images    → block-level ImageBlock (paragraph splits around it)
//   images in cells  → ImageBlock inside the cell's block list (incl. anchored/floating)

import type {
  Block,
  BookmarkRange,
  CharStyle,
  ColumnEntry,
  DocPosition,
  EmphasisMark,
  ImageBlock,
  PageBorderEdge,
  PageBorders,
  ParaStyle,
  LineNumbering,
  Paragraph,
  Run,
  SdtProps,
  RowProps,
  SectionBreakType,
  SectionPatch,
  SectionProps,
  ShapeBlock,
  ShapePreset,
  TabAlign,
  TableBlock,
  TableCell,
  TableRow,
  TabLeader,
  TabStop,
  UnderlineStyle,
} from "@cw/shared";
import type { NamedStyle, Stylesheet } from "@cw/shared";
import type { ListDefinition, ListLevel, ListNumberFormat } from "@cw/shared";
import type { TableCond, TableCondProps, TableStyle } from "@cw/shared";
import { AUTO_PARA_SPACING_PX, HIGHLIGHT_HEX, normalizeRuns } from "@cw/shared";
import { cellBordersFromIR, paraBordersFromIR, resolveCellBorders, runBorderFromIR, tableBordersFromIR, type BorderSources, type CellPosition } from "./borders";
import type { MediaStore } from "./media";
import type { NumberingData } from "./numbering";
import type { ResolvedTableStyle, StyleResolver, StylesData } from "./styles";
import type {
  IRBlock,
  IRBorders,
  IRInline,
  IRLineNumbering,
  IRListDefinition,
  IRPageBorders,
  IRParaProps,
  IRParagraph,
  IRRunProps,
  IRSdtProps,
  IRSection,
  IRTable,
  IRTableCell,
} from "./types";
import { WarningSink } from "./types";
import { emuToPx, halfPointsToPx, marginTwipsToPx, round2, round4, twipsToPx } from "./units";

const PAGE_BORDER_STYLES = new Set(["single", "double", "dashed", "dotted", "thick", "none"]);

/** IR newspaper columns → model columns (twips→px, per-column widths + sep). */
function irColumnsToModel(
  c: NonNullable<IRSection["columns"]>,
): { count: number; gapPx: number; sep?: boolean; cols?: ColumnEntry[] } {
  // OOXML default column gap is 720 twips (0.5") when w:space is absent.
  const out = { count: c.count, gapPx: round2(twipsToPx(c.spaceTwips ?? 720)) } as {
    count: number;
    gapPx: number;
    sep?: boolean;
    cols?: ColumnEntry[];
  };
  if (c.sep) out.sep = true;
  if (c.cols && c.cols.length === c.count) {
    out.cols = c.cols.map((col) => ({
      widthPx: round2(twipsToPx(col.wTwips)),
      spaceAfterPx: round2(twipsToPx(col.spaceTwips)),
    }));
  }
  return out;
}

/** IR page borders → model PageBorders (eighth-points→px, points→px, color). */
function irPageBordersToModel(b: IRPageBorders): PageBorders {
  const edge = (e: IRPageBorders["top"]): PageBorderEdge | undefined => {
    if (!e) return undefined;
    const style = (PAGE_BORDER_STYLES.has(e.style) ? e.style : "single") as PageBorderEdge["style"];
    const out: PageBorderEdge = {
      style,
      widthPx: e.sz !== undefined ? round2(e.sz / 6) : 1, // w:sz is eighth-points
      color: !e.color || e.color === "auto" ? "#000000" : `#${e.color}`,
    };
    if (e.space !== undefined) out.spacePx = round2((e.space * 96) / 72); // w:space is points
    return out;
  };
  const out: PageBorders = { offsetFrom: b.offsetFrom ?? "page" };
  const top = edge(b.top);
  if (top) out.top = top;
  const right = edge(b.right);
  if (right) out.right = right;
  const bottom = edge(b.bottom);
  if (bottom) out.bottom = bottom;
  const left = edge(b.left);
  if (left) out.left = left;
  return out;
}

/** IR line numbering → model LineNumbering (distance twips→px; counts as-is). */
function irLineNumberingToModel(ln: IRLineNumbering): LineNumbering {
  const out: LineNumbering = {};
  if (ln.countBy !== undefined) out.countBy = ln.countBy;
  if (ln.start !== undefined) out.start = ln.start;
  if (ln.restart !== undefined) out.restart = ln.restart;
  if (ln.distanceTwips !== undefined) out.distancePx = round2(twipsToPx(ln.distanceTwips));
  return out;
}

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
  /** Endnotes referenced (in document order) — the pipeline maps their bodies
   *  into Document.endnotes. `noteId` is the model key, `docxId` the source id. */
  endnoteRefs(): { docxId: string; noteId: string }[];
  /** The first `TOC` field's block id + verbatim instruction, or undefined. */
  tocField(): { blockId: string; instruction: string } | undefined;
}

/** Cell/table/table-style shading cascade pick (issue #150): return the first
 *  layer that DEFINES shading — a fill string OR an explicit clear (`null`) — so a
 *  clear at a more-specific layer suppresses a fill inherited from a less-specific
 *  one. `undefined` means "this layer didn't touch shading" and is skipped. */
function firstDefinedShd(...layers: (string | null | undefined)[]): string | null | undefined {
  for (const l of layers) if (l !== undefined) return l;
  return undefined;
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
  /** The document (body) section's non-inheriting OWN properties (line numbering,
   *  page-number restart). The body sectPr describes the LAST section, so it's the
   *  "following section" for the final in-paragraph break — used to stop those
   *  properties bleeding backward across a dropped break (see sectionOwnBleeds). */
  refSectionOwn?: { lineNumbering?: IRLineNumbering | undefined; pageNumberStart?: number | undefined },
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

  // Endnote markers numbered sequentially in document order (docx id → number),
  // mirroring footnotes. The pipeline maps the referenced bodies into
  // Document.endnotes keyed by `en<docxId>`.
  const endnoteNum = new Map<string, string>();
  const endnoteOrder: string[] = [];
  const endnoteNumber = (docxId: string): string => {
    let num = endnoteNum.get(docxId);
    if (num === undefined) {
      num = String(endnoteNum.size + 1);
      endnoteNum.set(docxId, num);
      endnoteOrder.push(docxId);
    }
    return num;
  };

  // Model list definitions, built lazily as numIds are referenced during the body
  // walk; the lists() getter then realizes any remaining DEFINED-but-unused ones
  // so authored list styles survive a round-trip before being applied.
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
    // Tri-state (issue #152): null = explicit numId=0 removal (opt out of an
    // inherited style list) — record the marker so export re-emits it; undefined =
    // simply not in a list (inherit).
    if (ref === null) {
      style.listCleared = true;
      return 0;
    }
    if (!ref) return 0;
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
    sectionChangesGeometry(props) ||
    !!props.sectionColumns ||
    props.sectionPageNumberStart !== undefined ||
    !!props.sectionPgBorders ||
    // An even/odd page-parity break or line numbering is a real, page-breaking
    // section even when the geometry is identical to the document section.
    props.sectionBreakType === "evenPage" ||
    props.sectionBreakType === "oddPage" ||
    !!props.sectionLineNumbering;

  /** Build a SectionPatch from a section-ending paragraph's geometry/columns. */
  const buildSectionPatch = (props: IRParaProps): SectionPatch => {
    const patch: SectionPatch = {};
    if (props.sectionPgSize) {
      patch.pageWidthPx = round2(twipsToPx(props.sectionPgSize.w));
      patch.pageHeightPx = round2(twipsToPx(props.sectionPgSize.h));
    }
    if (props.sectionMarginTwips) {
      patch.marginPx = marginTwipsToPx(props.sectionMarginTwips);
    }
    if (props.sectionColumns) patch.columns = irColumnsToModel(props.sectionColumns);
    if (props.sectionPageNumberStart !== undefined) patch.pageNumberStart = props.sectionPageNumberStart;
    if (props.sectionHeaderDistTwips !== undefined) patch.headerDistancePx = round2(twipsToPx(props.sectionHeaderDistTwips));
    if (props.sectionFooterDistTwips !== undefined) patch.footerDistancePx = round2(twipsToPx(props.sectionFooterDistTwips));
    if (props.sectionPgBorders) patch.pageBorders = irPageBordersToModel(props.sectionPgBorders);
    if (props.sectionLineNumbering) patch.lineNumbering = irLineNumberingToModel(props.sectionLineNumbering);
    return patch;
  };

  const lastParagraphOf = (mapped: Block[]): Paragraph | undefined => {
    for (let i = mapped.length - 1; i >= 0; i--) if (mapped[i]!.kind === "paragraph") return mapped[i] as Paragraph;
    return undefined;
  };

  // The first TOC field's anchor (block id + instruction), captured as blocks are
  // mapped (a paragraph may split into several blocks, so an ordinal can't find it).
  let tocAnchor: { blockId: string; instruction: string } | undefined;

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

    // Backward attribution: a w:sectPr describes the section it CLOSES, so whether a
    // heading starts a new page depends on ITS OWN section's type — the NEXT sectPr,
    // not the one before it. The forward `pending` above breaks the section AFTER a
    // page-type break, which silently misses a heading-led page-type section sitting
    // behind a "continuous" break (e.g. a full-page map between two chapters). We
    // track where the current section began and its first visible line so, when the
    // section closes as page-type, the break can land on that title directly.
    let sectionStartBi = 0;
    let sectionFirstVisible: Paragraph | null = null;
    const firstVisiblePara = (bs: Block[]): Paragraph | undefined => {
      for (const b of bs)
        if (b.kind === "paragraph" && b.runs.some((r) => r.text.trim().length > 0 && !r.style.hidden)) return b;
      return undefined;
    };

    // The non-inheriting OWN properties (line numbering, page-number restart) of
    // the section that FOLLOWS the break ending at `afterIndex`: the next
    // section-ending paragraph, or this body section when none follows. A
    // geometry-preserving break is normally dropped and its section flowed into
    // that following one — but these properties never inherit (see
    // effectiveSection), so if the following section declares one this section
    // lacks it would bleed BACKWARD across the merged boundary (e.g. line numbers
    // appearing on every page before a line-numbered section).
    const nextSectionOwn = (
      afterIndex: number,
    ): { lineNumbering?: IRLineNumbering | undefined; pageNumberStart?: number | undefined } => {
      for (let i = afterIndex + 1; i < blocks.length; i++) {
        const b = blocks[i]!;
        if (b.kind === "paragraph" && b.props.sectionBreak !== undefined) {
          return { lineNumbering: b.props.sectionLineNumbering, pageNumberStart: b.props.sectionPageNumberStart };
        }
      }
      return { lineNumbering: refSectionOwn?.lineNumbering, pageNumberStart: refSectionOwn?.pageNumberStart };
    };
    /** A geometry-preserving break must still be materialized (not flowed) when the
     *  section it closes differs from the following section in a non-inheriting OWN
     *  property — otherwise dropping the boundary lets that property bleed backward. */
    const sectionOwnBleeds = (props: IRParaProps, bi: number): boolean => {
      const next = nextSectionOwn(bi);
      return (
        JSON.stringify(props.sectionLineNumbering ?? null) !== JSON.stringify(next.lineNumbering ?? null) ||
        (props.sectionPageNumberStart ?? null) !== (next.pageNumberStart ?? null)
      );
    };

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
      const mapped: Block[] =
        irBlock.kind === "paragraph"
          ? mapParagraph(irBlock, media, resolveLink)
          : irBlock.kind === "math"
            ? [{ kind: "equation", id: id(), revision: 0, equation: { root: irBlock.root, display: irBlock.display }, align: irBlock.align ?? "center", ...(irBlock.scale !== undefined ? { scale: irBlock.scale } : {}) }]
            : [mapTable(irBlock, media, resolveLink)];

      // Custom-field membership: an IR block in a field's result region stamps its
      // fieldId onto every model block it produced (a paragraph may split).
      if (irBlock.fieldId) for (const b of mapped) b.fieldId = irBlock.fieldId;

      // Block-level content-control ancestry: stamp every produced model block.
      if (irBlock.sdtPath) for (const b of mapped) b.sdtPath = irBlock.sdtPath;

      // TOC field anchor: record the real model block (first one this paragraph
      // produced) so a headless render places the built TOC exactly here.
      if (irBlock.kind === "paragraph" && irBlock.tocField !== undefined && tocAnchor === undefined) {
        const p = mapped.find((b): b is Paragraph => b.kind === "paragraph") ?? mapped[0];
        if (p) tocAnchor = { blockId: p.id, instruction: irBlock.tocField };
      }

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
        if (sectionIsDistinct(props) || sectionOwnBleeds(props, bi)) {
          // Apply the section's geometry: sectionBreak sits on the paragraph that
          // ENDS the section (engine pages at the next block).
          let carrier = lastParagraphOf(mapped);
          if (!carrier) {
            carrier = emptyParagraph();
            mapped.push(carrier);
          }
          const breakType: SectionBreakType = props.sectionBreakType ?? "nextPage";
          carrier.style.sectionBreak = { type: breakType, props: buildSectionPatch(props) };
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
          // …and break THIS closing section's own title when it's titled — covers the
          // heading-led page section the forward `pending` misses behind a preceding
          // continuous break. Never the document's first section (it can't break).
          if (sectionStartBi > 0 && sectionFirstVisible && isHeadingLedSection(sectionStartBi - 1)) {
            sectionFirstVisible.style.pageBreakBefore = true;
          }
        }
        seenPageSection = true;
      }
      // Section boundary (page OR continuous): the next block opens a new section, so
      // reset the title tracker; otherwise remember this section's first visible line.
      if (irBlock.kind === "paragraph" && irBlock.props.sectionBreak !== undefined) {
        sectionStartBi = bi + 1;
        sectionFirstVisible = null;
      } else if (!sectionFirstVisible) {
        sectionFirstVisible = firstVisiblePara(mapped) ?? null;
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
            // A soft break with nothing before it is an empty visual line ("a\n\nb").
            if (runs.length > 0) flushPara();
            else blocks.push(paraOf([]));
          } else {
            // A page/column break flushes any preceding text but emits NO empty line
            // of its own — the pending break rides the NEXT emitted paragraph. Pushing
            // a blank "before" line here strands an empty paragraph that, after a
            // full-page figure, becomes an entire blank page (the break-only paragraph
            // Word places between a full-page map and the following section).
            if (runs.length > 0) flushPara();
            if (inline.page) pendingPageBreak = true;
            if (inline.column) pendingColumnBreak = true;
          }
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
        case "shape": {
          const shape = mapShape(inline, style.align);
          if (shape) {
            flushPara();
            // A shape can't carry a page/column break — give the break a carrier.
            if (pendingPageBreak || pendingColumnBreak) blocks.push(paraOf([]));
            blocks.push(shape);
            trailingBreak = false;
          }
          break;
        }
        case "run": {
          const effective = resolver.run(ir.props.styleId, inline.props);
          // Hidden text (w:vanish) is PRESERVED with style.hidden (set via
          // applyRunProps) — never displayed, never caret-reachable, protected
          // from deletion — so a round-trip re-emits it instead of losing it.
          runs.push(mapRun(inline.text, effective, resolveLink, inline.sdtPath, inline.fieldId));
          trailingBreak = false;
          break;
        }
        case "mathInline": {
          // Inline equation → a single-U+FFFC run with CharStyle.equation. Pass
          // the inline props (e.g. linkRelId/linkAnchor from a w:hyperlink) so a
          // hyperlinked equation keeps its link via mapRun's resolveLink.
          const eqRun = mapRun("￼", resolver.run(ir.props.styleId, inline.props ?? {}), resolveLink);
          eqRun.style = { ...eqRun.style, equation: { root: inline.root, display: false } };
          runs.push(eqRun);
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
    // w14:paraId/textId belong to the source <w:p> — give them to its FIRST model
    // paragraph; a soft-break split leaves continuation pieces without an id.
    if (paragraphs[0]) {
      if (ir.paraId) paragraphs[0].paraId = ir.paraId;
      if (ir.textId) paragraphs[0].textId = ir.textId;
    }
    // Resolve this paragraph's bookmark markers to model positions. The home is
    // the first emitted block; offsets clamp to its text (exact for the common
    // single-block paragraph; a soft-break split anchors into the first block).
    if (ir.bookmarkMarkers && ir.bookmarkMarkers.length > 0 && blocks[0]) {
      const home = blocks[0]!;
      const homeLen = home.kind === "paragraph" ? home.runs.reduce((s, r) => s + r.text.length, 0) : 0;
      // Marker offsets were recorded in IR-inline coordinates, where a few inlines
      // occupy a different width than in the final model: a footnote/endnote
      // reference is an empty (0-width) IR run but paints its number in the model,
      // and an inline equation is 0-width in the IR but a single object char. Build
      // the cumulative IR→model width shift indexed BY INLINE POSITION so each
      // marker is shifted past exactly the zero-width inlines that precede it in XML
      // order — a marker sitting just before such an inline (same offset) stays put
      // while one just after it advances past the rendered char. shiftByIndex[k] is
      // the total expansion among the first k inlines. (footnoteNumber/endnoteNumber
      // were already assigned during the run walk above, so re-reading them here is
      // a cache hit — it does not bump the counter.)
      const shiftByIndex: number[] = [0];
      for (const inline of ir.inlines) {
        let delta = 0;
        if (inline.kind === "run") {
          if (inline.props.footnoteId !== undefined) delta = footnoteNumber(inline.props.footnoteId).length;
          else if (inline.props.endnoteId !== undefined) delta = endnoteNumber(inline.props.endnoteId).length;
        } else if (inline.kind === "mathInline") {
          delta = 1; // U+FFFC object-replacement char
        }
        shiftByIndex.push(shiftByIndex[shiftByIndex.length - 1]! + delta);
      }
      const shiftFor = (idx: number | undefined): number => shiftByIndex[Math.min(idx ?? 0, ir.inlines.length)]!;
      for (const m of ir.bookmarkMarkers) {
        const pos: DocPosition = { blockId: home.id, offset: Math.min(m.offset + shiftFor(m.inlineIndex), homeLen) };
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
    // Linked ("Link to File") image: keep the external URL as the portable handle
    // (src is a live-but-non-portable copy of it). Export re-emits r:link.
    const externalUrl = media.externalUrl(inline.relId);
    if (externalUrl) image.externalSrc = externalUrl;
    // a:srcRect crop insets (fractions), rounded for a stable round-trip.
    if (inline.crop) {
      image.crop = {
        left: round4(inline.crop.left), top: round4(inline.crop.top),
        right: round4(inline.crop.right), bottom: round4(inline.crop.bottom),
      };
    }
    // Anchored with square/tight wrap → an honest float; the model flows
    // following text around it.
    if (inline.anchored && inline.anchorWrap === "square") image.wrap = "square";
    // wrapNone anchor → absolutely-positioned behind/in-front image: no flow
    // height, no reflow. Mutually exclusive with `wrap` (parser sets at most one).
    if (inline.anchorFloat) {
      const a = inline.anchorFloat;
      image.anchor = {
        behind: a.behind,
        offsetXPx: round2(emuToPx(a.offsetXEmu)),
        offsetYPx: round2(emuToPx(a.offsetYEmu)),
        relFromH: a.relFromH,
        relFromV: a.relFromV,
        ...(a.decorative ? { decorative: true } : {}),
        ...(a.z !== undefined ? { z: a.z } : {}),
      };
    }
    // wp14:anchorId/editId — persistent drawing identity, preserved verbatim.
    if (inline.anchorId || inline.editId) {
      image.drawingId = {
        ...(inline.anchorId ? { anchorId: inline.anchorId } : {}),
        ...(inline.editId ? { editId: inline.editId } : {}),
      };
    }
    return image;
  }

  const KNOWN_PRESETS = new Set<ShapePreset>(["rect", "ellipse", "line"]);

  function mapShape(
    inline: Extract<IRInline, { kind: "shape" }>,
    paraAlign: ParaStyle["align"],
  ): ShapeBlock | undefined {
    if (inline.widthEmu === undefined || inline.heightEmu === undefined) {
      warnings.add("shapes-unsized", "A drawing shape without explicit dimensions was skipped.");
      return undefined;
    }
    // PR 1 models rect/ellipse/line; any other preset maps to rect (a box) so the
    // shape still round-trips — later PRs widen the preset set.
    let preset: ShapePreset = "rect";
    if (KNOWN_PRESETS.has(inline.preset as ShapePreset)) preset = inline.preset as ShapePreset;
    else warnings.add("shape-preset-unsupported", `An unsupported shape preset "${inline.preset}" was imported as a rectangle.`);
    const shape: ShapeBlock = {
      kind: "shape",
      id: id(),
      revision: 0,
      geometry: { preset },
      widthPx: round2(emuToPx(inline.widthEmu)),
      heightPx: round2(emuToPx(inline.heightEmu)),
      align: paraAlign === "justify" ? "left" : paraAlign,
    };
    if (inline.fill) shape.fill = inline.fill;
    if (inline.stroke) {
      shape.stroke = "none" in inline.stroke ? inline.stroke : { color: inline.stroke.color, widthPt: round2(inline.stroke.widthPt) };
    }
    // wp14:anchorId/editId — persistent drawing identity, preserved verbatim.
    if (inline.anchorId || inline.editId) {
      shape.drawingId = {
        ...(inline.anchorId ? { anchorId: inline.anchorId } : {}),
        ...(inline.editId ? { editId: inline.editId } : {}),
      };
    }
    return shape;
  }

  function mapRun(rawText: string, effective: IRRunProps, resolveLink: LinkResolver, sdtPath?: string[], fieldId?: string): Run {
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
    if (effective.endnoteId !== undefined) {
      // Same as footnotes, but the engine paints these at the document end.
      text = endnoteNumber(effective.endnoteId);
      style.endnoteRef = `en${effective.endnoteId}`;
      style.verticalAlign = "super";
    }
    if (sdtPath && sdtPath.length > 0) {
      style.sdtPath = sdtPath;
      // Checkbox glyphs arrive in symbol fonts (MS Gothic / Wingdings private
      // chars) — normalize to the Unicode glyphs the editor toggles between. A
      // checkbox is always a leaf control, so key on the innermost id.
      const sdt = sdts[sdtPath[sdtPath.length - 1]!];
      if (sdt?.type === "checkbox" && text.length > 0) {
        text = sdt.checked ? "☒" : "☐";
        style.fontFamily = DEFAULT_CHAR.fontFamily;
      }
    }
    if (fieldId) style.fieldId = fieldId; // inline built-in field result run
    // Character-style reference (w:rStyle) — kept for round-trip + "what style?"
    // UI; the concrete formatting is already baked above. collectUsedStyleIds
    // picks the same id up from the IR so the stylesheet builder includes it.
    if (effective.styleId) style.charStyleId = effective.styleId;
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
    // cell over table over table-style — but tri-state (issue #150): the most
    // specific layer that DEFINES shading wins, and an explicit clear (null) at
    // that layer suppresses the inherited fill instead of falling through. `??`
    // would skip a null, so pick the first DEFINED (non-undefined) layer.
    const shd = firstDefinedShd(p.ir?.shd, ir.shd, styled.shd);
    if (shd === null) p.cell.shadingCleared = true;
    else if (shd) p.cell.shading = shd;
    if (hasBorderInfo) {
      const src: BorderSources = { cell: p.ir?.borders, table: ir.borders, style: styled.borders };
      p.cell.borders = resolveCellBorders(src, pos);
    }
  }

  function mapTable(ir: IRTable, media: MediaStore, resolveLink: LinkResolver): TableBlock {
    const buildCell = (irCell: IRTableCell): TableCell => {
      const blocks: Block[] = [];
      for (const b of irCell.blocks) {
        const mapped: Block[] =
          b.kind === "paragraph"
            ? mapParagraph(b, media, resolveLink)
            : b.kind === "math"
              ? [{ kind: "equation", id: id(), revision: 0, equation: { root: b.root, display: b.display }, align: b.align ?? "center", ...(b.scale !== undefined ? { scale: b.scale } : {}) }]
              : [mapTable(b, media, resolveLink)]; // nested table — model renders one level
        // Block-level control inside a cell: stamp every produced model block.
        if (b.sdtPath) for (const mb of mapped) mb.sdtPath = b.sdtPath;
        blocks.push(...mapped);
      }
      const cell: TableCell = { id: id(), blocks: blocks.length > 0 ? blocks : [emptyParagraph()] };
      if (irCell.gridSpan > 1) cell.colSpan = irCell.gridSpan;
      // Cell-margin cascade: this cell's w:tcMar over the table's w:tblCellMar
      // over Word's defaults, resolved per side (a cell may set only top/bottom).
      const side = (s: "top" | "right" | "bottom" | "left"): number =>
        irCell.marginTwips?.[s] ?? ir.cellMarginTwips?.[s] ?? WORD_CELL_MARGIN_TWIPS[s];
      cell.margin = marginTwipsToPx({ top: side("top"), right: side("right"), bottom: side("bottom"), left: side("left") });
      if (irCell.preferredWidth) {
        cell.preferredWidth =
          irCell.preferredWidth.type === "pct"
            ? { type: "pct", px: irCell.preferredWidth.frac }
            : { type: "abs", px: round2(twipsToPx(irCell.preferredWidth.twips)) };
      }
      if (irCell.vAlign) cell.vAlign = irCell.vAlign;
      if (irCell.textDirection) cell.textDirection = irCell.textDirection;
      if (irCell.noWrap) cell.noWrap = true;
      if (irCell.fitText) cell.fitText = true;
      if (irCell.hideMark) cell.hideMark = true;
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

    const rows: TableRow[] = placedRows.map((row, ri) => {
      const out: TableRow = { cells: row.map((p) => p.cell) };
      const irProps = ir.rows[ri]?.props;
      if (irProps) {
        const props: RowProps = {};
        if (irProps.heightTwips !== undefined && irProps.heightRule) {
          props.height = { value: round2(twipsToPx(irProps.heightTwips)), rule: irProps.heightRule };
        }
        if (irProps.cantSplit) props.cantSplit = true;
        if (irProps.tblHeader) props.repeatHeader = true;
        if (Object.keys(props).length > 0) out.props = props;
      }
      return out;
    });
    const table: TableBlock = { kind: "table", id: id(), revision: 0, rows };
    // Table-style reference + active bands (the style itself goes to
    // Document.tableStyles via buildTableStyles; cells are baked above).
    if (ir.styleId) table.styleId = ir.styleId;
    if (ir.look) table.condOverrides = { ...ir.look };
    // Always emit colFractions of the true width: dropped continuation cells mean
    // a row's cell count no longer equals the column count, so the layout engine
    // can't infer the column count from cells.length anymore.
    table.colFractions = columnFractions(ir.colWidthsTwips, width);
    if (ir.widthMode && ir.widthMode !== "fixed") table.widthMode = ir.widthMode;
    if (ir.preferredWidthTwips) table.preferredWidth = { type: "px", value: round2(twipsToPx(ir.preferredWidthTwips)) };
    else if (ir.preferredWidthPct) table.preferredWidth = { type: "pct", value: ir.preferredWidthPct };
    if (ir.align && ir.align !== "left") table.align = ir.align;
    // Minor & advanced table props (issue #61): indent + RTL column order drive
    // layout; overlap/caption/description round-trip as metadata.
    if (ir.indentTwips) table.indentPx = round2(twipsToPx(ir.indentTwips));
    if (ir.bidiVisual) table.bidiVisual = true;
    if (ir.overlap) table.overlap = ir.overlap;
    if (ir.caption) table.caption = ir.caption;
    if (ir.description) table.description = ir.description;
    // Table-level defaults (w:tblBorders/w:shd/w:tblCellMar). These are cascaded
    // onto the cells above for layout/paint; retained here so export re-emits them
    // at tblPr level (issue #48) rather than only as the baked per-cell copies.
    const defBorders = tableBordersFromIR(ir.borders);
    if (defBorders) table.defaultBorders = defBorders;
    // An explicitly-declared but empty/all-none w:tblBorders ("no borders") still
    // round-trips its container: keep an empty CellBorders so export re-emits the
    // <w:tblBorders/> intent rather than dropping the table-level declaration.
    else if (ir.bordersSpecified) table.defaultBorders = {};
    if (ir.shd) table.defaultShading = ir.shd;
    if (ir.cellMarginTwips) {
      const m = ir.cellMarginTwips;
      table.defaultCellMargin = marginTwipsToPx({
        top: m.top ?? WORD_CELL_MARGIN_TWIPS.top,
        right: m.right ?? WORD_CELL_MARGIN_TWIPS.right,
        bottom: m.bottom ?? WORD_CELL_MARGIN_TWIPS.bottom,
        left: m.left ?? WORD_CELL_MARGIN_TWIPS.left,
      });
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
      marginPx: ir.marginTwips ? marginTwipsToPx(ir.marginTwips) : { ...DEFAULT_SECTION.marginPx },
    };
    if (ir.columns) section.columns = irColumnsToModel(ir.columns);
    if (ir.pageNumberStart !== undefined) section.pageNumberStart = ir.pageNumberStart;
    if (ir.headerDistTwips !== undefined) section.headerDistancePx = round2(twipsToPx(ir.headerDistTwips));
    if (ir.footerDistTwips !== undefined) section.footerDistancePx = round2(twipsToPx(ir.footerDistTwips));
    if (ir.pageBorders) section.pageBorders = irPageBordersToModel(ir.pageBorders);
    if (ir.lineNumbering) section.lineNumbering = irLineNumberingToModel(ir.lineNumbering);
    if (ir.breakType) section.breakType = ir.breakType;
    return section;
  }

  return {
    mapBlocks,
    mapSection,
    emptyParagraph,
    // Realize EVERY defined numbering definition (not just referenced ones) so a
    // list style authored in the editor survives save→reopen before it's applied.
    lists: () => {
      for (const numId of numbering.keys()) listDefFor(numId);
      return Object.fromEntries(usedLists);
    },
    bookmarks: () => {
      const out: Record<string, BookmarkRange> = {};
      for (const [id, s] of bookmarkStarts) {
        out[s.name] = { start: s.pos, end: bookmarkEnds.get(id) ?? s.pos }; // point bookmark if no end
      }
      return out;
    },
    footnoteRefs: () => footnoteOrder.map((docxId) => ({ docxId, noteId: `fn${docxId}` })),
    endnoteRefs: () => endnoteOrder.map((docxId) => ({ docxId, noteId: `en${docxId}` })),
    tocField: () => tocAnchor,
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

// Import-side (OOXML w:val → model enum) maps. Intentionally many-to-one
// (start/num→left, end→right, middleDot→dot, hyphen→dash) and NOT the inverse of
// the exporter's TAB_VAL/TAB_LEADER in export/docx/mappings.ts — do not merge.
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
    // An explicit clear (issue #154) is a position-only "remove" marker — no align/leader.
    if (t.val === "clear") return { posPx: round2(twipsToPx(t.posTwips)), cleared: true } as TabStop;
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
      ? (() => {
          const total = widthsTwips.reduce((s, x) => s + x, 0);
          return widthsTwips.map((w) => Math.round((w / total) * 10000) / 10000);
        })()
      : Array.from({ length: width }, () => Math.round((1 / width) * 10000) / 10000);
  fractions[fractions.length - 1] =
    Math.round((1 - fractions.slice(0, -1).reduce((s, f) => s + f, 0)) * 10000) / 10000;
  return fractions;
}

// ---------------------------------------------------------------------------
// Style mapping (pure)

function mapCharStyle(props: IRRunProps): CharStyle {
  const style: CharStyle = { ...DEFAULT_CHAR };
  applyRunProps(style, props);
  return style;
}

/** Fold the full OOXML w:u/@w:val vocabulary onto the small set the model paints.
 *  Heavy/long variants collapse onto their base; "single"/"words"/unknown ⇒
 *  undefined (a plain solid underline — the model's default). */
const UNDERLINE_STYLE_MAP: Record<string, UnderlineStyle> = {
  double: "double",
  thick: "thick",
  dotted: "dotted",
  dottedHeavy: "dotted",
  dash: "dash",
  dashedHeavy: "dash",
  dashLong: "dash",
  dashLongHeavy: "dash",
  dotDash: "dotDash",
  dashDotHeavy: "dotDash",
  dotDotDash: "dotDotDash",
  dashDotDotHeavy: "dotDotDash",
  wave: "wave",
  wavyHeavy: "wave",
  wavyDouble: "wave",
};
function mapUnderlineStyle(raw: string): UnderlineStyle | undefined {
  return UNDERLINE_STYLE_MAP[raw];
}

/** Shared run-property mapping for full CharStyle and partial style-gallery patches. */
function applyRunProps(style: Partial<CharStyle>, props: IRRunProps): void {
  // ascii is the Latin slot the model lays out; hAnsi is its high-ANSI twin (used
  // when ascii is absent). cs/eastAsia are preserved for round-trip only, and only
  // when they name a DISTINCT face: Word writes w:cs (and often w:eastAsia) equal to
  // w:ascii for plain Latin runs, which carries no extra information.
  const latinFont = props.fontAscii ?? props.fontHAnsi;
  if (props.fontAscii) style.fontFamily = `${props.fontAscii}, serif`;
  else if (props.fontHAnsi) style.fontFamily = `${props.fontHAnsi}, serif`;
  if (props.fontCs && props.fontCs !== latinFont) style.fontFamilyComplexScript = `${props.fontCs}, serif`;
  if (props.fontEastAsia && props.fontEastAsia !== latinFont) style.fontFamilyEastAsia = `${props.fontEastAsia}, serif`;
  if (props.letterSpacingTwips !== undefined) style.letterSpacingPx = round2(twipsToPx(props.letterSpacingTwips));
  if (props.sizeHalfPoints !== undefined) style.fontSizePx = round2(halfPointsToPx(props.sizeHalfPoints));
  if (props.bold !== undefined) style.bold = props.bold;
  if (props.italic !== undefined) style.italic = props.italic;
  if (props.underline !== undefined) style.underline = props.underline;
  if (props.underline) {
    const us = props.underlineStyle ? mapUnderlineStyle(props.underlineStyle) : undefined;
    if (us) style.underlineStyle = us;
    if (props.underlineColor && props.underlineColor !== "auto") style.underlineColor = `#${props.underlineColor.toLowerCase()}`;
  }
  if (props.strikethrough !== undefined) style.strikethrough = props.strikethrough;
  if (props.color && props.color !== "auto") style.color = `#${props.color.toLowerCase()}`;
  // Tri-state (issue #155): null = explicit "No Color" clear of an inherited char-style
  // highlight → marker (omitting highlightColor); a name → resolved hex.
  if (props.highlight === null) style.highlightCleared = true;
  else if (props.highlight) {
    const hex = HIGHLIGHT_HEX[props.highlight];
    if (hex) style.highlightColor = hex;
  }
  if (props.vertAlign === "superscript") style.verticalAlign = "super";
  else if (props.vertAlign === "subscript") style.verticalAlign = "sub";
  if (props.vanish) style.hidden = true; // preserved, never displayed (see CharStyle.hidden)
  if (props.caps !== undefined) style.caps = props.caps; // w:caps — all-caps display
  if (props.smallCaps !== undefined) style.smallCaps = props.smallCaps; // w:smallCaps — small capitals
  if (props.rtl !== undefined) style.rtl = props.rtl; // explicit w:rtl="0" clears inherited RTL
  // Minor run typography & effects (w:rPr extras): unit-convert and collapse onto
  // the model's CharStyle fields. Each is additive — only set when present.
  if (props.doubleStrikethrough !== undefined) style.doubleStrikethrough = props.doubleStrikethrough;
  if (props.positionHalfPoints !== undefined) style.positionPx = round2(halfPointsToPx(props.positionHalfPoints));
  if (props.kerningHalfPoints !== undefined) style.kerningMinPx = round2(halfPointsToPx(props.kerningHalfPoints));
  if (props.widthScalePct !== undefined) style.widthScalePct = props.widthScalePct;
  if (props.emphasisMark) {
    const em = mapEmphasisMark(props.emphasisMark);
    if (em) style.emphasisMark = em;
  }
  if (props.outline !== undefined) style.outline = props.outline;
  if (props.shadow !== undefined) style.shadow = props.shadow;
  if (props.emboss !== undefined) style.emboss = props.emboss;
  if (props.imprint !== undefined) style.imprint = props.imprint;
  if (props.snapToGrid !== undefined) style.snapToGrid = props.snapToGrid;
  if (props.lang) style.lang = props.lang;
  if (props.runBorder) {
    const b = runBorderFromIR(props.runBorder);
    if (b) style.runBorder = b;
  }
  if (props.fitTextTwips !== undefined && props.fitTextTwips > 0) style.fitTextPx = round2(twipsToPx(props.fitTextTwips));
  if (props.symbol) {
    // Symbol glyph (w:sym): keep the marker for round-trip AND set the font so the
    // decoded code point paints in the symbol face (e.g. Wingdings), not body text.
    style.symbol = { font: props.symbol.font, char: props.symbol.char };
    style.fontFamily = `${props.symbol.font}, sans-serif`;
  }
}

/** Fold the OOXML w:em vocabulary onto the model's set; unknown values drop. */
const EMPHASIS_MARK_MAP: Record<string, EmphasisMark> = {
  dot: "dot",
  comma: "comma",
  circle: "circle",
  underDot: "underDot",
};
function mapEmphasisMark(raw: string): EmphasisMark | undefined {
  return EMPHASIS_MARK_MAP[raw];
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
    if (p.checkedSymbol !== undefined) props.checkedSymbol = p.checkedSymbol;
    if (p.uncheckedSymbol !== undefined) props.uncheckedSymbol = p.uncheckedSymbol;
    if (p.lockContent) props.lockContent = true;
    if (p.lockControl) props.lockControl = true;
    out[id] = props;
  }
  return out;
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
  applyParaProps(out, props);
  return out;
}

/** Shared paragraph-property mapping for full ParaStyle and partial style-gallery
 *  patches — mirrors applyRunProps for w:rPr. The full mapper layers defaults +
 *  its three extra fields (pageBreakBefore/outlineLevel/namedStyle) on top. */
function applyParaProps(out: Partial<ParaStyle>, props: IRParaProps): void {
  if (props.align) out.align = props.align;
  if (props.direction) out.direction = props.direction;
  if (props.lineHeight !== undefined) out.lineHeight = round2(props.lineHeight);
  // "auto" carries a multiplier and clears any fixed rule; exact/atLeast set one.
  if (props.lineRule === "exact" || props.lineRule === "atLeast") {
    out.lineRule = props.lineRule;
    if (props.lineExactTwips !== undefined) out.lineHeightPx = round2(twipsToPx(props.lineExactTwips));
  }
  if (props.spaceBeforeTwips !== undefined) out.spaceBeforePx = round2(twipsToPx(props.spaceBeforeTwips));
  if (props.spaceAfterTwips !== undefined) out.spaceAfterPx = round2(twipsToPx(props.spaceAfterTwips));
  // Auto-spacing (issue #160): bake a concrete approximation into the px field (Word
  // ignores the explicit value when autospacing is on) and keep the flag for export.
  // A per-context auto value + HTML-style margin collapsing are not reproduced.
  if (props.spaceBeforeAuto !== undefined) {
    out.spaceBeforeAuto = props.spaceBeforeAuto;
    if (props.spaceBeforeAuto) out.spaceBeforePx = AUTO_PARA_SPACING_PX;
  }
  if (props.spaceAfterAuto !== undefined) {
    out.spaceAfterAuto = props.spaceAfterAuto;
    if (props.spaceAfterAuto) out.spaceAfterPx = AUTO_PARA_SPACING_PX;
  }
  if (props.indentLeftTwips !== undefined) out.indentLeftPx = round2(twipsToPx(props.indentLeftTwips));
  if (props.indentRightTwips !== undefined) out.indentRightPx = round2(twipsToPx(props.indentRightTwips));
  if (props.indentFirstLineTwips !== undefined)
    out.indentFirstLinePx = round2(twipsToPx(props.indentFirstLineTwips));
  if (props.keepWithNext) out.keepWithNext = true;
  if (props.keepLinesTogether) out.keepLinesTogether = true;
  // Keep explicit false (w:contextualSpacing w:val="0") so a paragraph/style can
  // clear an inherited `true` through the cascade.
  if (props.contextualSpacing !== undefined) out.contextualSpacing = props.contextualSpacing;
  if (props.tabStops) out.tabStops = mapTabStops(props.tabStops);
  const pb = paraBordersFromIR(props.borders);
  // Tri-state (issue #153): real edges → borders; a w:pBdr present with no drawable
  // edge (empty / all-nil) → an explicit clear marker so export re-emits the empty
  // w:pBdr against a bordered named style instead of the style's box returning.
  if (pb) out.borders = pb;
  else if (props.bordersSpecified) out.bordersCleared = true;
  // w:shd tri-state (issue #147): a `null` resolved fill is an explicit clear —
  // record the marker and omit `shading` (which stays a concrete-fill-only field)
  // so the export can re-emit `<w:shd w:val="clear" w:fill="auto"/>` against a
  // named style that carries a fill. A concrete fill maps to `shading` as before.
  if (props.shd === null) out.shadingCleared = true;
  else if (props.shd !== undefined) out.shading = props.shd;
  mapMinorParaProps(props, out);
}

/** Minor w:pPr props (issue #62) shared by the full-paragraph and style-patch
 *  mappers. Each is a direct passthrough (no unit conversion). */
function mapMinorParaProps(props: IRParaProps, out: Partial<ParaStyle>): void {
  if (props.widowControl !== undefined) out.widowControl = props.widowControl;
  if (props.suppressLineNumbers !== undefined) out.suppressLineNumbers = props.suppressLineNumbers;
  if (props.textAlignment !== undefined) out.textAlignment = props.textAlignment;
  if (props.mirrorIndents !== undefined) out.mirrorIndents = props.mirrorIndents;
  if (props.adjustRightInd !== undefined) out.adjustRightInd = props.adjustRightInd;
  // Unmodeled CJK / hyphenation toggles (issue #161) — round-trip-only, no layout effect.
  if (props.snapToGrid !== undefined) out.snapToGrid = props.snapToGrid;
  if (props.suppressAutoHyphens !== undefined) out.suppressAutoHyphens = props.suppressAutoHyphens;
  if (props.kinsoku !== undefined) out.kinsoku = props.kinsoku;
  if (props.overflowPunct !== undefined) out.overflowPunct = props.overflowPunct;
  if (props.wordWrap !== undefined) out.wordWrap = props.wordWrap;
  if (props.topLinePunct !== undefined) out.topLinePunct = props.topLinePunct;
  if (props.autoSpaceDE !== undefined) out.autoSpaceDE = props.autoSpaceDE;
  if (props.autoSpaceDN !== undefined) out.autoSpaceDN = props.autoSpaceDN;
}

/** styles.xml → editor Stylesheet. ALL defined paragraph and character styles
 *  are carried over (not just the ones in use) so a style authored in the editor
 *  survives a save→reopen even before it's applied to any content — matching
 *  Word, which keeps style definitions regardless of use. Display names come from
 *  w:name; generated documents use opaque numeric styleIds. */
export function buildStylesheet(data: StylesData): Stylesheet | undefined {
  if (data.styles.size === 0) return undefined;
  const include = new Set<string>();
  for (const def of data.styles.values()) {
    if (def.type === "paragraph" || def.type === "character") include.add(def.id);
  }
  if (include.size === 0) return undefined;

  const styles: NamedStyle[] = [];
  for (const id of include) {
    const def = data.styles.get(id)!;
    const isChar = def.type === "character";
    const style: NamedStyle = {
      id,
      name: def.name ?? id,
      type: isChar ? "character" : "paragraph",
      char: mapCharPatch(def.rPr),
      para: isChar ? {} : mapParaPatch(def.pPr),
    };
    // basedOn is kept only when it resolves to an included style of the same kind.
    if (def.basedOnId && include.has(def.basedOnId) && data.styles.get(def.basedOnId)?.type === def.type) {
      style.basedOn = def.basedOnId;
    }
    styles.push(style);
  }
  // Default style first — the gallery reads top-to-bottom.
  styles.sort((a, b) =>
    a.id === data.defaultParaStyleId ? -1 : b.id === data.defaultParaStyleId ? 1 : a.name.localeCompare(b.name),
  );
  return { styles, defaultStyleId: data.defaultParaStyleId ?? styles[0]!.id };
}

/** styles.xml table styles → Document.tableStyles. ALL defined table styles are
 *  carried over (not just used ones) so an authored table style survives a
 *  save→reopen before being applied. Each w:tblStylePr conditional band maps to a
 *  TableCond; the whole-table base comes from the style's own rPr/pPr/tblBorders/tblShd. */
export function buildTableStyles(data: StylesData): Record<string, TableStyle> | undefined {
  const include = new Set<string>();
  for (const def of data.styles.values()) if (def.type === "table") include.add(def.id);
  if (include.size === 0) return undefined;

  const condProps = (rPr: IRRunProps, pPr: IRParaProps, borders: IRBorders | undefined, shd: string | undefined): TableCondProps => {
    const cp: TableCondProps = {};
    const char = mapCharPatch(rPr);
    if (Object.keys(char).length > 0) cp.char = char;
    const para = mapParaPatch(pPr);
    if (Object.keys(para).length > 0) cp.para = para;
    const cb = cellBordersFromIR(borders);
    if (cb) cp.borders = cb;
    if (shd !== undefined) cp.shading = shd;
    return cp;
  };

  const out: Record<string, TableStyle> = {};
  for (const id of include) {
    const def = data.styles.get(id)!;
    const conds: Partial<Record<TableCond, TableCondProps>> = {
      wholeTable: condProps(def.rPr, def.pPr, def.tblBorders, def.tblShd),
    };
    for (const [cond, c] of Object.entries(def.tblStylePr ?? {})) {
      conds[cond as TableCond] = condProps(c.rPr, c.pPr, c.borders, c.shd);
    }
    const style: TableStyle = { id, name: def.name ?? id, conds };
    if (def.basedOnId && include.has(def.basedOnId)) style.basedOn = def.basedOnId;
    if (def.rowBandSize !== undefined) style.rowBandSize = def.rowBandSize;
    if (def.colBandSize !== undefined) style.colBandSize = def.colBandSize;
    out[id] = style;
  }
  return out;
}

function mapParaStyle(props: IRParaProps): ParaStyle {
  const style: ParaStyle = { ...DEFAULT_PARA };
  applyParaProps(style, props);
  if (props.pageBreakBefore) style.pageBreakBefore = true;
  if (props.outlineLevel !== undefined) style.outlineLevel = props.outlineLevel;
  if (props.styleId) style.namedStyle = props.styleId;
  return style;
}
