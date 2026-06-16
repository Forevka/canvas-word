// Root of the fluent document-builder API. Composes the shared block-scope
// surface (body content, via StoryBuilder) with document-level concerns:
// page setup, header/footer bands, named-style registration, and build().
//
// The builder is a thin layer that mints plain model data — the same
// Document the editor, exporter, and collaboration layer consume. There is
// deliberately no hidden state: build() returns a deep clone, and the builder
// stays usable afterwards (rebuild-on-data-change just calls the author's
// function again with fresh data).

import type { BandContainer, Block, Document, ListLevel, ListNumberFormat, NamedStyle, SectionPatch, SectionProps, Stylesheet, TocOptions, TocSwitches } from "@cw/shared";
import { buildTocInstruction, bulletListDefinition, defaultStylesheet, generateTocIntoDoc, multilevelListDefinition, numberListDefinition, styleById } from "@cw/shared";
import { BuilderContext, type BuilderWarning } from "./blockFactory";
import { StoryBuilder } from "./storyBuilder";
import type { TableStylePreset } from "./tableStyles";
import { prepareTemplate } from "./template";
import { PAGE_SIZES, type PageSize, type PageSizeName } from "./units";

export interface PageSetup {
  pageSize?: PageSizeName | PageSize;
  /** Swaps width/height of the effective page size when needed. */
  orientation?: "portrait" | "landscape";
  /** Per-side margins, px; omitted sides keep their current value. */
  margins?: Partial<SectionProps["marginPx"]>;
  /** Newspaper columns (gapPx defaults to 48 ≈ Word's 0.5in). */
  columns?: { count: number; gapPx?: number };
  headerDistancePx?: number;
  footerDistancePx?: number;
  pageNumberStart?: number;
}

export interface CreateOptions {
  pageSize?: PageSizeName | PageSize;
  margins?: Partial<SectionProps["marginPx"]>;
  /** Start from a custom stylesheet (default: the editor's default gallery). */
  stylesheet?: Stylesheet;
  /** Deterministic id namespace (tests); default is a random per-builder seed. */
  idSeed?: string;
}

export interface TemplateOptions {
  /** Keep the template's body content and append after it (default: the body
   *  is discarded — the template contributes styles, lists, page setup, bands). */
  keepBody?: boolean;
  idSeed?: string;
}

export interface BandOptions {
  /** Which page band variant: default (all pages), first page, or even pages. */
  variant?: "default" | "first" | "even";
}

/** A `\nextPage` section break's per-section overrides. Geometry resolves like
 *  pageSetup; band callbacks compile to the section's own header/footer stories. */
export interface SectionBreakOptions {
  pageSize?: PageSizeName | PageSize;
  orientation?: "portrait" | "landscape";
  margins?: Partial<SectionProps["marginPx"]>;
  /** Newspaper columns for the new section; null = single column. */
  columns?: { count: number; gapPx?: number } | null;
  pageNumberStart?: number;
  header?: (s: StoryBuilder) => void;
  footer?: (s: StoryBuilder) => void;
  headerFirst?: (s: StoryBuilder) => void;
  footerFirst?: (s: StoryBuilder) => void;
  headerEven?: (s: StoryBuilder) => void;
  footerEven?: (s: StoryBuilder) => void;
}

/** How to build a custom list definition for .listDefinition(id, spec). */
export type ListDefinitionSpec =
  | { kind: "bullet"; char: string }
  | { kind: "number"; format: ListNumberFormat; suffix?: string }
  | { kind: "multilevel" }
  | { levels: ListLevel[] };

const DEFAULT_MARGIN = 96; // 1in @96dpi

function resolvePageSize(size: PageSizeName | PageSize): PageSize {
  return typeof size === "string" ? PAGE_SIZES[size] : size;
}

export class DocumentBuilder extends StoryBuilder {
  /** Start from a blank document (default stylesheet, US Letter, 1in margins). */
  static create(opts: CreateOptions = {}): DocumentBuilder {
    const size = resolvePageSize(opts.pageSize ?? "Letter");
    const doc: Document = {
      section: {
        pageWidthPx: size.pageWidthPx,
        pageHeightPx: size.pageHeightPx,
        marginPx: { top: DEFAULT_MARGIN, right: DEFAULT_MARGIN, bottom: DEFAULT_MARGIN, left: DEFAULT_MARGIN, ...opts.margins },
      },
      blocks: [],
      stylesheet: opts.stylesheet ?? defaultStylesheet(),
    };
    return new DocumentBuilder(new BuilderContext(doc, opts.idSeed));
  }

  /** Start from a .docx template: its stylesheet, list definitions, page setup,
   *  and header/footer bands carry over (body content is discarded unless
   *  keepBody). Import warnings surface on `builder.warnings`. */
  static async fromTemplate(docx: ArrayBuffer | Uint8Array, opts: TemplateOptions = {}): Promise<DocumentBuilder> {
    const { doc, warnings } = prepareTemplate(docx, opts.keepBody ?? false);
    const builder = new DocumentBuilder(new BuilderContext(doc, opts.idSeed));
    for (const w of warnings) builder.ctx.warn(w.code, w.message);
    return builder;
  }

  /** Set when .tableOfContents() ran — build() splices entries at the anchor. */
  private tocRequested = false;
  private tocOpts: TocOptions = {};

  private constructor(ctx: BuilderContext) {
    super(ctx, ctx.doc.blocks);
  }

  /** Lossy decisions made while building (unknown style ids, template import
   *  warnings, …) — surfaced, not swallowed. */
  get warnings(): readonly BuilderWarning[] {
    return this.ctx.warnings;
  }

  /** Compose a header band. Replaces the variant's existing story (including
   *  one carried over from the template). */
  header(build: (s: StoryBuilder) => void, opts: BandOptions = {}): this {
    return this.band("header", build, opts);
  }

  /** Compose a footer band. Same variant semantics as header(). */
  footer(build: (s: StoryBuilder) => void, opts: BandOptions = {}): this {
    return this.band("footer", build, opts);
  }

  private band(kind: "header" | "footer", build: (s: StoryBuilder) => void, opts: BandOptions): this {
    const variant = opts.variant ?? "default";
    const container: BandContainer =
      variant === "default" ? kind : (`${kind}${variant === "first" ? "First" : "Even"}` as BandContainer);
    const blocks: Block[] = [];
    build(new StoryBuilder(this.ctx, blocks));
    if (blocks.length === 0) blocks.push(this.ctx.paragraph([]));
    this.ctx.doc.section[container] = blocks;
    return this;
  }

  /** Merge page geometry onto the section (template values are the base). */
  pageSetup(setup: PageSetup): this {
    const section = this.ctx.doc.section;
    if (setup.pageSize !== undefined) {
      const size = resolvePageSize(setup.pageSize);
      section.pageWidthPx = size.pageWidthPx;
      section.pageHeightPx = size.pageHeightPx;
    }
    if (setup.orientation !== undefined) {
      const landscape = setup.orientation === "landscape";
      const [w, h] = [section.pageWidthPx, section.pageHeightPx];
      if (landscape !== w > h) {
        section.pageWidthPx = h;
        section.pageHeightPx = w;
      }
    }
    if (setup.margins !== undefined) section.marginPx = { ...section.marginPx, ...setup.margins };
    if (setup.columns !== undefined) section.columns = { count: setup.columns.count, gapPx: setup.columns.gapPx ?? 48 };
    if (setup.headerDistancePx !== undefined) section.headerDistancePx = setup.headerDistancePx;
    if (setup.footerDistancePx !== undefined) section.footerDistancePx = setup.footerDistancePx;
    if (setup.pageNumberStart !== undefined) section.pageNumberStart = setup.pageNumberStart;
    return this;
  }

  /** Register (or override) a named style. Register BEFORE applying it —
   *  withStyle() resolves at call time, not at build(). */
  style(def: NamedStyle): this {
    const sheet = this.ctx.stylesheet();
    const i = sheet.styles.findIndex((s) => s.id === def.id);
    if (i >= 0) sheet.styles[i] = def;
    else sheet.styles.push(def);
    return this;
  }

  /** Make `id` the document's default style and refresh the run/paragraph
   *  defaults from it. Call EARLY — only content added afterwards inherits it. */
  defaultStyle(id: string): this {
    const sheet = this.ctx.stylesheet();
    if (!styleById(sheet, id)) {
      this.ctx.warn(`style-missing:${id}`, `.defaultStyle("${id}") — no such style in the stylesheet; default unchanged.`);
      return this;
    }
    sheet.defaultStyleId = id;
    this.ctx.refreshDefaults();
    return this;
  }

  /** Register a custom list definition; reference it via .list(items,{listId:id}). */
  listDefinition(id: string, spec: ListDefinitionSpec): this {
    const lists = (this.ctx.doc.lists ??= {});
    if ("levels" in spec) lists[id] = { id, levels: spec.levels };
    else if (spec.kind === "bullet") lists[id] = bulletListDefinition(id, spec.char);
    else if (spec.kind === "number") lists[id] = numberListDefinition(id, spec.format, spec.suffix ?? ".");
    else lists[id] = { ...multilevelListDefinition(), id };
    return this;
  }

  /** Register (or override) a builder-only table-style preset — reusable cell
   *  formatting applied via .table(rows, { style: name }). Resolved to concrete
   *  cell properties at build time; nothing style-id-like enters the model. */
  tableStylePreset(name: string, preset: TableStylePreset): this {
    this.ctx.tableStyles.set(name, preset);
    return this;
  }

  /** Insert a table of contents, built from the document's headings at build()
   *  time (so headings added AFTER this call are still included). Pushes a
   *  placeholder anchor now; build() replaces it with the entries. */
  tableOfContents(opts: TocOptions = {}): this {
    this.tocRequested = true;
    this.tocOpts = opts;
    const sw: TocSwitches = {
      useOutlineLevels: true,
      hyperlinks: opts.hyperlink ?? true,
      hideInWeb: true,
      outlineRange: { from: 1, to: opts.maxLevel ?? 3 },
    };
    this.ctx.doc.tocInstruction = buildTocInstruction(sw);
    const anchor = this.ctx.paragraph([]);
    this.push(anchor);
    this.ctx.doc.tocAnchorBlockId = anchor.id;
    return this;
  }

  /** End the current section and start a new one on the next page, with the
   *  given per-section overrides (page size/orientation/margins, newspaper
   *  columns, page-number restart, and its own header/footer bands). */
  sectionBreak(opts: SectionBreakOptions = {}): this {
    const patch = this.compileSectionPatch(opts);
    const blocks = this.ctx.doc.blocks;
    const last = blocks[blocks.length - 1];
    if (last && last.kind === "paragraph") last.style.sectionBreak = { type: "nextPage", props: patch };
    else blocks.push(this.ctx.paragraph([], { sectionBreak: { type: "nextPage", props: patch } }));
    return this;
  }

  private compileSectionPatch(opts: SectionBreakOptions): SectionPatch {
    const patch: SectionPatch = {};
    const section = this.ctx.doc.section;
    if (opts.pageSize !== undefined || opts.orientation !== undefined) {
      let w = section.pageWidthPx;
      let h = section.pageHeightPx;
      if (opts.pageSize !== undefined) {
        const s = resolvePageSize(opts.pageSize);
        w = s.pageWidthPx;
        h = s.pageHeightPx;
      }
      if (opts.orientation !== undefined && (opts.orientation === "landscape") !== w > h) [w, h] = [h, w];
      patch.pageWidthPx = w;
      patch.pageHeightPx = h;
    }
    if (opts.margins !== undefined) patch.marginPx = { ...section.marginPx, ...opts.margins };
    if (opts.columns !== undefined) patch.columns = opts.columns === null ? null : { count: opts.columns.count, gapPx: opts.columns.gapPx ?? 48 };
    if (opts.pageNumberStart !== undefined) patch.pageNumberStart = opts.pageNumberStart;
    const band = (cb?: (s: StoryBuilder) => void): Block[] | undefined => {
      if (!cb) return undefined;
      const blocks: Block[] = [];
      cb(new StoryBuilder(this.ctx, blocks));
      if (blocks.length === 0) blocks.push(this.ctx.paragraph([]));
      return blocks;
    };
    const hd = band(opts.header);
    if (hd) patch.header = hd;
    const ft = band(opts.footer);
    if (ft) patch.footer = ft;
    const hf = band(opts.headerFirst);
    if (hf) patch.headerFirst = hf;
    const ff = band(opts.footerFirst);
    if (ff) patch.footerFirst = ff;
    const he = band(opts.headerEven);
    if (he) patch.headerEven = he;
    const fe = band(opts.footerEven);
    if (fe) patch.footerEven = fe;
    return patch;
  }

  /** Snapshot the document: deep clone, guaranteed at least one paragraph.
   *  The builder remains usable (and unaffected by mutations of the result). */
  build(): Document {
    this.flushPendingPageBreak();
    let doc = structuredClone(this.ctx.doc);
    // Resolve a requested TOC against the FINAL set of headings (splices entries
    // at the anchor + replaces it). Runs on the clone, so the live doc keeps the
    // anchor and a second build() regenerates cleanly (idempotent).
    if (this.tocRequested && doc.tocAnchorBlockId !== undefined) {
      doc = generateTocIntoDoc(doc, this.tocOpts).doc;
    }
    // The caret-home guard goes on the SNAPSHOT — building an empty document
    // must not leave a stray paragraph in the builder's live block list.
    if (doc.blocks.length === 0) doc.blocks.push(this.ctx.paragraph([]));
    return doc;
  }
}
