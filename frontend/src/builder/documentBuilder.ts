// Root of the fluent document-builder API. Composes the shared block-scope
// surface (body content, via StoryBuilder) with document-level concerns:
// page setup, header/footer bands, named-style registration, and build().
//
// The builder is a thin layer that mints plain model data — the same
// Document the editor, exporter, and collaboration layer consume. There is
// deliberately no hidden state: build() returns a deep clone, and the builder
// stays usable afterwards (rebuild-on-data-change just calls the author's
// function again with fresh data).

import type { BandContainer, Block, Document, NamedStyle, SectionProps, Stylesheet } from "@cw/shared";
import { defaultStylesheet } from "@cw/shared";
import { BuilderContext, type BuilderWarning } from "./blockFactory";
import { StoryBuilder } from "./storyBuilder";
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

  /** Snapshot the document: deep clone, guaranteed at least one paragraph.
   *  The builder remains usable (and unaffected by mutations of the result). */
  build(): Document {
    this.flushPendingPageBreak();
    const doc = structuredClone(this.ctx.doc);
    // The caret-home guard goes on the SNAPSHOT — building an empty document
    // must not leave a stray paragraph in the builder's live block list.
    if (doc.blocks.length === 0) doc.blocks.push(this.ctx.paragraph([]));
    return doc;
  }
}
