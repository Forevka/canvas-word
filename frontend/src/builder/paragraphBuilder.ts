// Paragraph scope. Returned by StoryBuilder.paragraph(); styles the paragraph
// it wraps, and re-implements the block-starting surface by delegating to its
// parent scope — so `.paragraph(..).withStyle(..).paragraph(..)` pops back up
// without an explicit end-of-scope call.
//
// Character formatting semantics: bold()/color()/font()/… patch EVERY run
// currently in the paragraph AND become the default for runs added later by
// text() in this scope — "make this paragraph bold" is the dominant authoring
// intent. Mixed formatting within a paragraph uses text(t, { …patch }).

import type { CharStyle, Document, NamedStyle, ParaStyle, Paragraph } from "@cw/shared";
import type { BuilderContext } from "./blockFactory";
import type { BandOptions, DocumentBuilder, PageSetup } from "./documentBuilder";
import type { ImageOptions, ListItem, ListOptions, StoryBuilder } from "./storyBuilder";
import type { CellContent, TableBuilder, TableOptions } from "./tableBuilder";

export interface SpacingOptions {
  /** Space before the paragraph, px. */
  before?: number;
  /** Space after the paragraph, px. */
  after?: number;
  /** Line height multiplier (1.0 = single). */
  lineHeight?: number;
}

export interface IndentOptions {
  /** Left indent, px. */
  left?: number;
  /** Right indent, px. */
  right?: number;
  /** First-line indent, px (negative = hanging). */
  firstLine?: number;
}

export class ParagraphBuilder<P extends StoryBuilder> {
  /** Char formatting accrued in this scope — future text() runs inherit it. */
  private charPatch: Partial<CharStyle> = {};

  constructor(
    private readonly parent: P,
    private readonly ctx: BuilderContext,
    private readonly para: Paragraph,
  ) {}

  // ---- paragraph scope -----------------------------------------------------

  /** Apply a named style from the stylesheet: sets the reference and patches
   *  exactly the fields the style defines (direct formatting applied AFTER this
   *  call wins — call order is precedence, like applying a style in the editor). */
  withStyle(id: string): this {
    const resolved = this.ctx.lookupStyle(id);
    if (!resolved) return this; // unknown style → warning already recorded
    this.para.style.namedStyle = id;
    Object.assign(this.para.style, resolved.para);
    for (const r of this.para.runs) Object.assign(r.style, resolved.char);
    Object.assign(this.charPatch, resolved.char);
    return this;
  }

  /** Append a run. Inherits the scope's accrued char formatting; `style` wins. */
  text(text: string, style?: Partial<CharStyle>): this {
    const run = this.ctx.run(text, { ...this.charPatch, ...style });
    // A paragraph created without text holds one empty placeholder run (the
    // model's empty-paragraph shape) — the first real text replaces it.
    if (this.para.runs.length === 1 && this.para.runs[0]!.text === "") this.para.runs[0] = run;
    else this.para.runs.push(run);
    return this;
  }

  private applyChar(patch: Partial<CharStyle>): this {
    Object.assign(this.charPatch, patch);
    for (const r of this.para.runs) Object.assign(r.style, patch);
    return this;
  }

  bold(on = true): this {
    return this.applyChar({ bold: on });
  }

  italic(on = true): this {
    return this.applyChar({ italic: on });
  }

  underline(on = true): this {
    return this.applyChar({ underline: on });
  }

  strikethrough(on = true): this {
    return this.applyChar({ strikethrough: on });
  }

  color(cssColor: string): this {
    return this.applyChar({ color: cssColor });
  }

  highlight(cssColor: string): this {
    return this.applyChar({ highlightColor: cssColor });
  }

  fontSize(px: number): this {
    return this.applyChar({ fontSizePx: px });
  }

  font(family: string): this {
    return this.applyChar({ fontFamily: family });
  }

  /** Make the paragraph's text a hyperlink (painted blue+underlined, Ctrl+click). */
  link(url: string): this {
    return this.applyChar({ link: url });
  }

  align(align: ParaStyle["align"]): this {
    this.para.style.align = align;
    return this;
  }

  spacing(opts: SpacingOptions): this {
    if (opts.before !== undefined) this.para.style.spaceBeforePx = opts.before;
    if (opts.after !== undefined) this.para.style.spaceAfterPx = opts.after;
    if (opts.lineHeight !== undefined) this.para.style.lineHeight = opts.lineHeight;
    return this;
  }

  indent(opts: IndentOptions): this {
    if (opts.left !== undefined) this.para.style.indentLeftPx = opts.left;
    if (opts.right !== undefined) this.para.style.indentRightPx = opts.right;
    if (opts.firstLine !== undefined) this.para.style.indentFirstLinePx = opts.firstLine;
    return this;
  }

  keepWithNext(on = true): this {
    this.para.style.keepWithNext = on;
    return this;
  }

  /** Escape the paragraph scope explicitly (rarely needed — any block-starting
   *  call below does it implicitly). */
  end(): P {
    return this.parent;
  }

  // ---- delegated block-starting surface (pops this scope) -------------------

  paragraph(text?: string, style?: Partial<CharStyle>): ParagraphBuilder<P> {
    return this.parent.paragraph(text, style) as ParagraphBuilder<P>;
  }

  table(rows: CellContent[][], opts?: TableOptions): P;
  table(build: (t: TableBuilder) => void, opts?: TableOptions): P;
  table(arg: CellContent[][] | ((t: TableBuilder) => void), opts?: TableOptions): P {
    return this.parent.table(arg as CellContent[][], opts) as P;
  }

  image(src: string | { data: Uint8Array | ArrayBuffer; mime: string }, opts: ImageOptions): P {
    return this.parent.image(src, opts) as P;
  }

  list(items: (string | ListItem)[], opts?: ListOptions): P {
    return this.parent.list(items, opts) as P;
  }

  bulletList(items: (string | ListItem)[]): P {
    return this.parent.bulletList(items) as P;
  }

  numberedList(items: (string | ListItem)[]): P {
    return this.parent.numberedList(items) as P;
  }

  pageBreak(): P {
    return this.parent.pageBreak() as P;
  }

  // ---- document-level delegators (only when the parent is the root builder) -

  header(this: ParagraphBuilder<DocumentBuilder>, build: (s: StoryBuilder) => void, opts?: BandOptions): DocumentBuilder {
    return this.parent.header(build, opts);
  }

  footer(this: ParagraphBuilder<DocumentBuilder>, build: (s: StoryBuilder) => void, opts?: BandOptions): DocumentBuilder {
    return this.parent.footer(build, opts);
  }

  pageSetup(this: ParagraphBuilder<DocumentBuilder>, setup: PageSetup): DocumentBuilder {
    return this.parent.pageSetup(setup);
  }

  style(this: ParagraphBuilder<DocumentBuilder>, def: NamedStyle): DocumentBuilder {
    return this.parent.style(def);
  }

  build(this: ParagraphBuilder<DocumentBuilder>): Document {
    return this.parent.build();
  }
}
