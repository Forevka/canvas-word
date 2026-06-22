// Paragraph scope. Returned by StoryBuilder.paragraph(); styles the paragraph
// it wraps, and re-implements the block-starting surface by delegating to its
// parent scope — so `.paragraph(..).withStyle(..).paragraph(..)` pops back up
// without an explicit end-of-scope call.
//
// Character formatting semantics: bold()/color()/font()/… patch EVERY run
// currently in the paragraph AND become the default for runs added later by
// text() in this scope — "make this paragraph bold" is the dominant authoring
// intent. Mixed formatting within a paragraph uses text(t, { …patch }).

import type { Block, CharStyle, Document, FieldSpec, IfOp, NamedStyle, PageNumFmt, ParaStyle, Paragraph, Run, SdtProps } from "@cw/shared";
import { buildInstruction, evaluateField, textOfRuns } from "@cw/shared";
import type { BuilderContext } from "./blockFactory";
import type { BandOptions, DocumentBuilder, ListDefinitionSpec, PageSetup, SectionBreakOptions } from "./documentBuilder";
import type { ImageOptions, ListItem, ListOptions } from "./storyBuilder";
import { StoryBuilder } from "./storyBuilder";
import type { CellContent, TableBuilder, TableOptions } from "./tableBuilder";
import type { TableStylePreset } from "./tableStyles";
import type { TocOptions } from "@cw/shared";

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

  /** Append a run, swapping out the empty placeholder a textless paragraph holds.
   *  Shared by text() and every inline-content emitter (field/sdt/footnote/…). */
  private pushRun(run: Run): this {
    if (this.para.runs.length === 1 && this.para.runs[0]!.text === "") this.para.runs[0] = run;
    else this.para.runs.push(run);
    return this;
  }

  /** Append a run. Inherits the scope's accrued char formatting; `style` wins. */
  text(text: string, style?: Partial<CharStyle>): this {
    return this.pushRun(this.ctx.run(text, { ...this.charPatch, ...style }));
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

  /** Remove a char property from the scope default AND every existing run (the
   *  exactOptional-safe way to toggle an optional CharStyle field off). */
  private clearChar(key: keyof CharStyle): this {
    delete (this.charPatch as unknown as Record<string, unknown>)[key];
    for (const r of this.para.runs) delete (r.style as unknown as Record<string, unknown>)[key];
    return this;
  }

  superscript(on = true): this {
    return on ? this.applyChar({ verticalAlign: "super" }) : this.clearChar("verticalAlign");
  }

  subscript(on = true): this {
    return on ? this.applyChar({ verticalAlign: "sub" }) : this.clearChar("verticalAlign");
  }

  letterSpacing(px: number): this {
    return this.applyChar({ letterSpacingPx: px });
  }

  /** Hidden text — kept in the model + .docx, never laid out/painted. */
  hidden(on = true): this {
    return on ? this.applyChar({ hidden: true }) : this.clearChar("hidden");
  }

  // ---- inline fields -------------------------------------------------------

  /** Register a built-in field def and append its (fieldId-tagged) result run.
   *  PAGE/NUMPAGES emit a live `{page}`/`{pages}` token; DATE/TIME/IF materialize
   *  via the shared evaluator so the builder agrees with the editor. */
  private emitField(spec: FieldSpec, opts?: { now?: Date; style?: Partial<CharStyle> }): this {
    const fields = (this.ctx.doc.fields ??= {});
    const fid = this.ctx.ids.next();
    fields[fid] = { id: fid, instruction: buildInstruction(spec), name: spec.type, kind: "builtin", spec };
    const out = evaluateField(spec, this.ctx.charDefault, { now: opts?.now ?? new Date() });
    const resultText = out.kind === "token" ? out.token : out.runs.map((r) => r.text).join("");
    return this.pushRun(this.ctx.run(resultText, { ...this.charPatch, ...opts?.style, fieldId: fid }));
  }

  /** Insert any built-in field from its typed spec. */
  field(spec: FieldSpec, opts?: { now?: Date; style?: Partial<CharStyle> }): this {
    return this.emitField(spec, opts);
  }

  pageField(numFmt?: PageNumFmt): this {
    return this.emitField(numFmt ? { type: "PAGE", numFmt } : { type: "PAGE" });
  }

  numPagesField(numFmt?: PageNumFmt): this {
    return this.emitField(numFmt ? { type: "NUMPAGES", numFmt } : { type: "NUMPAGES" });
  }

  dateField(format = "M/d/yyyy", opts?: { now?: Date }): this {
    return this.emitField({ type: "DATE", format }, opts);
  }

  timeField(format = "h:mm AM/PM", opts?: { now?: Date }): this {
    return this.emitField({ type: "TIME", format }, opts);
  }

  /** Conditional field. Branch results are plain text (matching the editor's
   *  field constructor — rich branch content stays raw-model only). */
  ifField(operandA: string, op: IfOp, operandB: string, ifTrue: string, ifFalse: string): this {
    return this.emitField({
      type: "IF", operandA, op, operandB,
      trueRuns: [this.ctx.run(ifTrue)], falseRuns: [this.ctx.run(ifFalse)],
    });
  }

  /** A non-built-in (host-resolved) field: verbatim instruction + cached result.
   *  The escape hatch for PAGEREF/REF/SEQ/STYLEREF/etc. (see crossReference). */
  customField(instruction: string, resultText: string, opts?: { name?: string; style?: Partial<CharStyle> }): this {
    const fields = (this.ctx.doc.fields ??= {});
    const fid = this.ctx.ids.next();
    const name = opts?.name ?? instruction.trim().split(/\s+/)[0]?.toUpperCase() ?? "FIELD";
    fields[fid] = { id: fid, instruction, name, kind: "custom" };
    return this.pushRun(this.ctx.run(resultText, { ...this.charPatch, ...opts?.style, fieldId: fid }));
  }

  /** A cross-reference to a bookmark, as a REF/PAGEREF field over customField. */
  crossReference(bookmarkName: string, opts?: { kind?: "ref" | "pageRef"; resultText?: string }): this {
    const kind = opts?.kind ?? "ref";
    const instruction = kind === "pageRef" ? ` PAGEREF ${bookmarkName} \\h ` : ` REF ${bookmarkName} \\h `;
    const resultText = opts?.resultText ?? (kind === "pageRef" ? "1" : bookmarkName);
    return this.customField(instruction, resultText, { name: kind === "pageRef" ? "PAGEREF" : "REF" });
  }

  // ---- inline content controls (SDT) ---------------------------------------

  private emitSdt(props: SdtProps, text: string, style?: Partial<CharStyle>): this {
    const sdts = (this.ctx.doc.sdts ??= {});
    const sid = this.ctx.ids.next();
    sdts[sid] = props;
    return this.pushRun(this.ctx.run(text, { ...this.charPatch, ...style, sdtPath: [sid] }));
  }

  /** Insert any content control from its props. */
  contentControl(props: SdtProps, text: string, style?: Partial<CharStyle>): this {
    return this.emitSdt(props, text, style);
  }

  richTextControl(text: string, opts?: { alias?: string; tag?: string }): this {
    return this.emitSdt({ type: "richText", ...opts }, text);
  }

  plainTextControl(text: string, opts?: { alias?: string; tag?: string }): this {
    return this.emitSdt({ type: "plainText", ...opts }, text);
  }

  checkbox(checked: boolean, opts?: { alias?: string; tag?: string }): this {
    return this.emitSdt({ type: "checkbox", checked, ...opts }, checked ? "☒" : "☐");
  }

  dropDown(selected: string, items: { display: string; value: string }[], opts?: { alias?: string; tag?: string }): this {
    return this.emitSdt({ type: "dropDown", listItems: items, ...opts }, selected);
  }

  comboBox(selected: string, items: { display: string; value: string }[], opts?: { alias?: string; tag?: string }): this {
    return this.emitSdt({ type: "comboBox", listItems: items, ...opts }, selected);
  }

  dateControl(text: string, dateFormat = "M/d/yyyy", opts?: { alias?: string; tag?: string }): this {
    return this.emitSdt({ type: "date", dateFormat, ...opts }, text);
  }

  // ---- footnotes + bookmarks -----------------------------------------------

  /** Append an auto-numbered footnote reference; the note body is a string (one
   *  paragraph) or a StoryBuilder callback (rich, multi-paragraph). */
  footnote(content: string | ((s: StoryBuilder) => void)): this {
    const footnotes = (this.ctx.doc.footnotes ??= {});
    const n = this.ctx.nextFootnoteNumber();
    const id = this.ctx.ids.next();
    let paras: Paragraph[];
    if (typeof content === "string") {
      paras = [this.ctx.paragraph([this.ctx.run(content, { fontSizePx: 12 })], { spaceAfterPx: 0 })];
    } else {
      const blocks: Block[] = [];
      content(new StoryBuilder(this.ctx, blocks));
      paras = blocks.filter((b): b is Paragraph => b.kind === "paragraph");
      if (paras.length < blocks.length) this.ctx.warn("footnote-non-paragraph", "Footnote bodies hold paragraphs only — non-paragraph blocks were dropped.");
      if (paras.length === 0) paras = [this.ctx.paragraph([])];
    }
    footnotes[id] = paras;
    return this.pushRun(this.ctx.run(String(n), { footnoteRef: id, verticalAlign: "super", fontSizePx: 11 }));
  }

  /** Append `text` and bookmark exactly that run's range in this paragraph. */
  bookmark(name: string, text: string, style?: Partial<CharStyle>): this {
    const bookmarks = (this.ctx.doc.bookmarks ??= {});
    const start = textOfRuns(this.para.runs).length;
    this.pushRun(this.ctx.run(text, { ...this.charPatch, ...style }));
    const end = textOfRuns(this.para.runs).length;
    bookmarks[name] = { start: { blockId: this.para.id, offset: start }, end: { blockId: this.para.id, offset: end } };
    return this;
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

  columnBreak(): P {
    return this.parent.columnBreak() as P;
  }

  bookmarkRange(name: string, build: (s: StoryBuilder) => void): P {
    return this.parent.bookmarkRange(name, build) as P;
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

  tableOfContents(this: ParagraphBuilder<DocumentBuilder>, opts?: TocOptions): DocumentBuilder {
    return this.parent.tableOfContents(opts);
  }

  sectionBreak(this: ParagraphBuilder<DocumentBuilder>, opts?: SectionBreakOptions): DocumentBuilder {
    return this.parent.sectionBreak(opts);
  }

  defaultStyle(this: ParagraphBuilder<DocumentBuilder>, id: string): DocumentBuilder {
    return this.parent.defaultStyle(id);
  }

  listDefinition(this: ParagraphBuilder<DocumentBuilder>, id: string, spec: ListDefinitionSpec): DocumentBuilder {
    return this.parent.listDefinition(id, spec);
  }

  tableStylePreset(this: ParagraphBuilder<DocumentBuilder>, name: string, preset: TableStylePreset): DocumentBuilder {
    return this.parent.tableStylePreset(name, preset);
  }
}
