// Paragraph scope. Returned by StoryBuilder.paragraph(); styles the paragraph
// it wraps, and re-implements the block-starting surface by delegating to its
// parent scope — so `.paragraph(..).withStyle(..).paragraph(..)` pops back up
// without an explicit end-of-scope call.
//
// Character formatting semantics: bold()/color()/font()/… patch EVERY run
// currently in the paragraph AND become the default for runs added later by
// text() in this scope — "make this paragraph bold" is the dominant authoring
// intent. Mixed formatting within a paragraph uses text(t, { …patch }).

import type { Block, CellBorder, CharStyle, Document, EmphasisMark, FieldSpec, IfOp, NamedStyle, PageNumFmt, ParaBorders, ParaStyle, Paragraph, Run, SdtProps, TableStyle, TabStop, UnderlineStyle } from "@cw/shared";
import { AUTO_PARA_SPACING_PX, buildInstruction, evaluateField, resolveStyle, styleById, textOfRuns } from "@cw/shared";
import type { BuilderContext } from "./blockFactory";
import type { BandOptions, DocumentBuilder, ListDefinitionSpec, PageSetup, SectionBreakOptions } from "./documentBuilder";
import { equationFromLatex, equationFromMathml } from "./mathInput";
import type { EquationOptions, ImageOptions, ListItem, ListOptions } from "./storyBuilder";
import { StoryBuilder } from "./storyBuilder";
import type { CellContent, TableBuilder, TableOptions } from "./tableBuilder";
import type { TableStylePreset } from "./tableStyles";
import type { TocOptions } from "@cw/shared";

export interface SpacingOptions {
  /** Space before the paragraph, px. */
  before?: number;
  /** Space after the paragraph, px. */
  after?: number;
  /** Line height multiplier (1.0 = single). Used when `lineRule` is omitted. */
  lineHeight?: number;
  /** Fixed line-spacing rule (docx w:lineRule). "exact" = the line is exactly
   *  `lineHeightPx` tall (taller content clips); "atLeast" = at least that,
   *  growing for a taller line. Set `lineHeightPx` alongside it. */
  lineRule?: "exact" | "atLeast";
  /** Fixed line height in px — used with `lineRule`. */
  lineHeightPx?: number;
  /** Word's "automatic" space before/after (OOXML w:beforeAutospacing/@w:afterAutospacing,
   *  issue #160). When true, bakes the approximate auto value and re-emits the attribute;
   *  overrides `before`/`after` respectively. */
  beforeAuto?: boolean;
  afterAuto?: boolean;
}

export interface IndentOptions {
  /** Left indent, px. */
  left?: number;
  /** Right indent, px. */
  right?: number;
  /** First-line indent, px (negative = hanging). */
  firstLine?: number;
}

/** Strict structural style equality for author-time run coalescing: every field
 *  of both styles must match (Object.is; object-typed fields like equation/
 *  sdtPath/runBorder compare by reference — builder-authored neighbors share
 *  them via the charPatch spread, and a content-equal-but-distinct object just
 *  skips the merge, which is harmless). Deliberately STRICTER than the editor's
 *  styleEq, which skips fields (charStyleId, rtl, endnoteRef, theme metadata)
 *  that must still keep author-time runs apart. */
function identicalStyles(a: CharStyle, b: CharStyle): boolean {
  const ka = Object.keys(a);
  if (ka.length !== Object.keys(b).length) return false;
  const ra = a as unknown as Record<string, unknown>;
  const rb = b as unknown as Record<string, unknown>;
  for (const k of ka) {
    if (!Object.is(ra[k], rb[k])) return false;
  }
  return true;
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
    // shading and the explicit-clear marker are mutually exclusive (issue #147):
    // whichever the applied style provides wins, so drop the stale counterpart a
    // prior .shading()/.clearShading() may have left — else export emits both.
    if (resolved.para.shading !== undefined) delete this.para.style.shadingCleared;
    else if (resolved.para.shadingCleared) delete this.para.style.shading;
    if (resolved.para.borders !== undefined) delete this.para.style.bordersCleared;
    else if (resolved.para.bordersCleared) delete this.para.style.borders;
    const charKeys = Object.keys(resolved.char);
    for (const r of this.para.runs) {
      Object.assign(r.style, resolved.char);
      // Author-explicit like applyChar: record the keys so table-style baking
      // preserves a styled value that equals the resolved default against a
      // conflicting band (#45) — .withStyle() previously skipped this.
      this.markRunCharKeys(r, charKeys);
    }
    Object.assign(this.charPatch, resolved.char);
    return this;
  }

  /** Append a run, swapping out the empty placeholder a textless paragraph holds.
   *  Shared by text() and every inline-content emitter (field/sdt/footnote/…).
   *  Coalesces into the previous run when the two styles are structurally
   *  identical (see identicalStyles — field/SDT/footnote/equation boundaries can
   *  never merge, since those fields differ): the model's invariant is that
   *  adjacent equal-styled runs are merged on every edit, and builder output
   *  should be canonical from the start instead of shipping fragmented runs
   *  from consecutive same-style text(). */
  private pushRun(run: Run): this {
    const runs = this.para.runs;
    if (runs.length === 1 && runs[0]!.text === "") {
      runs[0] = run;
      return this;
    }
    const last = runs[runs.length - 1];
    if (last && last.text.length > 0 && run.text.length > 0 && identicalStyles(last.style, run.style)) {
      last.text += run.text;
      // Union the merged run's explicit-key provenance into the survivor.
      const extra = this.ctx.explicitCharKeys.get(run);
      if (extra && extra.size > 0) this.markRunCharKeys(last, extra);
      return this;
    }
    runs.push(run);
    return this;
  }

  /** Append a run. Inherits the scope's accrued char formatting; `style` wins. */
  text(text: string, style?: Partial<CharStyle>): this {
    return this.pushRun(this.ctx.run(text, { ...this.charPatch, ...style }));
  }

  private applyChar(patch: Partial<CharStyle>): this {
    const keys = Object.keys(patch);
    if (keys.length === 0) return this; // e.g. effects({}) — nothing to do
    Object.assign(this.charPatch, patch);
    for (const r of this.para.runs) {
      Object.assign(r.style, patch);
      // Direct formatting applied to existing runs is author-explicit too — record
      // it so table-style baking preserves a value that equals the default against a
      // conflicting band (#45). Runs added LATER inherit charPatch via ctx.run, which
      // records their keys at creation, so this only needs the runs already present.
      this.markRunCharKeys(r, keys);
    }
    return this;
  }

  /** Merge `keys` into a run's explicit-char provenance (issue #45 — mirrors the
   *  data-driven cell path, where author-set CharStyle keys are tracked so an
   *  explicit value equal to the resolved default survives a conflicting band).
   *  Mutates the run's existing Set in place — each run owns its Set (minted in
   *  ctx.run / here), so chained formatting doesn't reallocate one per call. */
  private markRunCharKeys(run: Run, keys: Iterable<string>): void {
    const prev = this.ctx.explicitCharKeys.get(run);
    if (prev) {
      for (const k of keys) prev.add(k);
      return;
    }
    this.ctx.explicitCharKeys.set(run, new Set(keys));
  }

  bold(on = true): this {
    return this.applyChar({ bold: on });
  }

  italic(on = true): this {
    return this.applyChar({ italic: on });
  }

  /** Underline the paragraph's runs. `opts.style` selects the line style
   *  (double/dotted/dash/dotDash/dotDotDash/wave/thick; default a plain single
   *  line); `opts.color` paints a colored rule (CSS hex), else it follows the text. */
  underline(on = true, opts?: { style?: UnderlineStyle; color?: string }): this {
    const patch: Partial<CharStyle> = { underline: on };
    if (on && opts?.style) patch.underlineStyle = opts.style;
    if (on && opts?.color) patch.underlineColor = opts.color;
    return this.applyChar(patch);
  }

  strikethrough(on = true): this {
    return this.applyChar({ strikethrough: on });
  }

  /** Minor run typography & effects (OOXML w:rPr extras): double strikethrough
   *  (w:dstrike), baseline raise/lower in px (w:position; +up/−down), character
   *  width scaling as a percentage (w:w; 100 = normal), a kerning threshold in px
   *  (w:kern), emphasis marks (w:em), the outline/shadow/emboss/imprint text
   *  effects, a run border (w:bdr), and a fitText target width in px (w:fitText).
   *  Additive — only the provided fields are applied. */
  effects(opts: {
    doubleStrikethrough?: boolean;
    positionPx?: number;
    widthScalePct?: number;
    kerningMinPx?: number;
    emphasisMark?: EmphasisMark;
    outline?: boolean;
    shadow?: boolean;
    emboss?: boolean;
    imprint?: boolean;
    border?: CellBorder;
    fitTextPx?: number;
  }): this {
    const patch: Partial<CharStyle> = {};
    if (opts.doubleStrikethrough !== undefined) patch.doubleStrikethrough = opts.doubleStrikethrough;
    if (opts.positionPx !== undefined) patch.positionPx = opts.positionPx;
    if (opts.widthScalePct !== undefined) {
      // OOXML w:w (and the import/export round-trip) only preserve a positive
      // percentage in 1..600; reject out-of-range values rather than author state
      // the .docx pipeline would silently drop. Warn-and-skip (not throw) — the
      // builder's uniform invalid-input contract: a fluent chain keeps going and
      // the problem lands in .warnings, like withStyle()/list()/spacing().
      if (opts.widthScalePct <= 0 || opts.widthScalePct > 600) {
        this.ctx.warn("effects-width-scale-invalid", ".effects({ widthScalePct }) must be in 1..600; ignored.");
      } else {
        patch.widthScalePct = opts.widthScalePct;
      }
    }
    if (opts.kerningMinPx !== undefined) patch.kerningMinPx = opts.kerningMinPx;
    if (opts.emphasisMark !== undefined) patch.emphasisMark = opts.emphasisMark;
    if (opts.outline !== undefined) patch.outline = opts.outline;
    if (opts.shadow !== undefined) patch.shadow = opts.shadow;
    if (opts.emboss !== undefined) patch.emboss = opts.emboss;
    if (opts.imprint !== undefined) patch.imprint = opts.imprint;
    if (opts.border !== undefined) patch.runBorder = opts.border;
    if (opts.fitTextPx !== undefined) {
      // w:fitText is a positive width; a non-positive value would not round-trip.
      if (opts.fitTextPx <= 0) this.ctx.warn("effects-fit-text-invalid", ".effects({ fitTextPx }) must be greater than 0; ignored.");
      else patch.fitTextPx = opts.fitTextPx;
    }
    return this.applyChar(patch);
  }

  color(cssColor: string): this {
    return this.applyChar({ color: cssColor });
  }

  highlight(cssColor: string): this {
    return this.applyChar({ highlightColor: cssColor, highlightCleared: undefined });
  }

  /** Explicitly clear the run highlight (OOXML `<w:highlight w:val="none"/>`) —
   *  overrides a highlight the run's character style carries, so the clear survives
   *  export/re-import instead of the style's highlight silently returning (issue #155). */
  clearHighlight(): this {
    return this.applyChar({ highlightColor: undefined, highlightCleared: true });
  }

  fontSize(px: number): this {
    return this.applyChar({ fontSizePx: px });
  }

  font(family: string): this {
    return this.applyChar({ fontFamily: family });
  }

  /** Set the complex-script (bidi) font — OOXML w:rFonts/@w:cs. Preserved through
   *  the .docx round-trip; layout/paint still use the main font(). */
  fontComplexScript(family: string): this {
    return this.applyChar({ fontFamilyComplexScript: family });
  }

  /** Set the East-Asian (CJK) font — OOXML w:rFonts/@w:eastAsia. Preserved through
   *  the .docx round-trip; layout/paint still use the main font(). */
  fontEastAsia(family: string): this {
    return this.applyChar({ fontFamilyEastAsia: family });
  }

  /** Make the paragraph's text a hyperlink (painted blue+underlined, Ctrl+click). */
  link(url: string): this {
    return this.applyChar({ link: url });
  }

  /** Remove a char property from the scope default AND every existing run (the
   *  exactOptional-safe way to toggle an optional CharStyle field off). */
  private clearChar(key: keyof CharStyle): this {
    delete (this.charPatch as unknown as Record<string, unknown>)[key];
    for (const r of this.para.runs) {
      delete (r.style as unknown as Record<string, unknown>)[key];
      // No longer author-set — drop it from provenance so it falls back to value-equality.
      this.ctx.explicitCharKeys.get(r)?.delete(key);
    }
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

  /** All-caps display (OOXML w:caps): every letter renders UPPERCASED while the
   *  model text stays as authored (offset-transparent). Common on headings. An
   *  explicit `caps(false)` writes the OFF override (w:caps="0"), so it can clear
   *  casing inherited from a character style — mirroring bold(false). */
  caps(on = true): this {
    return this.applyChar({ caps: on });
  }

  /** Small-capitals display (OOXML w:smallCaps): letters render uppercased, with the
   *  originally-lowercase ones drawn smaller. Takes precedence over caps(). An
   *  explicit `smallCaps(false)` writes the OFF override (w:smallCaps="0"). */
  smallCaps(on = true): this {
    return this.applyChar({ smallCaps: on });
  }

  /** Force this run's text to a right-to-left embedding (OOXML w:rtl), regardless
   *  of its characters. For the paragraph's base direction use .direction("rtl"). */
  rtl(on = true): this {
    return on ? this.applyChar({ rtl: true }) : this.clearChar("rtl");
  }

  /** Append a symbol-font glyph (OOXML w:sym). `font` is the symbol font (e.g.
   *  "Wingdings"); `charHex` is the hex code point Word stores (e.g. "F0E0", usually
   *  a Private-Use value). The run renders the decoded glyph in `font` and re-emits
   *  w:sym on export. */
  symbol(font: string, charHex: string): this {
    const cp = parseInt(charHex, 16);
    const text = Number.isFinite(cp) && cp > 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : "�";
    return this.pushRun(this.ctx.run(text, { ...this.charPatch, fontFamily: `${font}, sans-serif`, symbol: { font, char: charHex.toUpperCase() } }));
  }

  /** Apply a registered character style (a type:"character" NamedStyle): bakes its
   *  formatting onto the runs AND sets the w:rStyle reference (kept for round-trip).
   *  Ids that are unknown OR refer to a paragraph style are ignored with a warning
   *  (w:rStyle must reference a character style), like .withStyle() for unknowns. */
  charStyle(id: string): this {
    const sheet = this.ctx.doc.stylesheet;
    const def = sheet ? styleById(sheet, id) : undefined;
    // NamedStyle.type defaults to "paragraph" when absent — only an explicit
    // "character" style is a valid w:rStyle target.
    if (!sheet || !def || def.type !== "character") {
      this.ctx.warn(
        `char-style-invalid:${id}`,
        `.charStyle("${id}") — no such character style (it is missing or a paragraph style); ignored.`,
      );
      return this;
    }
    // The style is validated present above — resolve directly (lookupStyle would
    // redundantly re-run the styleById existence check).
    return this.applyChar({ ...resolveStyle(sheet, id).char, charStyleId: id });
  }

  /** Append an INLINE equation from a LaTeX source — a single replaced glyph in the
   *  run flow (e.g. "the identity ", inlineEquation("e^{i\\pi}+1=0"), "."). */
  inlineEquation(latex: string, style?: Partial<CharStyle>): this {
    return this.pushRun(this.ctx.run("￼", { ...this.charPatch, ...style, equation: equationFromLatex(latex, false) }));
  }

  /** Append an inline equation from a presentation-MathML string. */
  inlineEquationMathml(mathml: string, style?: Partial<CharStyle>): this {
    return this.pushRun(this.ctx.run("￼", { ...this.charPatch, ...style, equation: equationFromMathml(mathml, false) }));
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

  pageField(numFmt?: PageNumFmt, opts?: { style?: Partial<CharStyle> }): this {
    return this.emitField(numFmt ? { type: "PAGE", numFmt } : { type: "PAGE" }, opts);
  }

  numPagesField(numFmt?: PageNumFmt, opts?: { style?: Partial<CharStyle> }): this {
    return this.emitField(numFmt ? { type: "NUMPAGES", numFmt } : { type: "NUMPAGES" }, opts);
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

  /** Shared body of footnote()/endnote() — the two differ only in which registry
   *  and counter they use and which ref key marks the reference run. */
  private emitNote(kind: "footnote" | "endnote", content: string | ((s: StoryBuilder) => void)): this {
    const notes = (this.ctx.doc[kind === "footnote" ? "footnotes" : "endnotes"] ??= {});
    const n = kind === "footnote" ? this.ctx.nextFootnoteNumber() : this.ctx.nextEndnoteNumber();
    const id = this.ctx.ids.next();
    let paras: Paragraph[];
    if (typeof content === "string") {
      paras = [this.ctx.paragraph([this.ctx.run(content, { fontSizePx: 12 })], { spaceAfterPx: 0 })];
    } else {
      const blocks: Block[] = [];
      content(new StoryBuilder(this.ctx, blocks));
      paras = blocks.filter((b): b is Paragraph => b.kind === "paragraph");
      if (paras.length < blocks.length) {
        this.ctx.warn(`${kind}-non-paragraph`, `${kind === "footnote" ? "Footnote" : "Endnote"} bodies hold paragraphs only — non-paragraph blocks were dropped.`);
      }
      if (paras.length === 0) paras = [this.ctx.paragraph([])];
    }
    notes[id] = paras;
    const refKey = kind === "footnote" ? { footnoteRef: id } : { endnoteRef: id };
    return this.pushRun(this.ctx.run(String(n), { ...refKey, verticalAlign: "super", fontSizePx: 11 }));
  }

  /** Append an auto-numbered footnote reference; the note body is a string (one
   *  paragraph) or a StoryBuilder callback (rich, multi-paragraph). */
  footnote(content: string | ((s: StoryBuilder) => void)): this {
    return this.emitNote("footnote", content);
  }

  /** Append an auto-numbered endnote reference; the note body is a string (one
   *  paragraph) or a StoryBuilder callback (rich, multi-paragraph). Endnotes
   *  collect at the END of the document instead of the page bottom. */
  endnote(content: string | ((s: StoryBuilder) => void)): this {
    return this.emitNote("endnote", content);
  }

  /** Append `text` and bookmark exactly that run's range in this paragraph. */
  bookmark(name: string, text: string, style?: Partial<CharStyle>): this {
    const bookmarks = (this.ctx.doc.bookmarks ??= {});
    // The appended run adds exactly text.length — no second full-run-list scan
    // (pushRun may coalesce into the previous run, but lengths still add up).
    const start = textOfRuns(this.para.runs).length;
    this.pushRun(this.ctx.run(text, { ...this.charPatch, ...style }));
    bookmarks[name] = { start: { blockId: this.para.id, offset: start }, end: { blockId: this.para.id, offset: start + text.length } };
    return this;
  }

  align(align: ParaStyle["align"]): this {
    this.para.style.align = align;
    // Author explicitly set align — record it (mirrors CellSpec.align in the data
    // path) so table-style baking preserves it even when it equals the default (#45).
    const prev = this.ctx.explicitParaKeys.get(this.para);
    this.ctx.explicitParaKeys.set(this.para, new Set(prev).add("align"));
    return this;
  }

  spacing(opts: SpacingOptions): this {
    if (opts.before !== undefined) this.para.style.spaceBeforePx = opts.before;
    if (opts.after !== undefined) this.para.style.spaceAfterPx = opts.after;
    // Auto-spacing (issue #160): bake the approximate auto value + set the export flag.
    if (opts.beforeAuto !== undefined) {
      this.para.style.spaceBeforeAuto = opts.beforeAuto;
      if (opts.beforeAuto) this.para.style.spaceBeforePx = AUTO_PARA_SPACING_PX;
    }
    if (opts.afterAuto !== undefined) {
      this.para.style.spaceAfterAuto = opts.afterAuto;
      if (opts.afterAuto) this.para.style.spaceAfterPx = AUTO_PARA_SPACING_PX;
    }
    if (opts.lineRule !== undefined) {
      // A fixed rule is meaningless without its height; never store a half-set rule
      // (it would export as w:line="0" and disagree with layout).
      if (opts.lineHeightPx === undefined) {
        this.ctx.warn("spacing-line-height-missing", ".spacing({ lineRule }) requires lineHeightPx; ignored.");
      } else {
        this.para.style.lineRule = opts.lineRule;
        this.para.style.lineHeightPx = opts.lineHeightPx;
      }
    } else if (opts.lineHeight !== undefined) {
      // Switching back to multiplier spacing — drop any previously-set fixed rule so
      // call-order precedence holds (a later .spacing({ lineHeight }) wins).
      delete (this.para.style as Partial<ParaStyle>).lineRule;
      delete (this.para.style as Partial<ParaStyle>).lineHeightPx;
    }
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

  /** Never split this paragraph across pages/columns (docx w:keepLines). */
  keepTogether(on = true): this {
    this.para.style.keepLinesTogether = on;
    return this;
  }

  /** Suppress before/after spacing between this paragraph and an adjacent
   *  same-style paragraph (docx w:contextualSpacing) — Word's list-style default. */
  contextualSpacing(on = true): this {
    this.para.style.contextualSpacing = on;
    return this;
  }

  /** Base writing direction (OOXML w:bidi). "rtl" lays the paragraph out
   *  right-to-left and mirrors start/end alignment + indents. */
  direction(dir: "ltr" | "rtl"): this {
    this.para.style.direction = dir;
    return this;
  }

  /** Outline level 0..8 (TOC levels 1..9; docx w:outlineLvl) — makes a paragraph a
   *  TOC entry without a heading style. Clamped to range. */
  outlineLevel(level: number): this {
    this.para.style.outlineLevel = Math.max(0, Math.min(8, Math.floor(level)));
    return this;
  }

  /** Explicit tab stops (docx w:tabs); a `\t` in run text advances to the next.
   *  Stored sorted by position, matching the layout engine's expectation. */
  tabStops(stops: TabStop[]): this {
    this.para.style.tabStops = [...stops].sort((a, b) => a.posPx - b.posPx);
    return this;
  }

  /** Paragraph borders (OOXML w:pBdr) — a box around the paragraph. Each edge
   *  reuses the table border value type (color + widthPx + optional line style);
   *  omit an edge to leave that side open. `between` round-trips but is not drawn
   *  for a standalone paragraph. */
  borders(borders: ParaBorders): this {
    this.para.style.borders = borders;
    delete this.para.style.bordersCleared; // a concrete box supersedes an explicit clear
    return this;
  }

  /** Explicitly clear the paragraph border box (OOXML empty `<w:pBdr/>`) — overrides a
   *  box the paragraph's named style carries, so the clear survives export/re-import
   *  instead of the style's box silently returning (issue #153). */
  clearBorders(): this {
    delete this.para.style.borders;
    this.para.style.bordersCleared = true;
    return this;
  }

  /** Paragraph shading — a CSS fill painted behind the paragraph (OOXML w:shd). */
  shading(cssColor: string): this {
    this.para.style.shading = cssColor;
    delete this.para.style.shadingCleared; // a concrete fill supersedes an explicit clear
    return this;
  }

  /** Explicitly clear paragraph shading (OOXML `<w:shd w:val="clear" w:fill="auto"/>`,
   *  Word's "Shading → No Color") — overrides a fill the paragraph's named style
   *  carries, so the clear survives export/re-import instead of the style's fill
   *  silently returning (issue #147). Distinct from simply not setting shading. */
  clearShading(): this {
    delete this.para.style.shading;
    this.para.style.shadingCleared = true;
    return this;
  }

  /** Explicitly remove list membership (OOXML `<w:numPr><w:numId w:val="0"/>`,
   *  Word's "opt out of the list") — overrides a list the paragraph's named style
   *  carries, so the opt-out survives export/re-import instead of the style's list
   *  silently returning (issue #152). Distinct from simply not being in a list. */
  clearList(): this {
    delete this.para.style.list;
    this.para.style.listCleared = true;
    return this;
  }

  /** Widow/orphan control (OOXML w:widowControl). Word's default is ON; pass
   *  `false` to let a lone first/last line break across a page boundary. */
  widowControl(on = true): this {
    this.para.style.widowControl = on;
    return this;
  }

  /** Exclude this paragraph from line numbering (OOXML w:suppressLineNumbers). */
  suppressLineNumbers(on = true): this {
    this.para.style.suppressLineNumbers = on;
    return this;
  }

  /** Vertical alignment of text within each line box (OOXML w:textAlignment):
   *  "top"/"center"/"bottom" hug the edge, "baseline" (default) rides the baseline. */
  textAlignment(v: "top" | "center" | "bottom" | "baseline"): this {
    this.para.style.textAlignment = v;
    return this;
  }

  /** Symmetric (mirrored) indents for facing-page layouts (OOXML w:mirrorIndents). */
  mirrorIndents(on = true): this {
    this.para.style.mirrorIndents = on;
    return this;
  }

  /** Auto-adjust the right indent to the document grid (OOXML w:adjustRightInd). */
  adjustRightInd(on = true): this {
    this.para.style.adjustRightInd = on;
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

  equation(latex: string, opts?: EquationOptions): P {
    return this.parent.equation(latex, opts) as P;
  }

  equationMathml(mathml: string, opts?: EquationOptions): P {
    return this.parent.equationMathml(mathml, opts) as P;
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

  contentControlRange(props: SdtProps, build: (s: StoryBuilder) => void): P {
    return this.parent.contentControlRange(props, build as (s: P) => void) as P;
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

  tableStyle(this: ParagraphBuilder<DocumentBuilder>, def: TableStyle): DocumentBuilder {
    return this.parent.tableStyle(def);
  }
}
