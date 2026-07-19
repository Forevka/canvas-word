// Public type surface for @forevka/wordcanvas/builder — the fluent
// programmatic document composer. Hand-written, self-contained (model types
// come from ./model, shared with wordcanvas.d.ts).
//
//   import { DocumentBuilder } from "@forevka/wordcanvas/builder";
//   const b = await DocumentBuilder.fromTemplate(templateBytes);
//   const doc = b.paragraph(data.title).withStyle("Heading1").build();
//
// Works in the browser (live preview via WordCanvas.setDocument) and in Node
// (server-side DOCX/PDF generation via the export subpath).

import type { CellBorders, CellMargin, CharStyle, Document, NamedStyle, PageNumFmt, ParaBorders, ParaStyle, SdtProps, Stylesheet, TableCondOverrides, TableStyle, TabStop } from "./model";

export type { Block, CharStyle, Document, NamedStyle, ParaStyle, Stylesheet, TableCondOverrides, TableStyle, TabStop } from "./model";

// ---------------------------------------------------------------------------
// Units

/** Inches → px @96dpi. */
export declare function inches(n: number): number;
/** Centimeters → px @96dpi. */
export declare function cm(n: number): number;
/** Points (1/72in) → px @96dpi. */
export declare function pt(n: number): number;
/** Twips (1/20pt) → px @96dpi. */
export declare function twips(n: number): number;

export interface PageSize {
  pageWidthPx: number;
  pageHeightPx: number;
}

/** Common page sizes in px @96dpi. */
export declare const PAGE_SIZES: Record<"Letter" | "Legal" | "A4" | "A3" | "Tabloid", PageSize>;
export type PageSizeName = keyof typeof PAGE_SIZES;

/** Base64-encode image bytes as a portable data: URL. */
export declare function bytesToDataUrl(bytes: Uint8Array | ArrayBuffer, mime: string): string;

// ---------------------------------------------------------------------------
// Options

export interface BuilderWarning {
  code: string;
  message: string;
}

export interface CreateOptions {
  pageSize?: PageSizeName | PageSize;
  margins?: Partial<{ top: number; right: number; bottom: number; left: number }>;
  /** Start from a custom stylesheet (default: the editor's default gallery). */
  stylesheet?: Stylesheet;
  /** Deterministic id namespace (tests); default is a random per-builder seed. */
  idSeed?: string;
}

export interface TemplateOptions {
  /** Keep the template's body content and append after it (default: discard —
   *  the template contributes styles, lists, page setup, and bands). */
  keepBody?: boolean;
  idSeed?: string;
}

export interface PageSetup {
  pageSize?: PageSizeName | PageSize;
  orientation?: "portrait" | "landscape";
  margins?: Partial<{ top: number; right: number; bottom: number; left: number }>;
  columns?: { count: number; gapPx?: number };
  headerDistancePx?: number;
  footerDistancePx?: number;
  pageNumberStart?: number;
}

export interface BandOptions {
  /** Which page band variant: default (all pages), first page, or even pages. */
  variant?: "default" | "first" | "even";
}

export interface ImageOptions {
  /** Required — the builder runs in Node too, where images can't be measured. */
  widthPx: number;
  heightPx: number;
  align?: "left" | "center" | "right";
  wrap?: "block" | "square";
}

export interface ListItem {
  text: string;
  /** Nesting level 0..8 (default 0, or ListOptions.level). */
  level?: number;
  style?: Partial<CharStyle>;
}

export interface ListOptions {
  kind?: "bullet" | "number";
  /** Reference an existing list definition (e.g. one carried by the template). */
  listId?: string;
  level?: number;
}

export interface CellSpec {
  text?: string;
  colSpan?: number;
  rowSpan?: number;
  /** Background fill (CSS color). */
  shading?: string;
  borders?: CellBorders;
  margin?: CellMargin;
  /** Char formatting for the cell's text. */
  style?: Partial<CharStyle>;
  align?: ParaStyle["align"];
  /** Preferred cell width (w:tcW): absolute px or percent of table width. */
  preferredWidth?: { px: number; type: "abs" | "pct" };
}

export type CellContent = string | CellSpec;
export type CellOptions = Omit<CellSpec, "text">;

export interface TableOptions {
  /** Column widths as fractions of the content width (normalized to sum 1). */
  colFractions?: number[];
  /** Bold every run in the first row. */
  headerRow?: boolean;
  /** Apply a registered builder-only table-style PRESET by name (baked into cells).
   *  Mutually exclusive with `styleId` — if both are given, `styleId` wins. */
  style?: string;
  /** Reference a REAL table style (DocumentBuilder.tableStyle) and bake its effective
   *  per-cell formatting onto the cells while keeping the styleId for docx round-trip. */
  styleId?: string;
  /** Which conditional bands of the referenced styleId are active (w:tblLook). */
  condOverrides?: TableCondOverrides;
  /** Column sizing strategy (w:tblLayout): "fixed" (default), "autofitContents", "autofitWindow". */
  widthMode?: "fixed" | "autofitContents" | "autofitWindow";
}

export interface EquationOptions {
  /** Horizontal placement of the display equation (default "center"). */
  align?: "left" | "center" | "right";
}

export interface SpacingOptions {
  before?: number;
  after?: number;
  lineHeight?: number;
  /** Fixed line-spacing rule (docx w:lineRule). Set `lineHeightPx` alongside it. */
  lineRule?: "exact" | "atLeast";
  /** Fixed line height in px — used with `lineRule`. */
  lineHeightPx?: number;
}

export interface IndentOptions {
  left?: number;
  right?: number;
  firstLine?: number;
}

// ---------------------------------------------------------------------------
// Scopes

/** Block-scope surface shared by the document body, header/footer bands, and
 *  table cells. Methods append eagerly and return the scope (or a paragraph
 *  scope that delegates back), so chains never need a seal step. */
export declare class StoryBuilder {
  paragraph(text?: string, style?: Partial<CharStyle>): ParagraphBuilder<this>;
  table(rows: CellContent[][], opts?: TableOptions): this;
  table(build: (t: TableBuilder) => void, opts?: TableOptions): this;
  image(src: string | { data: Uint8Array | ArrayBuffer; mime: string }, opts: ImageOptions): this;
  list(items: (string | ListItem)[], opts?: ListOptions): this;
  bulletList(items: (string | ListItem)[]): this;
  numberedList(items: (string | ListItem)[]): this;
  /** A display (block) equation from a LaTeX source (e.g. `\frac{a}{b}`). */
  equation(latex: string, opts?: EquationOptions): this;
  /** A display (block) equation from a presentation-MathML string. */
  equationMathml(mathml: string, opts?: EquationOptions): this;
  /** The NEXT block starts a new page. */
  pageBreak(): this;
}

/** Paragraph scope. Char formatting (bold/color/font/…) patches every run
 *  already in the paragraph and becomes the default for later text() runs;
 *  mixed formatting uses text(t, { …patch }). Any block-starting call pops
 *  back to the parent scope. */
export declare class ParagraphBuilder<P extends StoryBuilder> {
  /** Apply a named style: sets the reference and patches exactly the fields
   *  the style defines. Direct formatting applied AFTER this call wins. */
  withStyle(id: string): this;
  /** Append a run. Inherits the scope's accrued formatting; `style` wins. */
  text(text: string, style?: Partial<CharStyle>): this;
  bold(on?: boolean): this;
  italic(on?: boolean): this;
  underline(on?: boolean): this;
  strikethrough(on?: boolean): this;
  color(cssColor: string): this;
  highlight(cssColor: string): this;
  fontSize(px: number): this;
  font(family: string): this;
  link(url: string): this;
  /** Force this run's text to a right-to-left embedding (OOXML w:rtl). */
  rtl(on?: boolean): this;
  /** Apply a registered character style: bakes its formatting onto the runs AND
   *  sets the w:rStyle reference. Unknown ids are ignored with a warning. */
  charStyle(id: string): this;
  /** Append an INLINE equation from a LaTeX source (a single replaced glyph). */
  inlineEquation(latex: string, style?: Partial<CharStyle>): this;
  /** Append an inline equation from a presentation-MathML string. */
  inlineEquationMathml(mathml: string, style?: Partial<CharStyle>): this;
  // ---- inline content controls (OOXML w:sdt) ----
  /** Append an inline content control wrapping `text` with explicit properties. */
  contentControl(props: SdtProps, text: string, style?: Partial<CharStyle>): this;
  /** A rich-text content control. */
  richTextControl(text: string, opts?: { alias?: string; tag?: string }): this;
  /** A plain-text content control. */
  plainTextControl(text: string, opts?: { alias?: string; tag?: string }): this;
  /** A checkbox content control (☒/☐). */
  checkbox(checked: boolean, opts?: { alias?: string; tag?: string }): this;
  /** A drop-down content control (fixed list of `{ display, value }` items). */
  dropDown(selected: string, items: { display: string; value: string }[], opts?: { alias?: string; tag?: string }): this;
  /** A combo-box content control (list + free text). */
  comboBox(selected: string, items: { display: string; value: string }[], opts?: { alias?: string; tag?: string }): this;
  /** A date-picker content control. */
  dateControl(text: string, dateFormat?: string, opts?: { alias?: string; tag?: string }): this;
  // ---- inline fields (OOXML complex fields) ----
  /** The current page number (PAGE field), resolved at layout. */
  pageField(numFmt?: PageNumFmt): this;
  /** The total page count (NUMPAGES field), resolved at layout. */
  numPagesField(numFmt?: PageNumFmt): this;
  /** A DATE field with a Word/.NET-style format (default "M/d/yyyy"). */
  dateField(format?: string): this;
  /** A TIME field with a Word/.NET-style format (default "h:mm AM/PM"). */
  timeField(format?: string): this;
  /** A custom (host-resolved) field: verbatim instruction + cached result text. */
  customField(instruction: string, resultText: string): this;
  /** A conditional (IF) field — branch text by comparing two operands. */
  ifField(operandA: string, op: "=" | "<>" | "<" | "<=" | ">" | ">=", operandB: string, ifTrue: string, ifFalse: string): this;
  /** A cross-reference to a bookmark (REF, or PAGEREF when `pageRef`). */
  crossReference(bookmarkName: string, opts?: { kind?: "ref" | "pageRef" }): this;
  // ---- bookmarks + notes ----
  /** Append text and bookmark exactly that run's range. */
  bookmark(name: string, text: string, style?: Partial<CharStyle>): this;
  /** Append an auto-numbered footnote (single-paragraph body, or built via callback). */
  footnote(text: string): this;
  footnote(build: (s: StoryBuilder) => void): this;
  /** Append an auto-numbered endnote (collected at the document end). */
  endnote(text: string): this;
  endnote(build: (s: StoryBuilder) => void): this;
  /** Sub/superscript (measured smaller, baseline-shifted). */
  superscript(on?: boolean): this;
  subscript(on?: boolean): this;
  align(align: ParaStyle["align"]): this;
  spacing(opts: SpacingOptions): this;
  indent(opts: IndentOptions): this;
  keepWithNext(on?: boolean): this;
  /** Never split this paragraph across pages/columns (docx w:keepLines). */
  keepTogether(on?: boolean): this;
  /** Base writing direction (OOXML w:bidi); "rtl" lays the paragraph out right-to-left. */
  direction(dir: "ltr" | "rtl"): this;
  /** Outline level 0..8 (TOC levels 1..9) — a TOC entry without a heading style. */
  outlineLevel(level: number): this;
  /** Explicit tab stops (docx w:tabs); stored sorted by position. */
  tabStops(stops: TabStop[]): this;
  /** Paragraph borders (OOXML w:pBdr) — a box around the paragraph; each edge
   *  reuses the table CellBorder value type. */
  borders(borders: ParaBorders): this;
  /** Paragraph shading — a CSS fill painted behind the paragraph (OOXML w:shd). */
  shading(cssColor: string): this;
  /** Make THIS paragraph a list member (OOXML w:numPr) at `level` (0..8) — the
   *  single-paragraph counterpart to the story-level `.list()`, for items that
   *  carry formatting `.list()` cannot express (rich runs, a named style, spacing).
   *  `listId` should reference a definition registered via `.listDefinition()`. */
  listItem(listId: string, level?: number): this;
  /** Escape the paragraph scope explicitly (block-starting calls do it implicitly). */
  end(): P;

  // Delegated block-starting surface (pops this scope):
  paragraph(text?: string, style?: Partial<CharStyle>): ParagraphBuilder<P>;
  table(rows: CellContent[][], opts?: TableOptions): P;
  table(build: (t: TableBuilder) => void, opts?: TableOptions): P;
  image(src: string | { data: Uint8Array | ArrayBuffer; mime: string }, opts: ImageOptions): P;
  list(items: (string | ListItem)[], opts?: ListOptions): P;
  bulletList(items: (string | ListItem)[]): P;
  numberedList(items: (string | ListItem)[]): P;
  equation(latex: string, opts?: EquationOptions): P;
  equationMathml(mathml: string, opts?: EquationOptions): P;
  pageBreak(): P;

  // Document-level delegators (only when the parent is the root builder):
  header(this: ParagraphBuilder<DocumentBuilder>, build: (s: StoryBuilder) => void, opts?: BandOptions): DocumentBuilder;
  footer(this: ParagraphBuilder<DocumentBuilder>, build: (s: StoryBuilder) => void, opts?: BandOptions): DocumentBuilder;
  pageSetup(this: ParagraphBuilder<DocumentBuilder>, setup: PageSetup): DocumentBuilder;
  style(this: ParagraphBuilder<DocumentBuilder>, def: NamedStyle): DocumentBuilder;
  build(this: ParagraphBuilder<DocumentBuilder>): Document;
}

export declare class TableBuilder {
  row(cells: CellContent[]): this;
  row(build: (r: RowBuilder) => void): this;
  rows(data: CellContent[][]): this;
  colFractions(fractions: number[]): this;
}

export declare class RowBuilder {
  // The build (callback) overload is listed FIRST so a `(s) => …` arrow binds to it
  // and `s` types as StoryBuilder — CellSpec's fields are all optional, so a function
  // would otherwise match the `content: CellContent` overload and shadow the callback.
  cell(build: (s: StoryBuilder) => void, opts?: CellOptions): this;
  cell(content: CellContent, opts?: CellOptions): this;
}

/** Root of the fluent builder. Produces the same plain-data Document the
 *  editor renders (WordCanvas.setDocument) and the exporters write. build()
 *  returns a deep clone; the builder stays usable afterwards. */
export declare class DocumentBuilder extends StoryBuilder {
  /** Start from a blank document (default stylesheet, US Letter, 1in margins). */
  static create(opts?: CreateOptions): DocumentBuilder;
  /** Start from a .docx template: its stylesheet, list definitions, page setup,
   *  and header/footer bands carry over (body discarded unless keepBody). */
  static fromTemplate(docx: ArrayBuffer | Uint8Array, opts?: TemplateOptions): Promise<DocumentBuilder>;

  /** Lossy decisions (unknown style ids, template import warnings, …). */
  readonly warnings: readonly BuilderWarning[];
  /** Compose a header band (replaces the variant's existing story). */
  header(build: (s: StoryBuilder) => void, opts?: BandOptions): this;
  footer(build: (s: StoryBuilder) => void, opts?: BandOptions): this;
  /** Merge page geometry onto the section (template values are the base). */
  pageSetup(setup: PageSetup): this;
  /** Register (or override) a named style — BEFORE applying it via withStyle(). */
  style(def: NamedStyle): this;
  /** Register (or override) a REAL table style (OOXML w:style[type=table]) with
   *  conditional bands. Reference it via .table(rows, { styleId: def.id }). */
  tableStyle(def: TableStyle): this;
  /** Snapshot the document: deep clone, guaranteed at least one paragraph. */
  build(): Document;
}
