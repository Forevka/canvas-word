// Reverse builder: a `Document` model → TypeScript source that calls the fluent
// `DocumentBuilder` API to reconstruct an equivalent document. The inverse of the
// builder — take a `.docx` (imported to a Document) and get editable template
// code back.
//
// Fidelity strategy. Runs reproduce almost exactly through the generic
// `.text(text, patch)` escape hatch: we pass the full STRUCTURAL DELTA of a run's
// CharStyle against the resolved baseline the builder itself would produce (the
// same `resolveStyle` cascade `BuilderContext` uses), so every character field —
// including the self-contained `equation`/`symbol`/`charStyleId` — comes back
// verbatim. Paragraph properties map to their fluent methods the same way.
// Stylesheet, list definitions and table styles reproduce verbatim via
// `create({ stylesheet })` / `.listDefinition` / `.tableStyle`.
//
// What the fluent API cannot yet express is NOT silently dropped: each such field
// is recorded in `uncovered` (node path + field + sample value). That report is
// the backlog for extending the builder/model — the whole point of "analyze the
// frequently-used patches and grow the API to cover them".

import type {
  Block,
  Document,
  EquationBlock,
  ImageBlock,
  Paragraph,
  ParaStyle,
  Run,
  SectionPatch,
  SectionProps,
  Stylesheet,
  TableBlock,
  TableCell,
} from "@cw/shared";
import { DEFAULT_CHAR_STYLE, DEFAULT_PARA_STYLE, defaultStylesheet, resolveStyle } from "@cw/shared";
import { lit } from "./lit";

/** A model field the emitted code could not reproduce through the fluent API.
 *  Collected so it becomes a concrete backlog item for extending the builder. */
export interface UncoveredField {
  /** Human-readable location, e.g. `blocks[3].runs[1]` or `section`. */
  path: string;
  /** The model field name. */
  field: string;
  /** A sample of the dropped value (best-effort; large values are summarized). */
  value?: unknown;
  /** Why it could not be expressed / what would close the gap. */
  note?: string;
}

export interface CodegenOptions {
  /** Name of the generated function (default "buildDocument"). */
  functionName?: string;
  /** Import specifier for DocumentBuilder (default "@forevka/wordcanvas/builder"). */
  importFrom?: string;
}

export interface CodegenResult {
  /** The generated TypeScript module source. */
  code: string;
  /** Fields that could not be reproduced through the fluent API. */
  uncovered: UncoveredField[];
}

// ---------------------------------------------------------------------------
// Small value helpers

/** Structural equality for JSON-ish model leaf values (order-insensitive for
 *  object keys). Used to diff a concrete style against its resolved baseline. */
function deepEqual(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (typeof a !== typeof b) return false;
  if (a === null || b === null) return a === b;
  if (Array.isArray(a) || Array.isArray(b)) {
    if (!Array.isArray(a) || !Array.isArray(b) || a.length !== b.length) return false;
    return a.every((x, i) => deepEqual(x, b[i]));
  }
  if (typeof a === "object") {
    const oa = a as Record<string, unknown>;
    const ob = b as Record<string, unknown>;
    const ka = Object.keys(oa).filter((k) => oa[k] !== undefined);
    const kb = Object.keys(ob).filter((k) => ob[k] !== undefined);
    if (ka.length !== kb.length) return false;
    return ka.every((k) => deepEqual(oa[k], ob[k]));
  }
  return false;
}

/** Keys (union of both) where `target` differs structurally from `base`. */
function changedKeys(target: Record<string, unknown>, base: Record<string, unknown>): string[] {
  const keys = new Set<string>([...Object.keys(target), ...Object.keys(base)]);
  const out: string[] = [];
  for (const k of keys) {
    if (!deepEqual(target[k], base[k])) out.push(k);
  }
  return out;
}

/** Deep-equal against an empty/absent-friendly view: is `v` a "no-op" record
 *  (all values undefined or the record itself empty)? */
function isEmptyRecord(v: Record<string, unknown> | undefined): boolean {
  return !v || Object.keys(v).every((k) => v[k] === undefined);
}

// ---------------------------------------------------------------------------
// Line writer

class LineWriter {
  private readonly parts: string[] = [];
  private depth = 0;

  line(text = ""): void {
    this.parts.push(text ? "  ".repeat(this.depth) + text : "");
  }

  indent(): void {
    this.depth++;
  }

  dedent(): void {
    this.depth = Math.max(0, this.depth - 1);
  }

  toString(): string {
    return this.parts.join("\n");
  }
}

// ---------------------------------------------------------------------------
// CharStyle → run patch

/** Character fields we route through dedicated builder plumbing (registries) or
 *  cannot express — everything else is a plain patch field passed to .text(). */
const RUN_REGISTRY_FIELDS = ["footnoteRef", "endnoteRef", "fieldId", "sdtPath"] as const;

/** Build the `.text()` patch for a run: its CharStyle delta against `baseChar`,
 *  minus registry-backed markers (reported separately). Returns the patch plus
 *  the list of registry fields that were present (for the uncovered report). */
function runPatch(
  run: Run,
  baseChar: Record<string, unknown>,
): { patch: Record<string, unknown>; dropped: string[] } {
  const style = run.style as unknown as Record<string, unknown>;
  const patch: Record<string, unknown> = {};
  const dropped: string[] = [];
  for (const k of changedKeys(style, baseChar)) {
    if ((RUN_REGISTRY_FIELDS as readonly string[]).includes(k) && style[k] !== undefined) {
      dropped.push(k);
      continue;
    }
    patch[k] = style[k];
  }
  return { patch, dropped };
}

// ---------------------------------------------------------------------------
// Emitter

class Emitter {
  readonly writer = new LineWriter();
  readonly uncovered: UncoveredField[] = [];
  /** The stylesheet the GENERATED code will resolve baselines against — the doc's
   *  own, or the default gallery `DocumentBuilder.create()` installs when the doc
   *  carries none (so our deltas match what the rebuilt document resolves). */
  private readonly sheet: Stylesheet;
  private readonly charDefault: Record<string, unknown>;
  private readonly paraDefault: Record<string, unknown>;

  constructor(private readonly doc: Document) {
    this.sheet = doc.stylesheet ?? defaultStylesheet();
    const rd = resolveStyle(this.sheet, this.sheet.defaultStyleId);
    this.charDefault = { ...DEFAULT_CHAR_STYLE, ...rd.char };
    this.paraDefault = { ...DEFAULT_PARA_STYLE, ...rd.para };
  }

  private report(path: string, field: string, value?: unknown, note?: string): void {
    const entry: UncoveredField = { path, field };
    if (value !== undefined) entry.value = summarize(value);
    if (note !== undefined) entry.note = note;
    this.uncovered.push(entry);
  }

  /** Resolved char/para baseline for a paragraph — the exact state the builder
   *  reaches after `.paragraph()` + optional `.withStyle(namedStyle)`. */
  private baselineFor(namedStyle: string | undefined): { char: Record<string, unknown>; para: Record<string, unknown> } {
    const rs = namedStyle ? resolveStyle(this.sheet, namedStyle) : { char: {}, para: {} };
    return {
      char: { ...this.charDefault, ...rs.char },
      para: { ...this.paraDefault, ...rs.para },
    };
  }

  emit(opts: CodegenOptions): void {
    const fn = opts.functionName ?? "buildDocument";
    const from = opts.importFrom ?? "@forevka/wordcanvas/builder";
    const w = this.writer;
    w.line(`import { DocumentBuilder } from ${JSON.stringify(from)};`);
    w.line(`import type { Document } from ${JSON.stringify(from)};`);
    w.line("");
    w.line(`export function ${fn}(): Document {`);
    w.indent();
    this.emitCreate();
    this.emitDocLevel();
    for (const block of this.doc.blocks) this.emitBlock("b", block, "blocks");
    // Deferred: the body section's line numbering (see emitPageSetup) — applied
    // only after every section-break paragraph has been compiled, so it stays on
    // the final section instead of being carried backward onto earlier ones.
    if (this.doc.section.lineNumbering !== undefined) {
      w.line(`b.pageSetup({ lineNumbering: ${lit(this.doc.section.lineNumbering, "  ")} });`);
    }
    w.line("return b.build();");
    w.dedent();
    w.line("}");
  }

  /** `const b = DocumentBuilder.create({...})` + page geometry via `.pageSetup`. */
  private emitCreate(): void {
    const w = this.writer;
    const createOpts: Record<string, unknown> = {};
    const sheet = this.doc.stylesheet;
    // Reproduce the stylesheet verbatim when it differs from the builder's default
    // gallery (so run/para baselines resolve identically); omit it otherwise.
    if (sheet && !deepEqual(sheet, defaultStylesheet())) createOpts.stylesheet = sheet;
    w.line(`const b = DocumentBuilder.create(${isEmptyRecord(createOpts) ? "" : lit(createOpts, "  ")});`);
    this.emitPageSetup(this.doc.section);
  }

  private emitPageSetup(section: SectionProps): void {
    const w = this.writer;
    const setup: Record<string, unknown> = {
      pageSize: { pageWidthPx: section.pageWidthPx, pageHeightPx: section.pageHeightPx },
      margins: section.marginPx,
    };
    if (section.columns) {
      setup.columns = { count: section.columns.count, gapPx: section.columns.gapPx };
      if (section.columns.sep) this.report("section", "columns.sep", true, "column separator; no pageSetup option");
      if (section.columns.cols) this.report("section", "columns.cols", section.columns.cols, "explicit per-column widths; no pageSetup option");
    }
    if (section.headerDistancePx !== undefined) setup.headerDistancePx = section.headerDistancePx;
    if (section.footerDistancePx !== undefined) setup.footerDistancePx = section.footerDistancePx;
    if (section.pageNumberStart !== undefined) setup.pageNumberStart = section.pageNumberStart;
    if (section.breakType !== undefined) setup.breakType = section.breakType;
    // NB: the body section's `lineNumbering` is emitted LAST (see emit()), not here —
    // set early it would leak onto every earlier \sectionBreak section, which
    // compileSectionPatch carries the current section value forward into.
    w.line(`b.pageSetup(${lit(setup, "  ")});`);
    if (section.pageColorHex !== undefined) this.report("section", "pageColorHex", section.pageColorHex, "page background color; no builder method");
    if (section.pageBorders !== undefined) this.report("section", "pageBorders", section.pageBorders, "page border box; no builder method");
  }

  /** defaultTabStop, list definitions, table styles, header/footer bands, and the
   *  registries the fluent API cannot reconstruct (fields/sdts/notes/bookmarks). */
  private emitDocLevel(): void {
    const w = this.writer;
    const d = this.doc;
    if (d.defaultTabStopPx !== undefined) w.line(`b.defaultTabStop(${lit(d.defaultTabStopPx)});`);
    for (const [id, def] of Object.entries(d.lists ?? {})) {
      w.line(`b.listDefinition(${lit(id)}, ${lit({ levels: def.levels }, "  ")});`);
    }
    for (const def of Object.values(d.tableStyles ?? {})) {
      w.line(`b.tableStyle(${lit(def, "  ")});`);
    }
    for (const band of ["header", "footer", "headerFirst", "footerFirst", "headerEven", "footerEven"] as const) {
      const story = d.section[band];
      if (!story || story.length === 0) continue;
      const variant = band.endsWith("First") ? "first" : band.endsWith("Even") ? "even" : "default";
      const method = band.startsWith("header") ? "header" : "footer";
      const optArg = variant === "default" ? "" : `, { variant: ${lit(variant)} }`;
      w.line(`b.${method}((s) => {`);
      w.indent();
      for (const block of story) this.emitBlock("s", block, `section.${band}`);
      w.dedent();
      w.line(`}${optArg});`);
    }
    // Registries with no fluent reconstruction path in v1 — reported, not dropped.
    for (const id of Object.keys(d.fields ?? {})) this.report("fields", id, undefined, "inline/region field; reconstruct via .field()/.customField() (not yet automated)");
    for (const id of Object.keys(d.sdts ?? {})) this.report("sdts", id, undefined, "content control; reconstruct via .contentControl()/.contentControlRange() (not yet automated)");
    for (const id of Object.keys(d.footnotes ?? {})) this.report("footnotes", id, undefined, "footnote body; reconstruct via .footnote() (not yet automated)");
    for (const id of Object.keys(d.endnotes ?? {})) this.report("endnotes", id, undefined, "endnote body; reconstruct via .endnote() (not yet automated)");
    for (const name of Object.keys(d.bookmarks ?? {})) this.report("bookmarks", name, undefined, "bookmark range; reconstruct via .bookmark()/.bookmarkRange() (not yet automated)");
    if (d.tocInstruction !== undefined) this.report("document", "tocInstruction", d.tocInstruction, "table of contents; reconstruct via .tableOfContents() (not yet automated)");
    if (d.compatSettings !== undefined) this.report("document", "compatSettings", d.compatSettings, "settings.xml compatibility triples; no builder method");
    if (d.docId !== undefined) this.report("document", "docId", d.docId, "persistent w15:docId; no builder method");
  }

  // ---- blocks -------------------------------------------------------------

  private emitBlock(varName: string, block: Block, path: string): void {
    switch (block.kind) {
      case "paragraph":
        this.emitParagraph(varName, block, path);
        break;
      case "table":
        this.emitTable(varName, block, path);
        break;
      case "image":
        this.emitImage(varName, block, path);
        break;
      case "equation":
        this.emitEquation(varName, block, path);
        break;
      case "custom":
        this.report(path, "custom", (block as { customType?: string }).customType, "custom block; no builder method");
        break;
    }
  }

  private emitParagraph(varName: string, para: Paragraph, path: string): void {
    const w = this.writer;
    // Block-identity + membership markers the builder cannot set.
    if (para.paraId !== undefined) this.report(path, "paraId", para.paraId, "Word w14:paraId identity; no builder method");
    if (para.textId !== undefined) this.report(path, "textId", para.textId, "Word w14:textId identity; no builder method");
    if (para.fieldId !== undefined) this.report(path, "fieldId", para.fieldId, "block-level field membership; not yet automated");
    if (para.sdtPath !== undefined) this.report(path, "sdtPath", para.sdtPath, "block-level content control; not yet automated");

    // Leading page/column breaks are pending flags on the builder — emit before.
    if (para.style.pageBreakBefore) w.line(`${varName}.pageBreak();`);
    if (para.style.columnBreakBefore) w.line(`${varName}.columnBreak();`);

    const base = this.baselineFor(para.style.namedStyle);
    const chain: string[] = [];
    chain.push("paragraph()");
    if (para.style.namedStyle !== undefined) chain.push(`withStyle(${lit(para.style.namedStyle)})`);

    // Runs.
    for (let i = 0; i < para.runs.length; i++) {
      const run = para.runs[i]!;
      const { patch, dropped } = runPatch(run, base.char);
      for (const f of dropped) this.report(`${path}.runs[${i}]`, f, (run.style as unknown as Record<string, unknown>)[f], "registry-backed inline marker; not yet automated");
      chain.push(isEmptyRecord(patch) ? `text(${lit(run.text)})` : `text(${lit(run.text)}, ${lit(patch, "  ")})`);
    }

    // Paragraph properties (delta vs the resolved baseline).
    this.emitParaStyle(chain, para.style, base.para, path);

    // A section break is a property on the LAST paragraph of a section.
    if (para.style.sectionBreak) this.emitSectionBreak(chain, para.style.sectionBreak, path);

    w.line(`${varName}.${chain.join(".")};`);
  }

  /** Append paragraph-style method calls to `chain` for every changed field. */
  private emitParaStyle(chain: string[], style: ParaStyle, base: Record<string, unknown>, path: string): void {
    const s = style as unknown as Record<string, unknown>;
    const changed = new Set(changedKeys(s, base));
    // Fields consumed here but NOT via a scalar method — accounted for explicitly.
    const handledElsewhere = ["namedStyle", "pageBreakBefore", "columnBreakBefore", "sectionBreak", "tocEntry"];
    for (const f of handledElsewhere) changed.delete(f);

    if (changed.has("align")) chain.push(`align(${lit(style.align)})`);
    changed.delete("align");

    // spacing() — auto-spacing (w:beforeAutospacing) and the explicit px value are
    // independent: `beforeAuto:true` bakes an approximate value, but `beforeAuto:false`
    // (Word's w:beforeAutospacing="0") keeps the REAL w:before, so both must be emitted.
    const spacing: Record<string, unknown> = {};
    if (style.spaceBeforeAuto) {
      spacing.beforeAuto = true;
    } else {
      if (changed.has("spaceBeforeAuto") && style.spaceBeforeAuto === false) spacing.beforeAuto = false;
      if (changed.has("spaceBeforePx")) spacing.before = style.spaceBeforePx;
    }
    if (style.spaceAfterAuto) {
      spacing.afterAuto = true;
    } else {
      if (changed.has("spaceAfterAuto") && style.spaceAfterAuto === false) spacing.afterAuto = false;
      if (changed.has("spaceAfterPx")) spacing.after = style.spaceAfterPx;
    }
    // Line spacing: a fixed rule (lineRule + lineHeightPx) OR the multiplier
    // (lineHeight). `lineHeightPx` is only meaningful alongside a rule.
    if (style.lineRule !== undefined && (changed.has("lineRule") || changed.has("lineHeightPx"))) {
      spacing.lineRule = style.lineRule;
      spacing.lineHeightPx = style.lineHeightPx;
    } else if (changed.has("lineHeight")) {
      spacing.lineHeight = style.lineHeight;
    }
    for (const k of ["spaceBeforePx", "spaceAfterPx", "spaceBeforeAuto", "spaceAfterAuto", "lineHeight", "lineRule", "lineHeightPx"]) changed.delete(k);
    if (!isEmptyRecord(spacing)) chain.push(`spacing(${lit(spacing, "  ")})`);

    // indent()
    const indent: Record<string, unknown> = {};
    if (changed.has("indentLeftPx")) indent.left = style.indentLeftPx;
    if (changed.has("indentRightPx")) indent.right = style.indentRightPx;
    if (changed.has("indentFirstLinePx")) indent.firstLine = style.indentFirstLinePx;
    for (const k of ["indentLeftPx", "indentRightPx", "indentFirstLinePx"]) changed.delete(k);
    if (!isEmptyRecord(indent)) chain.push(`indent(${lit(indent, "  ")})`);

    // Toggle / value methods that map 1:1.
    const boolMethods: Record<string, string> = {
      keepWithNext: "keepWithNext",
      keepLinesTogether: "keepTogether",
      contextualSpacing: "contextualSpacing",
      widowControl: "widowControl",
      suppressLineNumbers: "suppressLineNumbers",
      mirrorIndents: "mirrorIndents",
      adjustRightInd: "adjustRightInd",
      snapToGrid: "snapToGrid",
      suppressAutoHyphens: "suppressAutoHyphens",
      kinsoku: "kinsoku",
      overflowPunct: "overflowPunct",
      wordWrap: "wordWrap",
      topLinePunct: "topLinePunct",
      autoSpaceDE: "autoSpaceDE",
      autoSpaceDN: "autoSpaceDN",
    };
    for (const [field, method] of Object.entries(boolMethods)) {
      if (!changed.has(field)) continue;
      chain.push(`${method}(${lit(s[field])})`);
      changed.delete(field);
    }

    if (changed.has("direction")) { chain.push(`direction(${lit(style.direction)})`); changed.delete("direction"); }
    if (changed.has("outlineLevel")) { chain.push(`outlineLevel(${lit(style.outlineLevel)})`); changed.delete("outlineLevel"); }
    if (changed.has("textAlignment")) { chain.push(`textAlignment(${lit(style.textAlignment)})`); changed.delete("textAlignment"); }
    if (changed.has("tabStops")) { chain.push(`tabStops(${lit(style.tabStops, "  ")})`); changed.delete("tabStops"); }

    if (changed.has("shadingCleared") && style.shadingCleared) { chain.push("clearShading()"); changed.delete("shadingCleared"); changed.delete("shading"); }
    else if (changed.has("shading")) { chain.push(`shading(${lit(style.shading)})`); changed.delete("shading"); changed.delete("shadingCleared"); }
    if (changed.has("bordersCleared") && style.bordersCleared) { chain.push("clearBorders()"); changed.delete("bordersCleared"); changed.delete("borders"); }
    else if (changed.has("borders")) { chain.push(`borders(${lit(style.borders, "  ")})`); changed.delete("borders"); changed.delete("bordersCleared"); }

    if (changed.has("listCleared") && style.listCleared) { chain.push("clearList()"); changed.delete("listCleared"); changed.delete("list"); }
    else if (changed.has("list") && style.list) {
      // Per-paragraph list membership — reproduced exactly (rich runs, named style,
      // spacing all preserved) via the single-paragraph .listItem() builder method.
      chain.push(`listItem(${lit(style.list.listId)}${style.list.level ? `, ${lit(style.list.level)}` : ""})`);
      changed.delete("list");
      changed.delete("listCleared");
    }

    for (const f of changed) this.report(path, f, s[f], "paragraph field with no fluent method");
  }

  private emitSectionBreak(chain: string[], sb: { type: string; props: SectionPatch }, path: string): void {
    const opts: Record<string, unknown> = {};
    const p = sb.props;
    if (sb.type !== "nextPage") opts.breakType = sb.type;
    if (p.pageWidthPx !== undefined && p.pageHeightPx !== undefined) opts.pageSize = { pageWidthPx: p.pageWidthPx, pageHeightPx: p.pageHeightPx };
    if (p.marginPx !== undefined) opts.margins = p.marginPx;
    if (p.columns !== undefined) opts.columns = p.columns === null ? null : { count: p.columns.count, gapPx: p.columns.gapPx };
    if (p.pageNumberStart !== undefined) opts.pageNumberStart = p.pageNumberStart;
    if (p.lineNumbering !== undefined) opts.lineNumbering = p.lineNumbering;
    // Band stories on a section break need callbacks — reported for v1.
    for (const band of ["header", "footer", "headerFirst", "footerFirst", "headerEven", "footerEven"] as const) {
      if (p[band] !== undefined) this.report(path, `sectionBreak.${band}`, undefined, "per-section band story on a section break; not yet automated");
    }
    if (p.pageColorHex !== undefined) this.report(path, "sectionBreak.pageColorHex", p.pageColorHex, "no sectionBreak option");
    if (p.pageBorders !== undefined) this.report(path, "sectionBreak.pageBorders", p.pageBorders, "no sectionBreak option");
    if (p.headerDistancePx !== undefined) this.report(path, "sectionBreak.headerDistancePx", p.headerDistancePx, "no sectionBreak option");
    if (p.footerDistancePx !== undefined) this.report(path, "sectionBreak.footerDistancePx", p.footerDistancePx, "no sectionBreak option");
    chain.push(`sectionBreak(${isEmptyRecord(opts) ? "" : lit(opts, "  ")})`);
  }

  private emitImage(varName: string, img: ImageBlock, path: string): void {
    const w = this.writer;
    const opts: Record<string, unknown> = { widthPx: img.widthPx, heightPx: img.heightPx };
    if (img.align !== "left") opts.align = img.align;
    if (img.wrap !== undefined) opts.wrap = img.wrap;
    if (img.crop !== undefined) opts.crop = img.crop;
    if (img.externalSrc !== undefined) opts.linked = true;
    if (img.anchor !== undefined) this.report(path, "anchor", img.anchor, "absolutely-positioned image; no builder .image() option");
    const src = img.externalSrc ?? img.src;
    w.line(`${varName}.image(${lit(src)}, ${lit(opts, "  ")});`);
  }

  private emitEquation(varName: string, eq: EquationBlock, path: string): void {
    // Block equations reconstruct from MathML; a reverse AST→MathML serializer is
    // not wired here yet, so report and skip (visible only as a gap, not garbage).
    const w = this.writer;
    this.report(path, "equation", undefined, "block equation; reconstruct via .equationMathml() (AST serializer not yet wired)");
    w.line(`// TODO: block equation at ${path} — reconstruct via ${varName}.equationMathml(...)`);
  }

  private emitTable(varName: string, table: TableBlock, path: string): void {
    const w = this.writer;
    if (table.fieldId !== undefined) this.report(path, "fieldId", table.fieldId, "table field membership; not yet automated");
    if (table.sdtPath !== undefined) this.report(path, "sdtPath", table.sdtPath, "table content control; not yet automated");

    const opts: Record<string, unknown> = {};
    if (table.colFractions !== undefined) opts.colFractions = table.colFractions;
    if (table.widthMode !== undefined && table.widthMode !== "fixed") opts.widthMode = table.widthMode;
    if (table.defaultBorders !== undefined) opts.borders = table.defaultBorders;
    if (table.defaultShading !== undefined) opts.shading = table.defaultShading;
    if (table.defaultCellMargin !== undefined) opts.cellMargin = table.defaultCellMargin;
    if (table.indentPx !== undefined && table.indentPx !== 0) opts.indent = table.indentPx;
    if (table.bidiVisual) opts.bidiVisual = true;
    if (table.overlap !== undefined) opts.overlap = table.overlap;
    if (table.caption !== undefined) opts.caption = table.caption;
    if (table.description !== undefined) opts.description = table.description;
    if (table.styleId !== undefined) {
      opts.styleId = table.styleId;
      if (table.condOverrides !== undefined) opts.condOverrides = table.condOverrides;
    }
    if (table.preferredWidth !== undefined) this.report(path, "preferredWidth", table.preferredWidth, "table preferred width; no .table() option");
    if (table.align !== undefined) this.report(path, "align", table.align, "table alignment; no .table() option");

    const optArg = isEmptyRecord(opts) ? "" : `, ${lit(opts, "    ")}`;
    w.line(`${varName}.table((t) => {`);
    w.indent();
    table.rows.forEach((row, ri) => {
      const rowOpts: Record<string, unknown> = {};
      if (row.props?.height) { rowOpts.height = row.props.height.value; rowOpts.heightRule = row.props.height.rule; }
      if (row.props?.cantSplit) rowOpts.cantSplit = true;
      if (row.props?.repeatHeader) rowOpts.header = true;
      const rowOptArg = isEmptyRecord(rowOpts) ? "" : `, ${lit(rowOpts, "    ")}`;
      w.line(`t.row((r) => {`);
      w.indent();
      row.cells.forEach((cell, ci) => this.emitCell("r", cell, `${path}.rows[${ri}].cells[${ci}]`));
      w.dedent();
      w.line(`}${rowOptArg});`);
    });
    w.dedent();
    w.line(`}${optArg});`);
  }

  private emitCell(rowVar: string, cell: TableCell, path: string): void {
    const w = this.writer;
    const opts: Record<string, unknown> = {};
    if (cell.colSpan !== undefined && cell.colSpan > 1) opts.colSpan = cell.colSpan;
    if (cell.rowSpan !== undefined && cell.rowSpan > 1) opts.rowSpan = cell.rowSpan;
    if (cell.shading !== undefined) opts.shading = cell.shading;
    else if (cell.shadingCleared) opts.shadingCleared = true;
    if (cell.borders !== undefined) opts.borders = cell.borders;
    if (cell.margin !== undefined) opts.margin = cell.margin;
    if (cell.preferredWidth !== undefined) opts.preferredWidth = cell.preferredWidth;
    if (cell.vAlign !== undefined && cell.vAlign !== "top") opts.vAlign = cell.vAlign;
    if (cell.textDirection !== undefined && cell.textDirection !== "lrTb") opts.textDirection = cell.textDirection;
    if (cell.noWrap) opts.noWrap = true;
    if (cell.fitText) opts.fitText = true;
    if (cell.hideMark) opts.hideMark = true;
    const optArg = isEmptyRecord(opts) ? "" : `, ${lit(opts, "      ")}`;
    w.line(`r.cell((c) => {`);
    w.indent();
    for (const block of cell.blocks) this.emitBlock("c", block, `${path}.blocks`);
    w.dedent();
    w.line(`}${optArg});`);
  }
}

/** Summarize a value for the uncovered report: keep small ones, elide large. */
function summarize(v: unknown): unknown {
  const json = (() => {
    try {
      return JSON.stringify(v);
    } catch {
      return "";
    }
  })();
  if (json && json.length <= 120) return v;
  return `${(json || String(v)).slice(0, 100)}…`;
}

/** Generate DocumentBuilder TypeScript that reconstructs `doc`. */
export function emitBuilderCode(doc: Document, opts: CodegenOptions = {}): CodegenResult {
  const e = new Emitter(doc);
  e.emit(opts);
  return { code: e.writer.toString() + "\n", uncovered: e.uncovered };
}
