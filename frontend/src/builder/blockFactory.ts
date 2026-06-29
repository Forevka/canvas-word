// Shared builder context + concrete block factories. The model keeps runs
// CONCRETE (every CharStyle/ParaStyle field present on every run/paragraph),
// so the context owns the resolved defaults — derived from the stylesheet's
// default style over editor-matching baselines — and every factory spreads
// them. Modeled on the run/para closures in model/sampleDoc.ts.

import type { CharStyle, Document, EquationBlock, ImageBlock, MathEquation, ParaStyle, Paragraph, Run, Stylesheet } from "@cw/shared";
import { createIdGenerator, DEFAULT_CHAR_STYLE, DEFAULT_PARA_STYLE, resolveStyle, styleById, type IdGenerator } from "@cw/shared";
import { builtinTableStyles, type TableStylePreset } from "./tableStyles";

export interface BuilderWarning {
  code: string;
  message: string;
}

/** Baseline concrete styles (match the editor's sample-doc/import defaults). */
const BASE_CHAR: CharStyle = DEFAULT_CHAR_STYLE;
const BASE_PARA: ParaStyle = DEFAULT_PARA_STYLE;

/** A concrete empty paragraph matching the document's default style — the
 *  caret home injected when a document arrives with no blocks (builder build()
 *  and the editor's programmatic setDocument both guard with this). */
export function emptyParagraphFor(doc: Document, id: string): Paragraph {
  const resolved = doc.stylesheet ? resolveStyle(doc.stylesheet, doc.stylesheet.defaultStyleId) : { char: {}, para: {} };
  return {
    kind: "paragraph",
    id,
    revision: 0,
    runs: [{ text: "", style: { ...BASE_CHAR, ...resolved.char } }],
    style: { ...BASE_PARA, ...resolved.para },
  };
}

/** Everything builder scopes share: id minting, the (mutable) document under
 *  construction, resolved defaults, and the deduplicating warning list. */
export class BuilderContext {
  readonly ids: IdGenerator;
  readonly doc: Document;
  readonly warnings: BuilderWarning[] = [];
  private readonly seenWarnings = new Set<string>();
  /** Concrete defaults for new runs/paragraphs (default style resolved). */
  charDefault: CharStyle;
  paraDefault: ParaStyle;
  /** Footnote marker counter — monotonic in insertion order (.footnote()). */
  private footnoteCounter = 0;
  /** Builder-only table-style presets (seeded with built-ins; .tableStylePreset adds). */
  readonly tableStyles = builtinTableStyles();
  /** Per-object provenance of author-supplied style keys: a run → the CharStyle
   *  keys, a paragraph block → the ParaStyle keys the author explicitly set. Lets
   *  table-style baking preserve an explicit value even when it equals the
   *  resolved default (issue #30). Keyed by object identity, which survives the
   *  grid rebuild in bakeTableStyleRows. */
  readonly explicitCharKeys = new WeakMap<object, ReadonlySet<string>>();
  readonly explicitParaKeys = new WeakMap<object, ReadonlySet<string>>();

  constructor(doc: Document, idSeed?: string) {
    this.doc = doc;
    this.ids = createIdGenerator(idSeed ?? `bld${Math.random().toString(36).slice(2, 6)}`);
    const resolved = doc.stylesheet ? resolveStyle(doc.stylesheet, doc.stylesheet.defaultStyleId) : { char: {}, para: {} };
    this.charDefault = { ...BASE_CHAR, ...resolved.char };
    this.paraDefault = { ...BASE_PARA, ...resolved.para };
  }

  /** Next footnote marker number (1-based, insertion order). */
  nextFootnoteNumber(): number {
    return ++this.footnoteCounter;
  }

  /** Recompute the run/paragraph defaults from the current stylesheet's default
   *  style — called when .defaultStyle() changes which style is the baseline. */
  refreshDefaults(): void {
    const resolved = this.doc.stylesheet ? resolveStyle(this.doc.stylesheet, this.doc.stylesheet.defaultStyleId) : { char: {}, para: {} };
    this.charDefault = { ...BASE_CHAR, ...resolved.char };
    this.paraDefault = { ...BASE_PARA, ...resolved.para };
  }

  /** A registered table-style preset by name, or undefined (+ warning). */
  tableStyle(name: string): TableStylePreset | undefined {
    const preset = this.tableStyles.get(name);
    if (!preset) this.warn(`table-style-missing:${name}`, `Table style "${name}" is not registered — .table({ style: "${name}" }) used no preset.`);
    return preset;
  }

  warn(code: string, message: string): void {
    if (this.seenWarnings.has(code)) return;
    this.seenWarnings.add(code);
    this.warnings.push({ code, message });
  }

  run(text: string, patch: Partial<CharStyle> = {}): Run {
    const run: Run = { text, style: { ...this.charDefault, ...patch } };
    // Record which char keys the author explicitly supplied via `patch`, keyed by
    // run identity. Table-style baking (bakeTableStyleRows/applyBandProps) consults
    // this provenance so an explicit value that EQUALS the resolved default still
    // survives a conflicting band (#30). Centralizing here covers every run-authoring
    // surface — data-driven cells, StoryBuilder callbacks, ParagraphBuilder.text,
    // lists, fields — so both cell paths behave identically (#45). Harmless for
    // non-table runs: the map is read only when a table style is baked.
    const keys = Object.keys(patch);
    if (keys.length > 0) this.explicitCharKeys.set(run, new Set(keys));
    return run;
  }

  paragraph(runs: Run[], patch: Partial<ParaStyle> = {}): Paragraph {
    return {
      kind: "paragraph",
      id: this.ids.next(),
      revision: 0,
      runs: runs.length > 0 ? runs : [this.run("")],
      style: { ...this.paraDefault, ...patch },
    };
  }

  image(src: string, widthPx: number, heightPx: number, align: ImageBlock["align"], wrap?: ImageBlock["wrap"], crop?: ImageBlock["crop"]): ImageBlock {
    const img: ImageBlock = { kind: "image", id: this.ids.next(), revision: 0, src, widthPx, heightPx, align };
    if (wrap) img.wrap = wrap;
    if (crop) img.crop = crop;
    return img;
  }

  /** A display (block) equation block from a ready MathEquation AST. */
  equation(equation: MathEquation, align?: EquationBlock["align"]): EquationBlock {
    const eq: EquationBlock = { kind: "equation", id: this.ids.next(), revision: 0, equation };
    if (align) eq.align = align;
    return eq;
  }

  /** The stylesheet, creating an empty one on demand for style registration. */
  stylesheet(): Stylesheet {
    if (!this.doc.stylesheet) this.doc.stylesheet = { styles: [], defaultStyleId: "Normal" };
    return this.doc.stylesheet;
  }

  /** Resolved char+para templates for a named style, or undefined (+ warning). */
  lookupStyle(id: string): { char: Partial<CharStyle>; para: Partial<ParaStyle> } | undefined {
    const sheet = this.doc.stylesheet;
    if (!sheet || !styleById(sheet, id)) {
      this.warn(`style-missing:${id}`, `Named style "${id}" is not in the stylesheet — .withStyle("${id}") was ignored.`);
      return undefined;
    }
    return resolveStyle(sheet, id);
  }
}
