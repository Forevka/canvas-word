// Block-scope base for the fluent builder: anything that holds a Block story —
// the document body (DocumentBuilder extends this), header/footer bands, and
// table cells — shares this surface. Methods append eagerly and return `this`
// (or a ParagraphBuilder that delegates back), so chains never need a seal step.

import type { Block, CharStyle, EquationBlock, Paragraph, ParaStyle, SdtProps } from "@cw/shared";
import { DEFAULT_BULLET_LIST_ID, DEFAULT_NUMBER_LIST_ID, defaultListDefinition, textOfRuns } from "@cw/shared";
import type { BuilderContext } from "./blockFactory";
import { equationFromLatex, equationFromMathml } from "./mathInput";
import { bytesToDataUrl } from "./media";
import { ParagraphBuilder } from "./paragraphBuilder";
import { TableBuilder, type CellContent, type TableOptions } from "./tableBuilder";

export interface ImageOptions {
  /** Intrinsic display size — required: the builder runs in Node too, where
   *  there is no DOM to measure an image (auto-measure is on the roadmap). */
  widthPx: number;
  heightPx: number;
  align?: "left" | "center" | "right";
  /** 'block' (default): own line. 'square': floats per align, text wraps. */
  wrap?: "block" | "square";
  /** Crop insets (OOXML a:srcRect), each a 0..1 fraction trimmed off that edge —
   *  so widthPx/heightPx describe the cropped box. Absent = no crop. */
  crop?: { left: number; top: number; right: number; bottom: number };
  /** Linked ("Link to File") image: `src` must be an http(s) URL whose bytes stay
   *  OUTSIDE the document. Export re-emits it as `a:blip r:link` + an External
   *  relationship rather than packing bytes into word/media. Ignored for raw-byte
   *  sources (those are always embedded as data:). */
  linked?: boolean;
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
  /** Default nesting level for all items (per-item `level` overrides). */
  level?: number;
}

export interface EquationOptions {
  /** Horizontal placement of the display equation (default "center"). */
  align?: EquationBlock["align"];
  /** Uniform size multiplier (drag-to-resize), default 1. Clamped to [0.25, 4]. */
  scale?: number;
}

export class StoryBuilder {
  protected readonly ctx: BuilderContext;
  protected readonly blocks: Block[];
  private pendingPageBreak = false;
  private pendingColumnBreak = false;

  constructor(ctx: BuilderContext, blocks: Block[]) {
    this.ctx = ctx;
    this.blocks = blocks;
  }

  /** Append a block, honoring a pending page/column break(). Only paragraphs
   *  carry the break flag, so a break before a table/image lands on an injected
   *  empty paragraph (same trick as the docx import mapper). */
  protected push(block: Block): void {
    if (this.pendingPageBreak || this.pendingColumnBreak) {
      if (block.kind === "paragraph") {
        if (this.pendingPageBreak) block.style.pageBreakBefore = true;
        if (this.pendingColumnBreak) block.style.columnBreakBefore = true;
      } else {
        const patch: Partial<ParaStyle> = {};
        if (this.pendingPageBreak) patch.pageBreakBefore = true;
        if (this.pendingColumnBreak) patch.columnBreakBefore = true;
        this.blocks.push(this.ctx.paragraph([], patch));
      }
      this.pendingPageBreak = false;
      this.pendingColumnBreak = false;
    }
    this.blocks.push(block);
  }

  /** Start a paragraph. The returned scope styles THIS paragraph and delegates
   *  block-starting calls back here, so the chain continues seamlessly. */
  paragraph(text?: string, style?: Partial<CharStyle>): ParagraphBuilder<this> {
    const runs = text !== undefined ? [this.ctx.run(text, style ?? {})] : [];
    const para = this.ctx.paragraph(runs);
    this.push(para);
    return new ParagraphBuilder(this, this.ctx, para);
  }

  /** Data-driven table: 2D array of cell text/specs (the common case). */
  table(rows: CellContent[][], opts?: TableOptions): this;
  /** Structural table: callback scope for spans, nested blocks, full control. */
  table(build: (t: TableBuilder) => void, opts?: TableOptions): this;
  table(arg: CellContent[][] | ((t: TableBuilder) => void), opts?: TableOptions): this {
    const t = new TableBuilder(this.ctx, opts);
    if (typeof arg === "function") arg(t);
    else t.rows(arg);
    this.push(t.toBlock());
    return this;
  }

  /** Insert an image from a URL (https:/data:) or raw bytes (inlined as data:). */
  image(src: string | { data: Uint8Array | ArrayBuffer; mime: string }, opts: ImageOptions): this {
    if (!Number.isFinite(opts?.widthPx) || !Number.isFinite(opts?.heightPx) || opts.widthPx <= 0 || opts.heightPx <= 0) {
      // Warn-and-skip (not throw) — the builder's uniform invalid-input contract:
      // a fluent chain keeps going and the problem lands in .warnings.
      this.ctx.warn("image-size-invalid", ".image() requires positive widthPx and heightPx (no DOM to auto-measure in Node); image skipped.");
      return this;
    }
    const url = typeof src === "string" ? src : bytesToDataUrl(src.data, src.mime);
    // Linked only makes sense for an external URL source (raw bytes are embedded).
    const externalSrc = opts.linked && typeof src === "string" ? url : undefined;
    this.push(this.ctx.image(url, opts.widthPx, opts.heightPx, opts.align ?? "left", opts.wrap, opts.crop, externalSrc));
    return this;
  }

  /** One paragraph per item, marked as list members. Definitions resolve to the
   *  template's (via listId), or the editor-default bullet/number definitions
   *  registered on demand. */
  list(items: (string | ListItem)[], opts: ListOptions = {}): this {
    const kind = opts.kind ?? "bullet";
    const lists = (this.ctx.doc.lists ??= {});
    let listId = opts.listId;
    if (listId !== undefined) {
      if (!lists[listId]) {
        this.ctx.warn(
          `list-missing:${listId}`,
          `List definition "${listId}" is not in the document — a default ${kind} definition was registered under that id.`,
        );
        lists[listId] = { ...defaultListDefinition(kind === "number" ? "decimal" : "bullet"), id: listId };
      }
    } else {
      listId = kind === "number" ? DEFAULT_NUMBER_LIST_ID : DEFAULT_BULLET_LIST_ID;
      if (!lists[listId]) lists[listId] = defaultListDefinition(kind === "number" ? "decimal" : "bullet");
    }
    for (const item of items) {
      const it = typeof item === "string" ? { text: item } : item;
      const level = Math.max(0, Math.min(8, it.level ?? opts.level ?? 0));
      this.push(this.ctx.paragraph([this.ctx.run(it.text, it.style ?? {})], { list: { listId, level } }));
    }
    return this;
  }

  /** A display (block) equation from a LaTeX source (e.g. `\frac{a}{b}`) — its own
   *  centered line. Uses the same LaTeX→math parser as the visual editor. */
  equation(latex: string, opts: EquationOptions = {}): this {
    this.push(this.ctx.equation(equationFromLatex(latex, true), opts.align, opts.scale));
    return this;
  }

  /** A display (block) equation from a presentation-MathML string. */
  equationMathml(mathml: string, opts: EquationOptions = {}): this {
    this.push(this.ctx.equation(equationFromMathml(mathml, true), opts.align, opts.scale));
    return this;
  }

  bulletList(items: (string | ListItem)[]): this {
    return this.list(items, { kind: "bullet" });
  }

  numberedList(items: (string | ListItem)[]): this {
    return this.list(items, { kind: "number" });
  }

  /** The NEXT block starts a new page. */
  pageBreak(): this {
    this.pendingPageBreak = true;
    return this;
  }

  /** The NEXT block starts a new newspaper column (Ctrl+Shift+Enter). A trailing
   *  column break with no following block is a no-op, matching Word. */
  columnBreak(): this {
    this.pendingColumnBreak = true;
    return this;
  }

  /** Bookmark everything the callback appends to this story (multi-paragraph
   *  span). For a precise inline span use ParagraphBuilder.bookmark(name, text). */
  bookmarkRange(name: string, build: (s: this) => void): this {
    const startIdx = this.blocks.length;
    build(this);
    const added = this.blocks.slice(startIdx).filter((b): b is Paragraph => b.kind === "paragraph");
    if (added.length === 0) {
      this.ctx.warn(`bookmark-empty:${name}`, `bookmarkRange("${name}") wrapped no paragraphs — no bookmark was recorded.`);
      return this;
    }
    const first = added[0]!;
    const last = added[added.length - 1]!;
    const bookmarks = (this.ctx.doc.bookmarks ??= {});
    bookmarks[name] = {
      start: { blockId: first.id, offset: 0 },
      end: { blockId: last.id, offset: textOfRuns(last.runs).length },
    };
    return this;
  }

  /** Wrap every block the callback appends in a BLOCK-LEVEL content control
   *  (OOXML block-level w:sdt) — the whole-paragraph(s)/table analog of an inline
   *  control. Properties go in Document.sdts; each wrapped block carries the id on
   *  Block.sdtPath (outer→inner, so nesting composes). Used by report scrape /
   *  write-back to tag editable regions. */
  contentControlRange(props: SdtProps, build: (s: this) => void): this {
    const sdts = (this.ctx.doc.sdts ??= {});
    const sid = this.ctx.ids.next();
    sdts[sid] = props;
    const startIdx = this.blocks.length;
    build(this);
    if (this.blocks.length === startIdx) {
      this.ctx.warn(`sdt-range-empty:${sid}`, "contentControlRange() wrapped no blocks — no control was recorded.");
      delete sdts[sid];
      return this;
    }
    // Prepend as the OUTER control: blocks an inner range already stamped keep
    // their inner ids, with this id ahead of them (outer→inner ancestry).
    for (let i = startIdx; i < this.blocks.length; i++) {
      const b = this.blocks[i]!;
      b.sdtPath = b.sdtPath ? [sid, ...b.sdtPath] : [sid];
    }
    return this;
  }

  /** A trailing pageBreak() with no following block still needs to exist —
   *  materialize it as an empty page-broken paragraph (called by build()). A
   *  pending column break alone is dropped (Word ignores a trailing one). */
  protected flushPendingPageBreak(): void {
    if (!this.pendingPageBreak) return;
    this.blocks.push(this.ctx.paragraph([], { pageBreakBefore: true }));
    this.pendingPageBreak = false;
    this.pendingColumnBreak = false;
  }
}
