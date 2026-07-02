// Child documents — lightweight sibling documents that SHARE a parent editor's
// live style context (stylesheet, list defs, content-control + field maps, page
// section) and render an arbitrary content slice with the REAL layout engine and
// canvas painter, instead of an HTML approximation.
//
// Two render modes share one core (build a child Document from the parent's style
// context + a content slice):
//   - render()      — paint-only static preview (style swatches, field results,
//                     read-only content). Reuses createPaintLayer with page chrome
//                     off and crops the surface to the content's measured size.
//   - mountEditor() — a full createEditor over the child slice (canvas-native
//                     editing); getBlocks() hands the edited content back so the
//                     caller can splice it into the parent.
//
// Each child owns its OWN layout engine, so its cache (keyed by block id+revision
// +width) can never collide with the parent's even when block ids overlap.

import { cellCondFlags, defaultStylesheet, effectiveCellProps, forEachImage, resolveStyle, resolveTableStyle, textOfRuns } from "@cw/shared";
import type {
  Block,
  CharStyle,
  Document,
  ListDefinition,
  ParaStyle,
  Paragraph,
  Run,
  SectionProps,
  Stylesheet,
  TableBlock,
  TableCell,
  TableRow,
  TableStyle,
} from "@cw/shared";
import { insertImage as insertImageCmd } from "../editor/commands";
import type { ResolvedBehavior, ResolvedTheme } from "../config";

/** A block-level rich fragment (paragraph runs + style), mirroring the editor's
 *  clipboard fragment — distinct from the review `Fragment` (which is `Run[]`). */
export interface BlockFragment {
  blocks: { runs: Run[]; style: ParaStyle }[];
  inline?: boolean;
}
import { createLayoutEngine } from "../layout/engine";
import type { LayoutTree, PlacedBlock } from "../layout/layoutTree";
import { createPaintLayer, type PaintScheduler } from "../paint/renderer";
import { parseOoxmlFragment } from "../import/docx/fragment";
// Type-only (erased at runtime) — keeps childDocument free of a runtime cycle with
// index.ts, which imports createChildDocument as a value.
import type { Editor, EditorOptions } from "../index";

/** A proxy over the parent document's media (its "relationships"): resolve an
 *  existing content-addressed image to a live URL, and register new image bytes
 *  back into the PARENT's store so a child-added image is shared, persists, and
 *  exports. Backed by the session media store. */
export interface ChildMedia {
  /** A live URL for a content-addressed image, or undefined if not in the store. */
  resolve(mediaId: string): string | undefined;
  /** Register new image bytes into the parent's store; resolves to the mediaId. */
  register(bytes: Uint8Array, mime: string): Promise<string>;
}

/** A live view of the parent document's style context. Pulled fresh on each
 *  render so children reflect the CURRENT parent styles (e.g. a redefined style). */
export interface StyleContext {
  stylesheet?: Stylesheet | undefined;
  lists?: Document["lists"];
  sdts?: Document["sdts"];
  fields?: Document["fields"];
  section: SectionProps;
  defaultStyleId?: string | undefined;
  /** Proxy over the parent's media/relationships — lets a child render the parent's
   *  images and register new ones back. Absent ⇒ images resolve by their live src only. */
  media?: ChildMedia | undefined;
}

/** What to render into a child surface. */
export type ChildContent =
  | { kind: "blocks"; blocks: Block[] }
  | { kind: "fragment"; fragment: BlockFragment }
  | { kind: "runs"; runs: Run[]; paraStyle?: ParaStyle }
  | { kind: "ooxml"; ooxml: string } // parsed via parseOoxmlFragment (media-free)
  | { kind: "styleSample"; styleId: string; sampleText?: string }
  // One sample paragraph per list level, rendered with the REAL engine against
  // the draft definition so the actual markers/indents show (List style editor).
  | { kind: "listSample"; listDef: ListDefinition; levels?: number }
  // A small synthetic table with cells baked from the draft table style, so the
  // header/banding/borders show through the real engine (Table style editor).
  | { kind: "tableSample"; tableStyle: TableStyle; rows?: number; cols?: number };

export interface ChildRenderOptions {
  /** Layout/wrap width in CSS px. Default: target.clientWidth. */
  widthPx?: number;
  /** Child section margins (all sides). Default 0. */
  padding?: number;
  /** Cap the displayed height (content beyond it is clipped). */
  maxHeightPx?: number;
  /** Fixed presentational scale (1 = 100%). Ignored when `fit` is set. */
  zoom?: number;
  /** Auto-scale the rendered content to fit `widthPx` × `maxHeightPx` (used for
   *  compact style swatches). Lays content out unwrapped, then scales down. */
  fit?: boolean;
}

/** A live interactive editor over a child slice (canvas-native editing). */
export interface ChildEditorHandle {
  /** The edited content as model blocks (hand back to the parent to splice). */
  getBlocks(): Block[];
  getDocument(): Document;
  /** Insert an image at the caret. The bytes are registered in the PARENT's media
   *  store (content-addressed), so it renders here, survives serialize/export, and
   *  resolves after commit-back. Dimensions default to the image's natural size.
   *  No-op without a media context. */
  insertImage(bytes: Uint8Array, mime: string, opts?: { widthPx?: number; heightPx?: number }): Promise<void>;
  destroy(): void;
}

export interface ChildDocument {
  /** Paint a static preview of `content` into `target` (a canvas surface is
   *  created/reused inside it). Safe to call repeatedly to re-render. */
  render(target: HTMLElement, content: ChildContent, opts?: ChildRenderOptions): void;
  /** Mount an interactive editor over `content` inside `target`. */
  mountEditor(target: HTMLElement, content: ChildContent, opts?: ChildRenderOptions): ChildEditorHandle;
  /** Tear down every surface/editor this child created. */
  destroy(): void;
}

/** Dependencies injected by the editor (keeps this module decoupled + cycle-free). */
export interface ChildDeps {
  /** Snapshot the parent's current style context. */
  getStyleContext: () => StyleContext;
  /** Construct a child editor (the real createEditor, injected by index.ts). */
  makeEditor: (container: HTMLElement, doc: Document, opts: EditorOptions) => Editor;
  /** Parent's resolved theme — so child previews/editors share its look. */
  theme?: ResolvedTheme;
  /** Parent's resolved behavior (zoom clamp). */
  behavior?: ResolvedBehavior;
}

// A tall page for the measuring pass — keeps all content on page 0 so its size
// can be read off the layout (cheap; this pass is never painted).
const MEASURE_H = 1_000_000;

const fallbackChar = (): CharStyle => ({
  fontFamily: "Georgia, serif",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#202124",
});

const fallbackPara = (): ParaStyle => ({
  align: "left",
  lineHeight: 1.5,
  spaceBeforePx: 0,
  spaceAfterPx: 0,
  indentFirstLinePx: 0,
  indentLeftPx: 0,
});

export function createChildDocument(deps: ChildDeps): ChildDocument {
  // Shared across this child's targets; renders are sequential (sync) so a single
  // engine is safe — each render() resets it, lays out, and hands the produced
  // tree (independent data) to the paint layer.
  const engine = createLayoutEngine();
  const surfaces = new Map<HTMLElement, { host: HTMLElement; paint: PaintScheduler }>();
  const editors = new Set<{ destroy: () => void }>();

  const para = (id: string, runs: Run[], style: ParaStyle): Paragraph => ({
    kind: "paragraph",
    id,
    revision: 0,
    runs: runs.length ? runs : [{ text: "", style: fallbackChar() }],
    style,
  });

  const styleSampleBlock = (styleId: string, sampleText: string | undefined, ctx: StyleContext): Paragraph => {
    const sheet = ctx.stylesheet ?? defaultStylesheet();
    const def = resolveStyle(sheet, ctx.defaultStyleId ?? sheet.defaultStyleId);
    const tgt = resolveStyle(sheet, styleId);
    const char: CharStyle = { ...fallbackChar(), ...def.char, ...tgt.char };
    // A swatch shows TYPOGRAPHY, not layout: keep the style's font/weight/color but
    // force flush-left with no indents/spacing/markers so it measures to the text's
    // true width (a centered style would otherwise push content across the wide
    // measuring page and scale the swatch to nothing).
    const style: ParaStyle = {
      ...fallbackPara(),
      ...def.para,
      ...tgt.para,
      align: "left",
      indentLeftPx: 0,
      indentFirstLinePx: 0,
      indentRightPx: 0,
      spaceBeforePx: 0,
      spaceAfterPx: 0,
      pageBreakBefore: false,
    };
    delete style.list;
    delete style.outlineLevel;
    delete style.tocEntry;
    delete style.sectionBreak;
    return para("cw-child-sample", [{ text: sampleText ?? "AaBbCcDd", style: char }], style);
  };

  const tableSampleBlock = (style: TableStyle, rowN: number, colN: number, ctx: StyleContext): TableBlock => {
    const resolved = resolveTableStyle({ [style.id]: style }, style.id);
    const base = resolveStyle(ctx.stylesheet ?? defaultStylesheet(), ctx.defaultStyleId ?? (ctx.stylesheet ?? defaultStylesheet()).defaultStyleId);
    const overrides = { firstRow: true, bandRows: true, rowBandSize: style.rowBandSize ?? 1, colBandSize: style.colBandSize ?? 1 };
    const rows: TableRow[] = [];
    for (let r = 0; r < rowN; r++) {
      const cells: TableCell[] = [];
      for (let c = 0; c < colN; c++) {
        const flags = cellCondFlags(r, c, rowN, colN, overrides);
        const props = effectiveCellProps(resolved, flags);
        const char: CharStyle = { ...fallbackChar(), ...base.char, ...props.char };
        const text = r === 0 ? `Head ${c + 1}` : `R${r}C${c + 1}`;
        const cell: TableCell = {
          id: `cw-tbl-${r}-${c}`,
          blocks: [para(`cw-tbl-p-${r}-${c}`, [{ text, style: char }], { ...fallbackPara(), spaceAfterPx: 0, spaceBeforePx: 0 })],
        };
        if (props.shading !== undefined) cell.shading = props.shading;
        if (props.borders && Object.keys(props.borders).length > 0) cell.borders = props.borders;
        if (props.margin) cell.margin = props.margin;
        cells.push(cell);
      }
      rows.push({ cells });
    }
    return { kind: "table", id: "cw-child-tbl", revision: 0, rows, colFractions: Array.from({ length: colN }, () => 1 / colN) };
  };

  const contentToBlocks = (content: ChildContent, ctx: StyleContext): Block[] => {
    switch (content.kind) {
      case "blocks":
        return content.blocks;
      case "ooxml":
        return parseOoxmlFragment(content.ooxml);
      case "fragment":
        return content.fragment.blocks.map((b, i) => para(`cw-child-${i}`, b.runs, b.style));
      case "runs":
        return [para("cw-child-0", content.runs, content.paraStyle ?? fallbackPara())];
      case "styleSample":
        return [styleSampleBlock(content.styleId, content.sampleText, ctx)];
      case "tableSample":
        return [tableSampleBlock(content.tableStyle, content.rows ?? 4, content.cols ?? 3, ctx)];
      case "listSample": {
        const def = content.listDef;
        const count = Math.max(1, Math.min(content.levels ?? 3, def.levels.length));
        const base = resolveStyle(ctx.stylesheet ?? defaultStylesheet(), ctx.defaultStyleId ?? (ctx.stylesheet ?? defaultStylesheet()).defaultStyleId);
        const char: CharStyle = { ...fallbackChar(), ...base.char };
        return Array.from({ length: count }, (_, i) =>
          para(`cw-child-list-${i}`, [{ text: `List item at level ${i + 1}`, style: char }], {
            ...fallbackPara(),
            ...base.para,
            spaceBeforePx: 0,
            spaceAfterPx: 2,
            list: { listId: def.id, level: i },
          }),
        );
      }
    }
  };

  /** The style context augmented with a draft list definition so a listSample
   *  renders against the unsaved definition (it isn't in the parent's lists yet). */
  const ctxFor = (content: ChildContent, ctx: StyleContext): StyleContext =>
    content.kind === "listSample"
      ? { ...ctx, lists: { ...(ctx.lists ?? {}), [content.listDef.id]: content.listDef } }
      : ctx;

  const buildDoc = (ctx: StyleContext, blocks: Block[], widthPx: number, padding: number, heightPx: number): Document => {
    const doc: Document = {
      section: {
        pageWidthPx: Math.max(1, widthPx),
        pageHeightPx: Math.max(1, heightPx),
        marginPx: { top: padding, right: padding, bottom: padding, left: padding },
      },
      blocks,
    };
    // Share the parent's style/numbering/control maps by reference (read-only).
    if (ctx.stylesheet) doc.stylesheet = ctx.stylesheet;
    if (ctx.lists) doc.lists = ctx.lists;
    if (ctx.sdts) doc.sdts = ctx.sdts;
    if (ctx.fields) doc.fields = ctx.fields;
    // Rehydrate content-addressed images from the parent's media store so a child
    // renders the parent's pictures even when blocks carry only a mediaId (a blank
    // src from a snapshot / serialized slice). Live srcs are left untouched.
    const m = ctx.media;
    if (m) forEachImage(doc, (b) => {
      if (b.mediaId && (!b.src || b.src === "")) {
        const url = m.resolve(b.mediaId);
        if (url) b.src = url;
      }
    });
    return doc;
  };

  const ensureSurface = (target: HTMLElement): { host: HTMLElement; paint: PaintScheduler } => {
    let s = surfaces.get(target);
    if (s) return s;
    const host = document.createElement("div");
    host.style.cssText = "overflow:hidden;position:relative;";
    target.appendChild(host);
    s = {
      host,
      paint: createPaintLayer(host, {
        chrome: false,
        ...(deps.theme ? { theme: deps.theme } : {}),
        ...(deps.behavior ? { zoomMin: deps.behavior.zoomMin, zoomMax: deps.behavior.zoomMax } : {}),
      }),
    };
    surfaces.set(target, s);
    return s;
  };

  const render = (target: HTMLElement, content: ChildContent, opts: ChildRenderOptions = {}): void => {
    const ctx = ctxFor(content, deps.getStyleContext());
    const blocks = contentToBlocks(content, ctx);
    const padding = opts.padding ?? 0;
    const fit = opts.fit === true;
    const surf = ensureSurface(target);
    // Measure the target's available content width by filling it — clientWidth on
    // the target would include its padding, which overflows and shows an x-scroll.
    surf.host.style.width = "100%";
    const avail = Math.max(1, surf.host.clientWidth || target.clientWidth || 1);
    const boxW = opts.widthPx ?? avail;
    // Fit mode lays out wide so a short sample doesn't wrap, then scales down.
    const layoutW = fit ? Math.max(boxW || 0, 2000) : Math.max(1, boxW);

    // Pass 1: measure content on a tall page (never painted).
    engine.reset();
    const measured = contentBounds(engine.layout(buildDoc(ctx, blocks, layoutW, padding, MEASURE_H)));
    const contentH = Math.max(1, measured.height + padding); // + bottom padding
    const contentW = Math.max(1, fit ? measured.width + padding : layoutW);

    // Pass 2: exact page so the canvas is tight.
    engine.reset();
    const tree = engine.layout(buildDoc(ctx, blocks, contentW, padding, contentH));

    let zoom = opts.zoom ?? 1;
    const maxH = opts.maxHeightPx ?? Infinity;
    if (fit) zoom = Math.min(1, (boxW || contentW) / contentW, maxH / contentH);
    const dispH = (fit ? contentH : Math.min(contentH, maxH)) * zoom;

    // Fit (swatches) crops tight; otherwise fill the width so the page matches the
    // host exactly — no horizontal overflow, content uses the full preview width.
    surf.host.style.width = fit ? `${contentW * zoom}px` : "100%";
    surf.host.style.height = `${dispH}px`;
    surf.paint.setTree(tree);
    surf.paint.setZoom(zoom);
  };

  const mountEditor = (target: HTMLElement, content: ChildContent, opts: ChildRenderOptions = {}): ChildEditorHandle => {
    const ctx = ctxFor(content, deps.getStyleContext());
    const blocks = contentToBlocks(content, ctx);
    const padding = opts.padding ?? 8;

    // Fill the target's width (no inner scrollbar: the outer container scrolls).
    const host = document.createElement("div");
    host.style.cssText = "width:100%;position:relative;";
    target.appendChild(host);
    const availW = Math.max(1, host.clientWidth || target.clientWidth || 400);
    const availH = target.clientHeight || 0;
    const widthPx = Math.max(1, opts.widthPx ?? availW);

    // Measure content, then build a page that fills the available preview box (so a
    // short control still gets a comfortably large editing area) yet grows to hold
    // taller content; the outer container provides the single scrollbar. Reuse the
    // shared measuring engine (reset like render() does) instead of allocating a
    // throwaway engine + caches per mount.
    engine.reset();
    const measured = contentBounds(engine.layout(buildDoc(ctx, blocks, widthPx, padding, MEASURE_H)));
    const contentH = Math.max(measured.height + padding, opts.maxHeightPx ?? availH, 200);

    const ed = deps.makeEditor(host, buildDoc(ctx, blocks, widthPx, padding, contentH), {
      paintOptions: { chrome: false },
      ...(deps.theme ? { theme: deps.theme } : {}),
      ...(deps.behavior ? { behavior: deps.behavior } : {}),
    });
    const media = ctx.media;
    // Seed a caret at the end of the content so toolbar-less inserts (e.g. images)
    // have a target even before the user clicks into the surface.
    const lastPara = [...ed.getDocument().blocks].reverse().find((b): b is Paragraph => b.kind === "paragraph");
    if (lastPara) {
      const pos = { blockId: lastPara.id, offset: textOfRuns(lastPara.runs).length };
      ed.setSelection({ anchor: pos, focus: pos });
    }

    const handle: ChildEditorHandle = {
      getBlocks: () => ed.getDocument().blocks,
      getDocument: () => ed.getDocument(),
      insertImage: async (bytes, mime, imgOpts): Promise<void> => {
        if (!media) return; // no parent relationships to register into
        const mediaId = await media.register(bytes, mime);
        const src = media.resolve(mediaId);
        if (!src) return;
        let w = imgOpts?.widthPx;
        let h = imgOpts?.heightPx;
        if (!w || !h) {
          try {
            const bmp = await createImageBitmap(new Blob([new Uint8Array(bytes)], { type: mime }));
            w = w ?? bmp.width;
            h = h ?? bmp.height;
            bmp.close();
          } catch {
            w = w ?? 240;
            h = h ?? 160;
          }
        }
        ed.dispatch(insertImageCmd(src, w, h, mediaId));
      },
      destroy: () => {
        ed.destroy();
        host.remove();
        editors.delete(handle);
      },
    };
    editors.add(handle);
    return handle;
  };

  return {
    render,
    mountEditor,
    destroy: () => {
      for (const s of surfaces.values()) {
        s.paint.destroy();
        s.host.remove();
      }
      surfaces.clear();
      for (const e of [...editors]) e.destroy();
      editors.clear();
    },
  };
}

// ---------------------------------------------------------------------------
// Content extent measuring (page 0 of a measure layout). Mirrors the editor's
// placed-block vertical-bounds helper; adds a right extent for fit-to-box swatches.

function contentBounds(tree: LayoutTree): { width: number; height: number } {
  const page = tree.pages[0];
  if (!page) return { width: 0, height: 0 };
  let maxBottom = 0;
  let maxRight = 0;
  for (const pb of page.blocks) {
    const b = blockBottom(pb);
    if (b > maxBottom) maxBottom = b;
    const r = blockRight(pb);
    if (r > maxRight) maxRight = r;
  }
  return { width: maxRight, height: maxBottom };
}

function blockBottom(pb: PlacedBlock): number {
  if (pb.table) return pb.table.y + pb.table.height;
  if (pb.image) return pb.y + pb.image.height;
  let bot = pb.y;
  for (const l of pb.lines) {
    const y = pb.y + l.y + l.height;
    if (y > bot) bot = y;
  }
  return bot;
}

function blockRight(pb: PlacedBlock): number {
  if (pb.table) return pb.table.x + pb.table.width;
  if (pb.image) return pb.x + pb.image.width;
  let r = pb.x;
  for (const l of pb.lines) {
    for (const f of l.fragments) {
      const x = pb.x + f.x + f.width;
      if (x > r) r = x;
    }
  }
  return r;
}
