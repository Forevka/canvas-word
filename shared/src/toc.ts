// Table-of-contents field instruction parsing + generation options. Pure and
// layout-free, so import, export, and the backend route all share it.
//
// The OOXML `TOC` field instruction (e.g. ` TOC \o "1-3" \h \z \u `) governs WHICH
// paragraphs become entries and some behaviors; the visual leader/right-tab live in
// the TOC paragraph styles, not the instruction (see TocOptions).

import type { CharStyle, Document, ParaStyle, Paragraph } from "./model/document";
import { DEFAULT_CHAR_STYLE } from "./model/defaults";
import { bodyParagraphs, textOfRuns } from "./model/text";
import { sliceRuns } from "./model/ops";
import { freshId } from "./ids";

/** Parsed `TOC` field switches. See ECMA-376 §17.16.5.68. */
export interface TocSwitches {
  /** \o "1-3" — include built-in heading levels in this (inclusive) range. */
  outlineRange?: { from: number; to: number };
  /** \u — also include paragraphs with a directly-applied outline level. */
  useOutlineLevels: boolean;
  /** \t "StyleName,level,…" — custom paragraph styles mapped to TOC levels. */
  customStyles?: Record<string, number>;
  /** \h — emit each entry as a hyperlink to its heading. */
  hyperlinks: boolean;
  /** \n "1-3" — omit page numbers for levels in this (inclusive) range. */
  hidePageNumberRange?: { from: number; to: number };
  /** \p "sep" — text between the entry label and the page number (default tab). */
  separator?: string;
  /** \z — hide tab/leader/page numbers in Web layout view (display-only). */
  hideInWeb: boolean;
}

const rangeOf = (s: string | undefined): { from: number; to: number } | undefined => {
  if (!s) return undefined;
  const m = s.match(/^(\d+)\s*-\s*(\d+)$/);
  if (m) return { from: Number(m[1]), to: Number(m[2]) };
  const n = s.match(/^\d+$/) ? Number(s) : NaN;
  return Number.isNaN(n) ? undefined : { from: n, to: n };
};

/** Parse a `TOC` field instruction string into its switches. */
export function parseTocInstruction(instr: string): TocSwitches {
  const sw: TocSwitches = {
    useOutlineLevels: /\\u\b/.test(instr),
    hyperlinks: /\\h\b/.test(instr),
    hideInWeb: /\\z\b/.test(instr),
  };
  const oRange = rangeOf(instr.match(/\\o\s+"([^"]+)"/)?.[1]);
  if (oRange) sw.outlineRange = oRange;
  const nRange = rangeOf(instr.match(/\\n\s+"?([0-9-]+)"?/)?.[1]);
  if (nRange) sw.hidePageNumberRange = nRange;
  const sep = instr.match(/\\p\s+"([^"]*)"/)?.[1];
  if (sep !== undefined) sw.separator = sep;
  const t = instr.match(/\\t\s+"([^"]+)"/)?.[1];
  if (t) {
    const pairs = t.split(",");
    const map: Record<string, number> = {};
    for (let i = 0; i + 1 < pairs.length; i += 2) {
      const name = pairs[i]!.trim();
      const lvl = Number(pairs[i + 1]!.trim());
      if (name && lvl >= 1) map[name] = lvl;
    }
    if (Object.keys(map).length > 0) sw.customStyles = map;
  }
  return sw;
}

/** Synthesize a `TOC` field instruction from switches (inverse of
 *  parseTocInstruction; round-trips the switches it models). Switch order follows
 *  Word's own emit so a re-parse is stable. */
export function buildTocInstruction(sw: TocSwitches): string {
  const parts: string[] = ["TOC"];
  if (sw.outlineRange) parts.push(`\\o "${sw.outlineRange.from}-${sw.outlineRange.to}"`);
  if (sw.useOutlineLevels) parts.push("\\u");
  if (sw.customStyles && Object.keys(sw.customStyles).length > 0) {
    const t = Object.entries(sw.customStyles).map(([name, lvl]) => `${name},${lvl}`).join(",");
    parts.push(`\\t "${t}"`);
  }
  if (sw.hyperlinks) parts.push("\\h");
  if (sw.hidePageNumberRange) parts.push(`\\n "${sw.hidePageNumberRange.from}-${sw.hidePageNumberRange.to}"`);
  if (sw.separator !== undefined) parts.push(`\\p "${sw.separator}"`);
  if (sw.hideInWeb) parts.push("\\z");
  return ` ${parts.join(" ")} `;
}

/** Is `level` covered by an inclusive range (undefined range = no). */
export const inRange = (r: { from: number; to: number } | undefined, level: number): boolean =>
  !!r && level >= r.from && level <= r.to;

// ---------------------------------------------------------------------------
// Heading detection (shared by the editor's TOC generation and the headless
// route) — robust to opaque style ids: tries the effective outline level, then
// the resolved style NAME ("Heading 1"), then a `namedStyle` like "heading1",
// then \t custom-style mappings.

/** The friendly name of a style id, following basedOn (for opaque numeric ids). */
function styleNameOf(doc: Document, id: string | undefined): string | undefined {
  if (!id || !doc.stylesheet) return undefined;
  const byId = new Map(doc.stylesheet.styles.map((s) => [s.id, s]));
  let cur = byId.get(id);
  const seen = new Set<string>();
  while (cur && !seen.has(cur.id)) {
    if (cur.name && /(?:^|\s)heading\s*\d/i.test(cur.name)) return cur.name;
    seen.add(cur.id);
    cur = cur.basedOn ? byId.get(cur.basedOn) : undefined;
  }
  return byId.get(id)?.name;
}

/** TOC level (1-based) for a paragraph, or null if it's not a heading. */
export function tocHeadingLevel(p: Paragraph, doc: Document, sw?: TocSwitches): number | null {
  const named = p.style.namedStyle;
  // \t custom styles (match by opaque id or by friendly name).
  if (sw?.customStyles && named) {
    const lv = sw.customStyles[named] ?? sw.customStyles[styleNameOf(doc, named) ?? ""];
    if (lv) return lv;
  }
  if (p.style.outlineLevel !== undefined) return p.style.outlineLevel + 1;
  const direct = named?.match(/^heading\s*(\d)$/i);
  if (direct) return Number(direct[1]);
  const m = styleNameOf(doc, named)?.match(/(?:^|\s)heading\s*(\d)/i);
  return m ? Number(m[1]) : null;
}

/** All TOC-eligible headings in document order (body + table cells), honoring the
 *  switches' `\o` range (when present) and `maxLevel`. */
export function detectTocHeadings(
  doc: Document,
  sw?: TocSwitches,
  maxLevel = 3,
): { block: Paragraph; level: number }[] {
  const cap = sw?.outlineRange?.to ?? maxLevel;
  const out: { block: Paragraph; level: number }[] = [];
  for (const p of bodyParagraphs(doc)) {
    const level = tocHeadingLevel(p, doc, sw);
    if (level === null || level < 1 || level > cap) continue;
    if (sw?.outlineRange && (level < sw.outlineRange.from || level > sw.outlineRange.to)) continue;
    out.push({ block: p, level });
  }
  return out;
}

// ---------------------------------------------------------------------------
// Imported-TOC unification: an imported TOC field flattens to plain paragraphs
// (label + "#anchor" hyperlink + cached "\t<number>" text). To round-trip it as a
// LIVE field on export, mark those paragraphs as tocEntry and STRIP the cached
// number text — so the number becomes layout-driven (paint decoration), exactly
// like an editor-created TOC, and export wraps it in a real TOC/PAGEREF field.

/** Mark a document's imported TOC entries as tocEntry and strip their cached page
 *  numbers. Gated by the caller (only when a TOC field exists). Returns the count.
 *  An entry is a body paragraph with a "#anchor" link resolving to a bookmarked
 *  block plus a trailing "\t<number>"; level is inferred from its left indent. */
export function markImportedTocEntries(doc: Document, indentStepPx = 20): number {
  if (!doc.bookmarks) return 0;
  const contentWidthPx = doc.section.pageWidthPx - doc.section.marginPx.left - doc.section.marginPx.right;
  let count = 0;
  for (const b of doc.blocks) {
    if (b.kind !== "paragraph" || b.style.tocEntry) continue;
    const anchor = b.runs.find((r) => r.style.link?.startsWith("#"))?.style.link?.slice(1);
    if (!anchor) continue;
    const targetId = doc.bookmarks[anchor]?.start.blockId;
    if (!targetId) continue;
    const full = textOfRuns(b.runs);
    const m = full.match(/\t\d+\s*$/); // trailing tab + page number
    if (!m) continue;
    b.runs = sliceRuns(b.runs, 0, full.length - m[0].length); // drop "\t<number>"
    if (b.runs.length === 0) b.runs = [{ text: "", style: { ...TOC_BASE_CHAR } }];
    const level = Math.min(9, Math.max(1, Math.round((b.style.indentLeftPx || 0) / indentStepPx) + 1));
    b.style.tocEntry = { targetId, level };
    // Right-aligned dot leader so the painted/exported number aligns (if none set).
    if (!b.style.tabStops?.length && contentWidthPx > 0) {
      b.style.tabStops = [{ posPx: contentWidthPx, align: "right", leader: "dot" }];
    }
    b.revision++;
    count++;
  }
  return count;
}

// ---------------------------------------------------------------------------
// Generation options — everything that used to be hardcoded in
// buildTocParagraphs, exposed so the editor AND the backend route can pass it.

export interface TocLevelStyle {
  char?: Partial<CharStyle>;
  para?: Partial<ParaStyle>;
}

/** TOC entries default to the body font/color, one notch smaller. */
const TOC_BASE_CHAR: CharStyle = { ...DEFAULT_CHAR_STYLE, fontSizePx: 13 };

/** Per-level left-indent step (px) when none is supplied via TocOptions. */
const TOC_INDENT_STEP_PX = 20;

/** Resolve the char + paragraph style for a TOC entry at `level`. Precedence:
 *  built-in defaults < the document's own TOC style (`inherit`, from its TOC field
 *  paragraph) < caller `TocOptions`. When `inherit` is given we DON'T impose the
 *  built-in level-1 bold / size bumps — we respect the document's look (matching the
 *  source Word doc); without it, the editor's default look is used. The para carries
 *  the right-aligned leader tab stop (callers add `tocEntry`). */
export function tocEntryStyle(
  opts: TocOptions,
  level: number,
  contentWidthPx: number,
  inherit?: { char?: CharStyle | undefined; para?: ParaStyle | undefined },
): { char: CharStyle; para: ParaStyle } {
  const indentStep = opts.indentStepPx ?? TOC_INDENT_STEP_PX;
  const leader = opts.leader ?? "dot";
  // Shared paragraph base for every entry, before indent + per-level overrides.
  const paraBase = { align: "left", lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 2, indentFirstLinePx: 0 } as const;
  let char: CharStyle;
  let para: ParaStyle;
  if (inherit && (inherit.char || inherit.para)) {
    char = { ...TOC_BASE_CHAR, ...inherit.char, ...opts.baseChar, ...opts.levels?.[level]?.char };
    const baseIndent = inherit.para?.indentLeftPx ?? 0;
    para = {
      ...paraBase,
      ...inherit.para,
      indentLeftPx: baseIndent + (level - 1) * indentStep,
      ...opts.levels?.[level]?.para,
    };
  } else {
    const base = { ...TOC_BASE_CHAR, ...opts.baseChar };
    char = { ...base, fontSizePx: level === 1 ? 14 : 13, bold: level === 1, ...opts.levels?.[level]?.char };
    para = {
      ...paraBase,
      indentLeftPx: (level - 1) * indentStep, ...opts.levels?.[level]?.para,
    };
  }
  if (leader !== "none" && contentWidthPx > 0) para.tabStops = [{ posPx: contentWidthPx, align: "right", leader }];
  else delete para.tabStops; // leader:"none" overrides any inherited leader tab
  return { char, para };
}

/** The TOC title paragraph's resolved styles, or null when titles are disabled. */
export function tocTitleStyle(opts: TocOptions): { text: string; char: CharStyle; para: ParaStyle } | null {
  if (opts.title === null) return null;
  const t = opts.title ?? {};
  const base = { ...TOC_BASE_CHAR, ...opts.baseChar };
  return {
    text: t.text ?? "Table of Contents",
    char: { ...base, fontSizePx: 20, bold: true, ...t.char },
    para: {
      align: "left", lineHeight: 1.4, spaceBeforePx: 8, spaceAfterPx: 12,
      indentFirstLinePx: 0, indentLeftPx: 0, namedStyle: t.namedStyle ?? "tocTitle", ...t.para,
    },
  };
}

export interface TocOptions {
  /** TOC title paragraph; null/omit to skip a title. */
  title?: { text?: string; char?: Partial<CharStyle>; para?: Partial<ParaStyle>; namedStyle?: string } | null;
  /** Base char style applied to every entry before per-level overrides. */
  baseChar?: Partial<CharStyle>;
  /** Per-level (1-based) style overrides. */
  levels?: Record<number, TocLevelStyle>;
  /** Deepest heading level to include (default 3). */
  maxLevel?: number;
  /** Per-level left indent step in px (default 20). */
  indentStepPx?: number;
  /** Dot-leader between label and page number (default "dot"). */
  leader?: "dot" | "dash" | "underscore" | "none";
  /** Show page numbers (default true). */
  includePageNumbers?: boolean;
  /** Emit entries as hyperlinks (default true). */
  hyperlink?: boolean;
}

// ---------------------------------------------------------------------------
// TOC generation — entries are REAL paragraphs tagged with `tocEntry` (so they
// lay out, paint, hit-test, and export as a live field for free). Page numbers
// are paint-only: the layout engine resolves each entry's number per relayout,
// so they never go stale. Shared by the editor (insertTocCmd) and the headless
// PDF/DOCX render path.

const TOC_LEVELS = 3;

/** Build TOC paragraphs (optional title + one entry per heading) from a document's
 *  headings. Every hardcoded look is an overridable `TocOptions` default. The page
 *  number is paint-only (layout post-pass); a right-aligned dot-leader tab stop
 *  right-aligns it. Mutates nothing. */
export function buildTocParagraphs(doc: Document, opts: TocOptions = {}): Paragraph[] {
  // Honor the document's own TOC field switches (\o level range, \t custom styles, …)
  // so a C#-emitted ` TOC \o "1-5" ` lists all five levels, not the default three.
  const sw = doc.tocInstruction ? parseTocInstruction(doc.tocInstruction) : undefined;
  const maxLevel = opts.maxLevel ?? sw?.outlineRange?.to ?? TOC_LEVELS;
  const contentWidthPx = doc.section.pageWidthPx - doc.section.marginPx.left - doc.section.marginPx.right;

  // Inherit the document's own TOC styling from the TOC field's paragraph (its pStyle,
  // e.g. "TOC 1", resolved on import → the doc's intended font/size/indent). Absent for
  // an editor-inserted TOC in a doc with no TOC field → the built-in default look.
  const anchor = doc.tocAnchorBlockId
    ? doc.blocks.find((b): b is Paragraph => b.kind === "paragraph" && b.id === doc.tocAnchorBlockId)
    : undefined;
  const inherit = anchor ? { char: anchor.runs[0]?.style, para: anchor.style } : undefined;

  const out: Paragraph[] = [];

  // Title: an explicit opts.title wins; otherwise add the default title only when we're
  // NOT filling an existing TOC field (a doc with a TOC field usually has its own title
  // heading already, so a generated one would duplicate it).
  const wantTitle = opts.title !== undefined ? opts.title !== null : anchor === undefined;
  if (wantTitle) {
    const title = tocTitleStyle(opts);
    if (title) {
      out.push({
        kind: "paragraph", id: freshId(), revision: 0,
        runs: [{ text: title.text, style: title.char }],
        style: title.para,
      });
    }
  }

  // Scan body AND table-cell paragraphs so headings inside tables are listed.
  for (const { block, level } of detectTocHeadings(doc, sw, maxLevel)) {
    const text = textOfRuns(block.runs).replace(/\v/g, " ").trim();
    if (text.length === 0) continue;
    const { char, para } = tocEntryStyle(opts, level, contentWidthPx, inherit);
    out.push({
      kind: "paragraph", id: freshId(), revision: 0,
      runs: [{ text, style: char }],
      style: { ...para, tocEntry: { targetId: block.id, level } },
    });
  }
  return out;
}

/** Fill an EMPTY TOC field: build entries from the document's headings and splice
 *  them at the captured field location (`doc.tocAnchorBlockId`). Page numbers stay
 *  paint-only — the layout engine resolves them on render. Returns a NEW doc (input
 *  not mutated).
 *
 *  If the TOC field ALREADY has entries (imported and marked as `tocEntry`), they are
 *  PRESERVED untouched — they carry the source document's own styling and their page
 *  numbers are recomputed by layout, so rebuilding from headings (with our defaults)
 *  would only throw the author's TOC away. No headings / nowhere to anchor → unchanged. */
export function generateTocIntoDoc(
  doc: Document,
  opts: TocOptions = {},
): { doc: Document; generated: number; headings: number } {
  // Already populated (e.g. a Word/C#-rendered TOC) → keep it as-is.
  if (doc.blocks.some((b) => b.kind === "paragraph" && b.style.tocEntry)) {
    return { doc, generated: 0, headings: 0 };
  }
  if (doc.tocAnchorBlockId === undefined) return { doc, generated: 0, headings: 0 };
  const at = doc.blocks.findIndex((b) => b.id === doc.tocAnchorBlockId);
  if (at < 0) return { doc, generated: 0, headings: 0 };

  const fresh = buildTocParagraphs(doc, opts);
  const headings = fresh.reduce((n, p) => n + (p.style.tocEntry ? 1 : 0), 0);
  if (headings === 0) return { doc, generated: 0, headings: 0 };

  const blocks = [...doc.blocks];
  blocks.splice(at, 1, ...fresh); // replace the empty TOC field paragraph
  return { doc: { ...doc, blocks }, generated: headings, headings };
}
