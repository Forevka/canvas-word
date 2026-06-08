# canvas-word — .docx Import Design

A TypeScript docx importer running in a Web Worker. No Rust/WASM: docx import is
data-dense and compute-light — the cost of marshaling a large object graph across the
WASM boundary would eat any parsing win, and the hard part (OOXML semantics → our model)
iterates fastest in TS next to `src/model/document.ts`. The worker makes import
perceptually free regardless of duration; the real cost of opening a 70-page document is
our own cold layout pass, not XML parsing.

> **Note:** `DOMParser` is *not* available in Web Workers (it's a window-only DOM API).
> We use a pure-JS XML parser instead — which also means the entire importer is DOM-free
> and runs under vitest in Node with zero browser setup.

## Dependencies

- **`fflate`** — zip inflation (tiny, zero-dep, fast).
- **`txml`** — XML → ordered node tree (`{tagName, attributes, children}`), ~5KB, very
  fast. Order preservation is non-negotiable for `document.xml` (run order *is* the
  text). Avoid `fast-xml-parser`'s default mode — it groups children by tag name and
  destroys ordering; its `preserveOrder` mode works but produces a clumsier shape.

## Module structure

```
src/import/docx/
  importDocx.ts       ← main-thread API (the only file the app imports)
  worker.ts           ← worker entry: message protocol only, no logic
  pipeline.ts         ← orchestrates the stages below (pure, no worker refs)
  zip.ts              ← fflate wrapper: bytes → Map<partName, Uint8Array>
  xml.ts              ← txml helpers: parse, el(), attr(), children(), w: namespace
  contentTypes.ts     ← [Content_Types].xml → main part lookup
  relationships.ts    ← *.rels → Map<rId, {type, target}>
  styles.ts           ← styles.xml → StyleResolver (cascade, memoized)
  theme.ts            ← theme1.xml → {majorFont, minorFont, colorScheme}
  units.ts            ← twips/half-pt/EMU/ST_LineRule → px conversions
  documentParser.ts   ← document.xml body walk → IR (intermediate rep)
  mapToModel.ts       ← IR + resolved styles → canvas-word Document
  media.ts            ← image parts → Blob URLs + intrinsic sizes
  types.ts            ← IR types, ImportWarning, ImportResult
```

Everything except `importDocx.ts`/`worker.ts` is pure functions over
`Uint8Array`/parsed XML — fully testable headless, same philosophy as the command layer.

## Main-thread API

```ts
// importDocx.ts
export interface ImportResult {
  doc: Document;              // our model — pure data, structured-clone-safe by design
  warnings: ImportWarning[];  // every lossy decision, surfaced not swallowed
  mediaUrls: string[];        // blob: URLs to revoke when the doc is discarded
}

export async function importDocx(
  file: File | ArrayBuffer,
  opts?: { onProgress?: (phase: ImportPhase, pct: number) => void },
): Promise<ImportResult>;
```

- Worker created Vite-natively:
  `new Worker(new URL("./worker.ts", import.meta.url), { type: "module" })` — no build
  config needed.
- The file's `ArrayBuffer` is **transferred** (zero-copy), not cloned.
- Lazy singleton worker, terminated after ~30s idle. Concurrent imports queue (rare; not
  worth a pool).
- `URL.createObjectURL` *is* available in workers, and the URLs are valid on the main
  thread (same origin) — images become blob URLs worker-side and arrive in
  `ImageBlock.src` as plain strings. Caller revokes via `mediaUrls` on document close.

### Worker protocol

```ts
type ToWorker   = { id: number; buf: ArrayBuffer };                        // transferred
type FromWorker =
  | { id: number; type: "progress"; phase: ImportPhase; pct: number }
  | { id: number; type: "done"; result: ImportResult }                     // structured clone
  | { id: number; type: "error"; code: ImportErrorCode; message: string };

type ImportErrorCode = "NOT_ZIP" | "ENCRYPTED" | "NO_DOCUMENT_PART" | "MALFORMED_XML";
```

`ENCRYPTED` detection: password-protected docx is an OLE compound file, not a zip —
sniff the `D0 CF 11 E0` magic and report "password-protected files not supported"
instead of "invalid zip".

## Pipeline stages

```
buf ─► unzip ─► contentTypes ─► rels ─► styles+theme ─► parse document.xml ─► IR
                                                              │
        media parts ─► Blob URLs ◄────── drawing rIds ────────┘
                                                              ▼
                                              mapToModel ─► Document + warnings
```

1. **Unzip** (`fflate.unzipSync` — sync is correct, we're in a worker). Skip inflating
   `word/media/*` until referenced — wrap entries lazily.
2. **Locate parts**: `[Content_Types].xml` → main document part (don't hardcode
   `word/document.xml`); `word/_rels/document.xml.rels` → styles, numbering, theme,
   headers/footers, images.
3. **StyleResolver** (the heart — this is where docx fidelity lives or dies):

   ```ts
   resolveRun(pStyleId, rStyleId, directRpr): CharStyle
   resolvePara(pStyleId, directPpr): ParaStyle
   ```

   Resolution order per spec, simplified to what our model can express:
   `docDefaults` → paragraph-style `basedOn` chain (root-first) → character-style chain
   → direct formatting. Toggle properties (bold/italic) use **XOR semantics** across
   style layers — implement that correctly; it's the #1 source of "everything is bold"
   bugs. Theme indirections resolved here:
   `w:rFonts w:asciiTheme="minorHAnsi"` → theme minor latin font; `w:themeColor` →
   color-scheme hex. Memoize by `(pStyleId, rStyleId)` — 70 pages reuse a handful of
   styles.
4. **Body walk** (`documentParser.ts`) — recursive over body children:
   - `w:p` → paragraph IR: `pPr` (style ref, `jc`, `spacing`, `ind`, numbering ref) +
     inline walk: `w:r` (rPr + `w:t`/`w:tab`/`w:br`/`w:drawing`), `w:hyperlink`
     (unwrap, apply `Hyperlink` char style), `w:fldSimple`/field codes (take cached
     result text), `w:sdt` inline (unwrap `sdtContent` transparently)
   - `w:tbl` → rows → `w:tc` → nested block walk (cells hold paragraphs — matches
     `TableCell.blocks: Paragraph[]`); record `gridSpan`/`vMerge` in IR even though the
     model can't hold them yet
   - `w:sdt` block-level → unwrap to its content (content controls become invisible —
     warning emitted)
   - body-level `w:sectPr` → page size/margins/header-footer refs
5. **Media**: `w:drawing` → `a:blip r:embed` rId → media part → `Blob` → object URL.
   Size from `wp:extent` (EMUs); if absent, `createImageBitmap` (worker-available) for
   intrinsic size.
6. **mapToModel**: IR → our `Document`. Fresh ids with an import-distinct prefix
   (`i0`, `i1`… — `freshBlockId` in commands.ts uses `n…`, `sampleDoc` uses `b…`; no
   collisions). Run normalization (merge adjacent equal-styled runs) — same invariant
   the ops maintain, applied once at the end. `revision: 0` everywhere.

## Unit conversions (`units.ts`, all pure, all unit-tested)

| OOXML unit | Used for | → px (96dpi) |
|---|---|---|
| twips (1/20 pt) | margins, indents, spacing | `tw / 15` |
| half-points | font size (`w:sz`) | `hp * 2 / 3` |
| EMU (914400/inch) | image extents | `emu / 9525` |
| `w:line` @ `lineRule="auto"` | line height | multiplier = `line / 240` |
| `w:line` @ `exact`/`atLeast` | line height | twips → px, **approximated** to multiplier vs font size + warning (`ParaStyle.lineHeight` is multiplier-only) |

## Lossy mappings — explicit decisions, all emitting `ImportWarning`

The model can't hold everything in a rich docx. Each gap gets a deliberate policy rather
than a silent drop:

| docx feature | Model gap | Policy (phase 1) |
|---|---|---|
| Headers/footers (rich) | `SectionProps.header?: Block[]` — full block stories with `{page}`/`{pages}` tokens in run text | Default variant mapped as a block story (own rels for images); `PAGE`/`NUMPAGES` fields → `{page}`/`{pages}` tokens. First/even variants ignored with warning |
| Images in table cells | — | Faithful: `TableCell.blocks: Block[]` holds images and nested tables |
| Floating images (`wp:anchor`) | `ImageBlock.wrap` | Square/tight wrap → `wrap: "square"` float with `positionH` alignment; overlap/top-and-bottom wraps demote to block flow with warning |
| Linked (external) images | — | http(s) targets pass through by URL; `file:` targets skipped with warning |
| Numbering/lists | — | Faithful: `numbering.xml` → `Document.lists` (`ListDefinition`) + `ParaStyle.list`; markers render paint-only. Level indent is de-duplicated against the paragraph indent (engine adds them); bullet glyphs normalized from Symbol/Wingdings code points |
| Hyperlinks | `CharStyle.link` | Faithful: external `r:id` → URL (via the part's rels), `w:anchor` → `#bookmark`; warns if a target rel is missing |
| Highlight | `CharStyle.highlightColor` | Faithful: 16 named colors → hex |
| Super/subscript | `CharStyle.verticalAlign` | Faithful: `w:vertAlign` → `"super"`/`"sub"` |
| Footnotes | `Document.footnotes` + `CharStyle.footnoteRef` | Faithful: `footnotes.xml` bodies (own rels) keyed `fn<id>`; ref runs numbered sequentially in document order; separator/continuation pseudo-notes skipped; tables-in-notes dropped (warning) |
| `w:keepLines` | `ParaStyle.keepLinesTogether` | Faithful |
| Newspaper columns | `SectionProps.columns` | Faithful: `w:cols` (count > 1) with px gap (0.5in default) |
| Page-number restart | `SectionProps.pageNumberStart` | Faithful: `w:pgNumType/@w:start` |
| `w:br` (soft line break) | no soft breaks | Split into a new paragraph with `spaceBefore/After: 0` |
| `w:tab` | no tab stops | Replace with fixed spaces (warning) |
| `gridSpan` | — | Faithful: maps to `TableCell.colSpan` (and `w:tblGrid` → `colFractions`) |
| `vMerge` | no row spans | Continuation cells stay as separate (empty) cells; warning |
| Explicit page breaks (`w:br type="page"`, `w:pageBreakBefore`) | — | Faithful: map to `ParaStyle.pageBreakBefore` (inline breaks split the paragraph; the follower carries the break) |
| Multiple sections | single `SectionProps` | Page *geometry*: body-level `sectPr` wins (last section), with warning. Page *boundaries*: non-continuous section breaks set `pageBreakBefore` on the following block |

This table doubles as the **model-evolution backlog** — when the model later gains lists
or rich headers, the importer seam already collects the data (the IR keeps it; only
`mapToModel` discards).

## Testing

- All stages are DOM-free → plain **vitest in Node**. Fixtures: a handful of small real
  `.docx` files in `src/import/docx/__fixtures__/` (Word-generated, covering styles
  cascade, tables, images, sdt, headers).
- Unit level: `units.test.ts` (pure math), `styles.test.ts` (cascade + toggle-XOR
  against minimal `styles.xml` snippets), `relationships.test.ts`.
- Golden level: fixture → full pipeline → snapshot of the resulting `Document` JSON +
  warnings list.
- Perf probe: extend the existing `?stress=` pattern — a `?docx=` URL param in `main.ts`
  that imports a file and logs phase timings, to confirm on a real 70-pager that
  parse ≪ layout.

## Build order

1. ✅ **Skeleton + happy path** — zip → document.xml → paragraphs/runs, worker wiring,
   `?docx=` dev hook + 📂 toolbar button. Tables, sectPr, sdt/hyperlink/field
   unwrapping, hidden text (`w:vanish`) landed here too. 8.9 MB real report: ~160ms.
2. ✅ **StyleResolver + theme** — docDefaults, `basedOn` chains (cycle-guarded),
   toggle-XOR (§17.7.3), `w:default` paragraph style, theme fonts (`asciiTheme`) and
   colors (`themeColor`), memoized per (pStyle, rStyle). Parts located via real
   relationship parsing (`relationships.ts`), conventional names as fallback.
3. ✅ **Images + headers/footers** — DrawingML (`wp:inline`/`wp:anchor` via
   `mc:AlternateContent` unwrapping — how Word actually ships drawings) and legacy VML
   (`w:pict`); media parts → Blob URLs created lazily (media zip entries aren't even
   inflated unless referenced); external (`TargetMode="External"`) rels kept — linked
   http(s) images pass through by URL. Header/footer parts → `Block[]` stories with
   their own rels; complex/simple `PAGE`/`NUMPAGES` fields → `{page}`/`{pages}` tokens
   (header/footer only — body fields keep their cached result text).
4. ✅ **Numbering, hyperlinks, highlight, super/subscript** — `numbering.xml` →
   `Document.lists` + `ParaStyle.list` (real list model, paint-only markers, indent
   de-duplication, Symbol/Wingdings bullet normalization); external/anchor hyperlinks →
   `CharStyle.link` (resolved through each part's rels); `w:highlight` →
   `highlightColor`; `w:vertAlign` → `verticalAlign`.

Still on the backlog (lower impact / need model or layout support): table borders &
shading, cell vertical-merge row spans, footnotes (`footnotes.xml` → `Document.footnotes`),
tab-stop positions, first/even header variants, east-asian/complex-script fonts, a
warnings-summary UI toast.
