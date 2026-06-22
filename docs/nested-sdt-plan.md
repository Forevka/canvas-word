# Nested content-control (w:sdt) support — implementation plan

> Status: **IMPLEMENTED** (model `sdtPath`, import push/pop stacks, export LCP nesting, editor path-awareness; 6 new tests + full suites green). No snapshot migration — the `sdtId`→`sdtPath` behavior change is documented in `CHANGELOG.md` instead (old persisted snapshots drop control membership on load; re-import from `.docx`). Companion to `docs/scrape-writeback-plan.md` — outer-section + inner-field controls now round-trip, so the AppraiSys side keeps its existing section-SDT structure instead of flattening it.
>
> Verified: `node test-fixtures/scrape/roundtrip-test.mjs` — the section SDT survives import→edit→export→re-import as a SINGLE control with inner field controls nested inside (`doc.sdts` stays 6, no fragmentation), and an inner-control edit preserves the section ancestry.

## Context / problem

Today SDT membership is **flat**: `CharStyle.sdtId?: string` (one id per run), props in `Document.sdts: Record<id, SdtProps>`, and **no block-level SDT marker**. Consequences (measured in `test-fixtures/scrape/`):
- Nested controls collapse on **import** — only one id survives per run (inline walk keeps the outer; block walk via `tagBlockSdt`'s `=== undefined` guard keeps the inner — inconsistent).
- **Export** `runsXml` groups runs by a single `=== sdtId`, so it serializes at most one level and **fragments** a multi-block/section control into several controls.
- A control wrapping a whole table / run-less block isn't representable except via the surrounding runs.

**Goal:** round-trip arbitrary nesting (inline-in-inline, block-in-block, inline-in-block) and keep block-level controls (whole paragraphs/tables) intact.

## Design — path-based, mirroring the `fieldId` precedent

`fieldId` already solves the dual inline/block marker problem in this repo (`CharStyle.fieldId` + `Block.fieldId` + `Document.fields`, with `fieldRegionXml` emitting a block-spanning marker). Nesting = add **ordered paths** on top of that pattern.

- `CharStyle.sdtPath?: string[]` — **inline** ancestry within a paragraph, outer→inner. Replaces scalar `sdtId`.
- `Block.sdtPath?: string[]` on Paragraph/ImageBlock/TableBlock — **block-region** ancestry (mirrors `Block.fieldId`); makes run-less block controls representable.
- `Document.sdts` unchanged (props keyed by id; ids are the path leaves).
- **Central invariant:** a run's full enclosing chain = `(block.sdtPath ?? []).concat(run.style.sdtPath ?? [])`. Block ids are the outer prefix; inline ids the inner suffix.

**Layout/paint untouched** — `sdtId` is grep-clean in `frontend/src/layout/**` and `frontend/src/paint/**`; it's an invisible marker (`setSdtProps` → `dirtyBlockIds:[]`). **Persistence** is a structure-agnostic JSON clone — only a one-time legacy normalization is needed.

**Decision: fully migrate** off the scalar `sdtId` (no derived shadow field). Removing it makes the TypeScript compiler flag all ~30 read sites — the safest way to catch every consumer. Expose helpers instead.

## Phases & exact touch points

### 1. Model + helpers (`shared`)
- `shared/src/model/document.ts`: `CharStyle.sdtId` (:32) → `sdtPath?: string[]`; add `sdtPath?: string[]` to `Paragraph` (:134), `ImageBlock` (:148), `TableBlock` (:261) beside each `fieldId`.
- New `shared/src/model/sdt.ts`: `innermostSdtId(style)`, `blockInnermostSdtId(block)`, `fullSdtChain(block, style)`, `sdtPathEq(a,b)` (treats `[]`/`undefined` equal), `commonPrefixLen(a,b)`, `inSdt(style,id)`, `blockInSdt(block,id)`. Export from `shared/src/index.ts`. Never store `sdtPath: []` — use `undefined`.

### 2. ops (`shared/src/model/ops.ts`)
- `styleEq` (:129): `a.sdtId === b.sdtId` → `sdtPathEq(a.sdtPath, b.sdtPath)`. **Only correctness-critical op change** — keeps nested boundaries from merging runs at any level. Run-surgery (`splitRunsAt`, `sliceRuns`, `insertTextInRuns`, `applyStylePatchToRuns` :194, etc.) copies `r.style` so the array rides along; **treat `sdtPath` as immutable** (always replace, never mutate in place — the shallow `{...style}` aliases the array).
- `setSdtProps` (:870): no change. No new op. transform/replay: nothing (they don't touch `CharStyle` internals).

### 3. Import (`frontend/src/import/docx`)
- `types.ts`: run inline `sdtId` (:143) → `sdtPath?: string[]`; add `sdtPath?` to `IRParagraph` (:239) and `IRTable` (:308).
- `documentParser.ts`: add **two** stacks to `ParseCtx` — `inlineSdtStack: string[]` and `blockSdtStack: string[]` (init at :107/:140, shared across body/bands/footnotes like `nextSdt`). 
  - Inline `w:sdt` (:317): mint id, register props, `inlineSdtStack.push(id)` → recurse `walkInlines` → `pop()`. In `parseRun` (~:298) snapshot `run.sdtPath = [...inlineSdtStack]` if non-empty. Delete the old per-inline overwrite loop (:330).
  - Block `w:sdt` (:173): same with `blockSdtStack`; stamp `p.sdtPath`/`t.sdtPath = [...blockSdtStack]` where blocks are emitted in `walkBlocks` (:161-171). **Delete `tagBlockSdt` (:799-811).**
  - Two stacks cleanly split inline-in-block: block ids on the block, inline ids on the run, no double-counting. Cross-part id uniqueness preserved (minter unchanged).
- `mapToModel.ts`: `mapRun` (:632) param `sdtId?` → `sdtPath?: string[]`, set `style.sdtPath`; checkbox glyph normalization (:651-659) keys on the **innermost** id (`sdtPath[len-1]` — a checkbox is always a leaf). Thread `irBlock.sdtPath` onto mapped blocks (mirror `fieldId` stamping at :354), including cell blocks in `mapTable`.

### 4. Export (`frontend/src/export/docx/documentXml.ts`) — the hard part
Replace flat grouping with **longest-common-prefix recursion**.
- Inline `runsXml` (:150): `nestSdt(runs, depth)` — at each depth, runs with `path.length <= depth` flush to the existing `fieldRunsXml`; runs sharing `path[depth]` group into a `w:sdt` and recurse at `depth+1`. Bottoms out into the current sdt→field nesting order.
- Block-level `buildDocumentXml` (:736) + `cellXml` (:352): new `emitBlocks(blocks, depth)` groups contiguous blocks by `sdtPath[depth]` LCP, wraps each group in `w:sdt`/`w:sdtContent`, recurses; the leaf span runs the **existing** toc/`fieldId`-region/`blockXml` dispatch. Block `w:sdt` can directly contain `w:p`/`w:tbl`, so it's **simpler** than `fieldRegionXml` (no synthetic-paragraph edge handling) — run-less blocks wrap intact.
- **Composition order** (valid OOXML, falls out of recursion): block-sdt ⊃ block-field-region ⊃ paragraph ⊃ inline-sdt ⊃ inline-field.
- Bookmark-bracketing guard (:273): generalize `runs.some(r => r.style.sdtId)` → `r.style.sdtPath?.length`.

### 5. Editor commands (`frontend/src/editor/commands.ts`) — operate on innermost, never destroy ancestry
- `findSdtRanges` (:836): membership = `inSdt(run.style, id) || blockInSdt(block, id)` (i.e. `id ∈ fullSdtChain`).
- `sdtAtPosition` (:860): return **innermost** id (run path leaf → block path leaf → `dominantCellSdt` fallback). Add `sdtStackAtPosition(doc,pos): string[]` (full chain) for lock-up-the-path and "select outer".
- Mutation commands (`setSdtContent` :1081, `replaceSdtContent` :1117, `replaceSdtBlockSpan` :1156, `replaceSdtCellContent` :1207): when re-tagging, preserve the **prefix up to and including the edited id** (`origPath.slice(0, origPath.indexOf(id)+1)`) instead of `{sdtId:id}`. `insertContentControl` (:1025): **append** to the existing path when wrapping inside a control. `removeContentControl` (:1261): filter the **one** id out of each path (not wipe); block edit if `lockControl` or any ancestor has `lockContent` (`chainLocked` helper). `toggleSdtCheckbox`: unchanged.
- Read-site migration in `frontend/src/index.ts` (:381 `stripSdtMarker`, :476/525/535/689/692/1244/2073/2217-2236/2627), `editorApp.ts` (:1533), `sdtPopup.ts` (:115): compiler-guided once `sdtId` is removed; semantics preserved (innermost).

### 6. Builder + fixtures + inspector
- `paragraphBuilder.ts` `emitSdt` (:209): add `beginSdt(props)`/`endSdt()` stack so nested paths are buildable; leaf convenience methods emit single-element paths.
- `template.ts` orphan-GC (:51): collect ids over `r.style.sdtPath` and `b.sdtPath`.
- `sampleDoc.ts`: add a nested fixture (section block control over 2 paragraphs with an inner inline control; plus a section-wrapping-a-table case).
- `ui/sdtInspector.ts`: add ancestry breadcrumb (via `sdtStackAtPosition`).

### 7. Migration / back-compat — NOT DONE (by decision)
No migration shim. `serialize.ts` is unchanged (`SNAPSHOT_VERSION` stays 1). The
`sdtId`→`sdtPath` change is a documented behavior change (see `CHANGELOG.md`):
`.docx` import is unaffected (it writes `sdtPath` directly); old persisted
snapshots / op logs that carried the scalar `sdtId` simply drop control membership
on load (text intact) — re-import from the source `.docx` or re-save to restore.

### 8. Tests (hand-OOXML `simpleDocx` style of `import/docx/fields.test.ts`; round-trip style of `export/docx/fieldRoundtrip.test.ts`)
- `import/docx/sdtNested.test.ts`: inline-in-inline, block-in-block, inline-in-block path assertions; checkbox-as-leaf glyph.
- `export/docx/sdtNested.test.ts`: nested `w:sdt` emission; block control wrapping `w:tbl` → `<w:sdt><w:sdtContent><w:tbl>`.
- `export/docx/sdtRoundtrip.test.ts`: 2-level nested + section-wrapping-a-table survive import→export→import (`sdtPathEq`).
- commands test: edit-inside-nested preserves outer ancestry; remove-inner keeps outer; outer lock blocks inner edit.

## Edge cases
Checkbox keys on innermost; run-less block controls wrap intact (no synthetic para); fieldId+sdt co-location → sdt outer (recursion order); bookmarks bracket outer `w:p`; placeholder is per-id; mixed-depth runs handled by `path.length <= depth` base case; path arrays immutable (aliasing); `[]`≡`undefined` in `sdtPathEq`.

## Risks
- **Read-site sprawl** (~30 sites) — mitigated by fully removing `sdtId` so the compiler flags each; a miss drops UX, not data (data lives in `sdtPath`).
- **`replaceSdt*` ancestry prefix** is the subtlest new code — a bug drops the outer control on inspector save; cover with the edit-inside-nested test.
- **Op-log back-compat** — replaying pre-change logs could re-introduce scalar `sdtId`; resolve via op-log lifetime check / `normalizeCharStyle`.
- **Block-sdt interactive removal** not op-backed in v1 (import/inspector-driven only) — document as follow-up.

## Verification
`npm test -w @cw/shared && npm test -w @forevka/wordcanvas` (new sdtNested + roundtrip + serialize tests green), then re-run `node test-fixtures/scrape/roundtrip-test.mjs` against `scrape-sample.docx` — the outer section SDT must now survive as a **single** control with the inner field controls nested inside it (no fragmentation), and an edit inside an inner control must preserve the outer section ancestry.
