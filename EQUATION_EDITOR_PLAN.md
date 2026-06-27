# Implementation Plan — MathML Equations + Visual Equation Editor

Status: proposal / pre-implementation. Grounded in a codebase survey (model, layout,
import/export, editor UI). All `path:line` references are from that survey and should be
re-verified at implementation time (the tree moves).

### Product decisions (locked)
- **v1 = display AND inline math together** (no phasing of inline to a later release).
- **Bundled math font = STIX Two Math** (OFL).
- **Input = visual editor primary + LaTeX/AsciiMath shortcuts** as a secondary method.
- **No opaque fallback — everything must be fully editable.** Anything Word/OMML can
  express must round-trip *and* be editable in the visual editor from day one.

> **Scope consequence of "fully editable":** there is no escape hatch. The converter must
> cover the **full common OMML element set** (not an MVP subset), the typesetter must lay out
> every one of those constructs, and the editor must let you select/replace/navigate any of
> them. We make this tractable by building the editor to operate **generically on the MathML
> AST** (select any subtree, wrap, delete, retarget a slot) rather than as a fixed set of
> per-template behaviors — so "fully editable" reduces to **converter coverage + typesetter
> coverage**, with the long pole being typesetting breadth, tracked by golden tests.

---

## 1. Goal & scope

Add first-class mathematical equations to WordCanvas:

1. **Canonical internal form = MathML** (Presentation MathML). Equations are stored,
   edited, and reasoned about as a MathML AST.
2. **Visual (WYSIWYG) equation editor** — a Word-style structured editor: templates
   (fraction, radical, scripts, n-ary), placeholder slots you tab between, a symbol
   palette, and LaTeX-style shortcuts (`\frac`, `^`, `_`, `\alpha`) that expand into
   structure. Not just a text box.
3. **Full pipeline parity** — equations render on the `<canvas>`, paint into **PDF**
   export, and round-trip through **`.docx`** (which uses **OMML**, not MathML).

### Why this is big (the three things that don't exist yet)
- **No inline replaced element.** `Run = { text, style }`; offsets index into the
  concatenated run text (`shared/src/model/document.ts:150`, `model/position.ts:4`,
  `model/text.ts:183`). Nothing occupies "one character slot but paints as a custom box."
- **No math typesetting engine.** Zero hits for `oMath|omml|mathml|equation` in
  `shared/src`. Fractions, scripts, radicals, stretchy delimiters, and spacing rules
  must all be built.
- **OMML ≠ MathML.** `.docx` math is OMML (`m:oMath`/`m:oMathPara`); the `m:` namespace
  isn't even declared (`export/docx/xmlWrite.ts:41` `WML_NS` has no `xmlns:m`). A
  bidirectional MathML↔OMML converter is required for round-trip.

---

## 2. Key design decisions (with recommendations)

### D1 — Canonical form: **MathML AST** (locked)
Store a parsed Presentation-MathML tree on the model (not a raw string), so editing,
caret movement, and layout operate on structure.

- Import OMML → MathML AST. Export MathML AST → OMML.
- **Full common OMML element set is converted, typeset, and editable** (no opaque
  placeholder — see locked decisions). We still keep the original source string on the
  `EquationDef` as a *integrity check / debugging* aid and to losslessly re-emit attributes
  we don't yet model on a node, but every construct is structurally editable. (The verbatim
  string is belt-and-suspenders, not a fallback render path.)

### D2 — Display AND inline math in v1 (locked)
- **Display/block equations** (`m:oMathPara`) map onto the existing `ImageBlock` path — new
  `Block` variant, mirror `placeImage`/`imageParagraphXml`.
- **Inline equations** (`m:oMath` inside a line): a run carrying a **single sentinel code
  unit** (`U+FFFC` OBJECT REPLACEMENT CHARACTER) plus an `equationId` on `CharStyle` (mirrors
  the `fieldId` marker, `document.ts:46`), so the "offset indexes concatenated text"
  invariant is preserved and every caret / segmentation / range op keeps working unchanged.
  The math payload lives in a `Document.equations` registry keyed by id (mirrors
  `Document.fields`, `document.ts:514`).
- Both ship in v1. To de-risk, **build order is still display-first** (it isolates the math
  engine from text-offset internals), but inline lands in the same release — the sentinel-run
  work and its `text.ts` consumer audit are in-scope for v1, not deferred.

### D3 — MathML↔OMML converter: **hand-written, subset-driven** (recommended)
Microsoft's `OMML2MML.xsl`/`MML2OMML.xsl` need an XSLT engine — unavailable in a
zero-dependency Node/browser bundle. Hand-write a focused converter over the common
Presentation-MathML subset. Both directions are pure functions over the AST/XML, unit-
testable in isolation, and independent of layout/paint.

**Element coverage = the full common OMML surface** (locked "fully editable" → no subset):
`math, mrow, mi, mn, mo, mtext, mspace, mfrac` (incl. bevelled & no-bar), `msup, msub,
msubsup, msqrt, mroot, mfenced` (+ `mo` fences/stretchy), `munder, mover, munderover`
(limits/accents/bars), `mtable/mtr/mtd` (matrices/arrays/cases), `mstyle`, `mmultiscripts`
(pre/post-scripts), `mphantom`, and OMML-specific constructs that have no 1:1 MathML element
(`m:nary` n-ary operators with limits, `m:box`, `m:groupChr`, `m:eqArr`, `m:func`, `m:limLow/
limUpp`, `m:bar`, `m:borderBox`) mapped to the nearest MathML structure. Only truly exotic
tails (`maction`, semantic `annotation-xml`, MathML 3 elementary-math `mstack`) may lag, and
those are surfaced as warnings — they are not silently dropped.

### D4 — Math font: **bundle STIX Two Math + build-time metric extraction** (locked)
Real math needs a font with an OpenType **MATH** table (axis height, rule thickness,
script percents, big-operator / stretchy-delimiter glyph variants & assembly). The bundled
clone fonts are Latin metric clones with no MATH table, and the measure host bakes vertical
metrics from clone ratios while ignoring real glyph bounding boxes
(`fonts/clones.ts:63`, `export/shared/fontkitContext.ts:55`) — unusable for math.

Plan:
- Bundle **STIX Two Math** (OFL) — the chosen face (MathJax's default; broad symbol coverage).
- **Build-time extraction step**: parse the font's MATH table + glyph metrics (italic
  correction, bboxes, variant/assembly tables) into a JSON data file. Runtime consumes the
  JSON for layout constants and the font file for glyph drawing — keeps runtime zero-dep and
  sidesteps any fontkit MATH-table gaps.
- **Export embedding** reuses the existing per-doc custom-font registration/embedding path
  (see `memory/custom-fonts.md`): to embed math glyphs in PDF/DOCX the math font is
  registered like any other font. On-screen rendering uses it directly. Call this out in
  docs exactly like the `cjk.fallbackFont` gap.

### D5 — Editor surface: **floating panel with its own render + live preview** first
True in-canvas inline math editing (own caret inside the document canvas) is the hardest UI.
Start with a **floating, non-blocking equation editor panel** (clone `fieldConstructor.ts` +
`makeFloatingDialog`) that hosts the structured editing surface and a live preview rendered
with the real math engine. Later, optionally promote to in-canvas "equation edit mode."

---

## 3. Architecture & data flow

```
                ┌────────────────────────── canonical: MathML AST ──────────────────────────┐
  .docx (OMML) ─►  import: m:oMath → OMML-XML → [OMML→MathML] → AST  ─┐                       │
  paste MathML ─►  parse MathML → AST ───────────────────────────────┤                       │
  visual editor ─► structure edits on AST ───────────────────────────┤  Document.equations   │
                                                                      │  + EquationBlock /     │
                                                                      │    inline equationId   │
                                                                      └─► math layout engine ──┼─► math box
                                                                                  (typeset)    │   {w,asc,desc,
                                                                                               │    glyphs[], rules[]}
                                          ┌────────────────────────────────────────────────────┘
                                          ├─► canvas painter (paint/renderer.ts)   ─► screen
                                          ├─► PDF painter (export/pdf/paintBlock)  ─► .pdf
                                          └─► export: AST → [MathML→OMML] → m:oMath ─► .docx
```

Two painters consume one positioned math box (the codebase deliberately mirrors canvas/PDF
draw calls and shares only pure-geometry helpers — follow the `paintStyle.ts pageBorderSegments`
pattern so canvas/PDF stay in lockstep).

---

## 4. Workstreams

### WS1 — Model & ops  (`shared/src`)
- New module `model/math.ts`: the MathML AST node types + helpers (mirrors `model/sdt.ts`).
- `Document.equations?: Record<string, EquationDef>` registry (mirror `fields`,
  `document.ts:478-515`). `EquationDef` holds the AST (+ optional verbatim source for
  unsupported constructs).
- **Block math:** add `EquationBlock` to the `Block` union (`document.ts:324`); clone the
  `ImageBlock` story for ops (`insertBlock`/`setEquationProps`, mirror `setImageProps`
  `ops.ts:57,600`) and serialization (`persist/serialize.ts`; AST is plain JSON, serializes
  free — no media-blob handling needed).
- **Inline math (phase B):** `CharStyle.equationId?` marker (mirror `fieldId`,
  `document.ts:46`) + a sentinel `U+FFFC` in run text; `setEquation` op (mirror `setField`,
  `ops.ts:74,905`). Audit `text.ts` graphemes/word-break helpers to treat the sentinel as an
  atomic cluster.
- Barrel export in `shared/src/index.ts`; regenerate `frontend/types/model.d.ts`.

### WS2 — MathML↔OMML + docx round-trip  (`frontend/src/import`, `frontend/src/export`)
- `mathml/fromOmml.ts` and `mathml/toOmml.ts` — pure converters (D3).
- `mathml/parse.ts` / `mathml/serialize.ts` — MathML string ↔ AST (for paste & storage).
- **Import wiring:**
  - inline: `import/docx/documentParser.ts:306` `walkInlines` → add `case "m:oMath":`.
  - block: `documentParser.ts:166` `walkBlocks` → add `case "m:oMathPara":`.
  - IR type: `import/docx/types.ts:144` `IRInline` → add `{ kind: "math"; ... }`.
  - IR→model: `import/docx/mapToModel.ts:517` → add `case "math":`.
  - (txml matches qualified names literally — `m:*` parses with no namespace registration.)
- **Export wiring:**
  - declare the namespace: `export/docx/xmlWrite.ts:41` `WML_NS` → add
    `"xmlns:m": "http://schemas.openxmlformats.org/officeDocument/2006/math"`.
  - block: `export/docx/documentXml.ts:491` `blockXml` → add an `m:oMathPara` branch
    (template: `imageParagraphXml:646`).
  - inline: emit `m:oMath` as a sibling in `runsXml`/`paragraphXml` (`documentXml.ts:155,267`).
- Unknown OMML → keep verbatim string on `EquationDef`, re-emit byte-for-byte.

### WS3 — Math typesetting engine  (`frontend/src/layout/math/` — new)
The core subsystem. Input: MathML AST + math-font data (D4). Output: a positioned math box
`{ width, ascent, descent, glyphs:[{char,font,size,x,y}], rules:[{x,y,w,h}] }` — pure data.
- TeX/MathML box-layout rules driven by the extracted MATH constants: fraction
  numerator/denominator shifts + rule thickness; super/subscript shifts + italic
  correction; radical with drawn rule + degree; `mrow` with inter-atom spacing classes
  (ord/op/bin/rel/punct); fenced/stretchy delimiters via glyph variants/assembly;
  munder/mover limits; `mtable` row/col alignment.
- Glyph metrics read directly from fontkit (`export/shared/fontkitContext.ts:20`), bypassing
  the clone-ratio vertical metrics. Reuse `metrics.ts` (`measureTextWidth`, `charStyleToFont`)
  for plain text spans and script sizing (`SUB_SUPER_SCALE` precedent, `metrics.ts:50`).
- **Coverage = full set (D3).** Build order ramps complexity — (1) mi/mn/mo + mrow spacing
  classes, (2) mfrac + scripts + radicals, (3) fences/stretchy delimiters via variants &
  assembly, (4) munder/over limits + accents + n-ary, (5) mtable/matrices + multiscripts —
  but all five land in v1. Each layer is gated by golden render tests before the next.

### WS4 — Layout integration + paint  (`frontend/src/layout`, `frontend/src/paint`, `frontend/src/export/pdf`)
- **Block placement:** `placeEquation` mirroring `placeImage` (`layout/engine.ts:1423`),
  emitting a `PlacedBlock` with a math-box payload.
- **Inline placement (phase B):** set pretext `extraWidth = mathBoxWidth` + `break:'never'`
  in `layout/prepareCache.ts:119` (`toItems`), and feed the math box ascent/descent into the
  line-metric reduction in `breakNextLine` (`layout/engine.ts:304-339`) so the line grows to
  fit. Add a hit-test entry in `layout/geometry.ts`.
- **Painters (two mirrored paths):** add a math branch in `paint/renderer.ts:856` (canvas,
  `ctx.fillText` glyphs + `fillRect`/stroke rules) and `export/pdf/paintBlock.ts:44` (pdfkit
  `doc.text` + `doc.rect/.moveTo.lineTo`). Compute rules (fraction bars, radicals, brackets)
  as a **shared pure segment list** (pattern: `paint/paintStyle.ts:118` `pageBorderSegments`)
  so both painters stay identical.

### WS5 — Visual equation editor UI  (`frontend/src/ui/equationEditor.ts` — new)
- Clone `ui/fieldConstructor.ts` + `ui/floatingDialog.ts` (`makeFloatingDialog:32`) for a
  draggable, non-blocking panel with scoped CSS, AbortController teardown, Escape-to-close,
  modal `mousedown` propagation stop (so canvas selection doesn't steal focus).
- **Editing surface (generic AST editor):** renders the equation with the real math engine
  and maintains its own **math caret/selection over the AST** with *generic* structure
  operations — select any subtree, delete, wrap-in-template, retarget into any empty slot,
  navigate into/out of any node. Genericity is what makes the locked "fully editable"
  requirement bounded: any node the converter/typesetter supports is automatically editable,
  no per-construct UI code.
- **Templates & palette:** buttons inserting fraction / radical / script / n-ary / matrix
  templates with empty slots; a categorized **symbol palette** (Greek, operators, relations,
  arrows, big operators).
- **Keyboard (visual + LaTeX shortcuts, locked):** `^`→superscript slot, `_`→subscript slot,
  `/`→fraction, `\name`→symbol/template expansion via a LaTeX/AsciiMath token table,
  Tab/Shift-Tab between slots, arrows within structure. LaTeX typing is the power-user accel;
  the palette/templates are the discoverable primary path.
- **Live preview** of the result rendered in the document's real font (child-document trick,
  `index.ts:2173` `openFieldConstructor` precedent), so what you build is what lands.
- Apply → `editor.dispatch(insertEquation(ast))` / `editEquationCmd(id, ast)`.

### WS6 — Commands, selection, ribbon, click-to-edit  (`frontend/src/editor`, `frontend/src/index.ts`, `editorApp.ts`)
- Commands `insertEquation(ast)` / `editEquationCmd(id, ast)` in `editor/commands.ts`
  (clone `insertFieldCmd`/`editFieldCmd`, `commands.ts:1015,1031`); one `dispatch` = one undo
  step via the existing `UndoManager` (`editor/undo.ts`).
- **Ribbon:** add an "Equation" button/group on the Insert tab after the image button
  (`editorApp.ts:~1531`); add an icon to `ui/icons.ts`. The Layout-tab page-layout button
  (`editorApp.ts:1619`) is the exact toggle-a-floating-dialog idiom.
- **Selection:** extend `hitTestSelectableObject` (`layout/geometry.ts`) to math boxes; add
  `equationSelected` / `inEquation` to `CurrentFormat` (`index.ts:2960-2998`); gate buttons
  with `enable(...)`.
- **Click-to-edit:** context-menu "Edit Equation…" in `buildContextEntries`
  (`index.ts:~2487`, clone the field-edit entry) and a double-click branch in
  `input/selectionController.ts:388` opening the panel seeded with the existing AST.
- **In-canvas edit mode (optional, later):** early guards in `input/imeProxy.ts` (`onInsertText`
  / `onSplitParagraph`) and `input/keymap.ts` to route `^`/`_`/`\` to equation commands when a
  math caret is active (precedent: the `sdtBlocksEdit()` guards, `index.ts:1838`).

### WS7 — Tests, examples, docs
- Unit: OMML↔MathML conversion (golden pairs, both directions, lossless verbatim fallback).
- Round-trip: `.docx` import→export equality for the supported subset (precedent:
  `export/roundtripStability.test.ts`).
- Render goldens: math-box layout snapshots; canvas vs PDF parity check.
- Builder: a `DocumentBuilder` helper to author equations headless (memory: `builder-features`).
- Showcase: add equations to `sampleDoc`; an `examples/` equation demo.
- Docs: README "How it compares" + `frontend/README.md` limitations — move charts/equations
  off the roadmap line; document the bundled-math-font requirement for export embedding
  (parallel to `cjk.fallbackFont`). Update landing `#matrix` (equations cell), and the
  competitor grid / feature matrix artifacts.

---

## 5. Milestones (build order — all land in v1)

- **[DONE] M0 — Spike:** ✅ OMML↔MathML converter + MathML string parse/serialize. Shipped:
  `shared/src/model/math.ts` (AST), `frontend/src/mathml/{parse,serialize,toOmml,fromOmml,
  xmlNode}.ts`, 22 tests (`mathml.test.ts`). STIX MATH-table → JSON extraction still pending
  (folded into M4 — the default `MathFont` provider uses clone metrics meanwhile).
- **[DONE] M3-core — Typesetter:** ✅ `frontend/src/layout/math/{mathBox,mathFont,layoutMath}.ts`
  — row spacing classes, scripts, fractions, radicals, fences, limits, n-ary, matrix. Pure
  (injected `MathFont`), 11 tests (`layoutMath.test.ts`). Remaining: stretchy delimiters /
  glyph assembly + real STIX constants (need the MATH table — with M4).
- **[DONE] M1 — Model integration (BLOCK/display math):** ✅ `EquationBlock` added to the `Block`
  union; `CharStyle.equation` added for future inline; `setEquation` op (invertible). Layout
  wired by treating an equation as an image-like box (`equationBox` cached typeset) in body +
  cell + band paths (`layout/engine.ts`, `PlacedEquation` in `layoutTree.ts`). Painted on canvas
  (`paint/paintMath.ts` + `renderer.ts`) AND PDF (`export/pdf/paintBlock.ts`). Tests:
  `equationRender.test.ts`, `equationOps.test.ts`. Architecture note: chose inline-run-free BLOCK
  representation to avoid the U+FFFC-sentinel pretext surgery for v1; inline is M5.
- **[DONE] M2 — docx round-trip (BLOCK):** ✅ export emits block-level `m:oMathPara`
  (`documentXml.ts` + `xmlns:m` in `WML_NS`); import parses `m:oMathPara` → `EquationBlock`
  (`documentParser.ts`/`types.ts`/`mapToModel.ts`). Round-trip test green
  (`equationRoundtrip.test.ts`). Inline `m:oMath` on import is WARNED (not dropped silently),
  pending M5.
- **[DONE] M4 — Visual editor (insert + edit):** ✅ `ui/equationEditor.ts` floating dialog
  (MathML editor + template/symbol palette + LIVE canvas preview + Display/Inline toggle) wired
  to Insert ribbon → Equation (`editorApp.ts`); commands `insertEquation`/`insertInlineEquation`/
  `editEquationCmd`/`editInlineEquationCmd`. **Right-click → "Edit Equation…"** on any equation
  (block via `hitTestEquation`, inline via caret-run detection) seeds the dialog and applies in
  place (`index.ts` buildContextEntries). REMAINING: bundle STIX Two Math + MATH-table extraction
  (uses clone metrics today); a LaTeX/AsciiMath input mode (MathML + palette today).
- **[DONE] M5 — Inline math:** ✅ `U+FFFC` sentinel run + `CharStyle.equation` (`styleEq` ref
  guard). Layout: `toItems` reserves the box width via pretext `extraWidth`+`break:'never'`;
  `breakNextLine` swaps box ascent/descent into the line and rides the box on the fragment;
  `geometry.clustersOf` makes it one atomic caret cluster. Paint: canvas (`renderer.ts`) + PDF
  (`paintBlock.ts paintLine`). Round-trip: inline `m:oMath` export (`singleRun`) + import
  (`mathInline` IR → equation run). Tests: `inlineEquation.test.ts`, `equationHit.test.ts`.
  Demo: an inline Pythagorean identity + ½ in a sentence in sampleDoc.
- **[DONE] M6a — STIX Two Math:** ✅ bundled `StixTwoMath-Regular.ttf` (OFL) and routed all math
  glyphs to it across canvas (FontFace in `editorFonts.ts`, awaited before first paint), fontkit
  measurement (`measureHost.ts`), and PDF embedding (`fontRegistry` via `cloneFamilyFor`
  passthrough + `MATH_FONT_FAMILY` in clones.ts). Identifiers remap to the Mathematical
  Alphanumeric block (`mathAlpha.ts`) so STIX renders TRUE italic/bold/blackboard/script glyphs
  (no faux-slant). **Stretchy delimiters**: fences grow to their content height via a size-variant
  approximation centered on the math axis (`delimiterBox` in `layoutMath.ts`) — matrix brackets /
  big parens now scale (verified: parens 22px→52px around a fraction). fontkit can't read the MATH
  table, so constants stay TeX-default. Tests: `mathAlpha.test.ts` + stretchy case in
  `layoutMath.test.ts`. 47 math tests, 701 frontend, 0 regressions.
- **[DONE] M6b — MATH constants + big operators + LaTeX:** ✅
  - **True MATH-table constants**: `scripts/extract-math-constants.mjs` parses STIX's OpenType
    MATH table (fontkit can't) → generated `stixMathConstants.ts`; `mathFont.ts` `stixMathConstants()`
    feeds real axisHeight/rule thickness/script percents/fraction & radical gaps, used whenever the
    math font is STIX.
  - **Big-operator size variants**: `opAtom` enlarges ∑ ∏ ∫ … to `displayOperatorMinHeight` in
    DISPLAY style (centered on the axis), with over/under limits for ∑-class and scripts for ∫;
    inline stays compact. Threaded a `display` flag through `MathStyle`/`equationBox`.
  - **LaTeX input mode**: `mathml/latex.ts` recursive-descent LaTeX→AST (frac/sqrt/scripts/fences/
    matrices/accents/Greek/ops/funcs/styles/spacing; lenient). Wired into `equationEditor.ts` with a
    LaTeX/MathML toggle (LaTeX default for new, one-way LaTeX→MathML convert on switch) + mode-aware
    palette. Tests: `latex.test.ts` (12) + engine layout of a LaTeX quadratic. 714 frontend tests.
- **[REMAINING] — true glyph-assembly stretchy (vs size-scaling), AsciiMath input, structured
  slot-caret editing, docs/landing/`#matrix` announcement.**

---

## 6. Risks & open questions

1. **Math layout engine scope (now the #1 risk).** "Fully editable, no fallback" removes the
   MVP escape hatch, so the typesetter must cover the full OMML surface — the dominant cost by
   far. Mitigate with the layered build order (engine layers 1→5), golden tests gating each
   layer, and leaning hard on documented TeX/MathML rules + STIX MATH constants. This is where
   schedule risk concentrates; treat M3 as the critical path.
2. **fontkit MATH-table support.** May not parse the MATH table natively → prefer the
   build-time JSON extraction (D4) so runtime never depends on it.
3. **Bundle size / font licensing.** A math font adds weight; confirm OFL and the code-split
   lazy-load story (math chunk loads on first equation, like the existing lazy editor/export
   chunks).
4. **Inline offset invariant.** The sentinel-char approach must keep every `text.ts` consumer
   correct (graphemes, word break, selection). Audit + tests before shipping inline.
5. **Converter fidelity (no fallback).** Since unsupported constructs can't degrade to a
   placeholder, gaps in OMML coverage become hard bugs, not soft warnings. Mitigate with a
   broad corpus of real-world OMML round-trip fixtures and the verbatim-string integrity check
   (D1) to detect any node we drop or mangle.
6. **Editor caret-in-tree UX.** A structured math caret is its own mini-editor; the generic-
   AST design (WS5) keeps it bounded — slot navigation + templates + LaTeX expansion cover
   "fully editable" without per-construct UI. In-canvas (vs floating-panel) editing stays an
   optional later enhancement.

## 7. Resolved product decisions
All four open questions are locked (see top of doc): **inline + display in v1**, **STIX Two
Math**, **visual + LaTeX shortcuts**, **fully editable (no opaque fallback)**. No open product
questions remain; the engineering critical path is M3 (full typesetter coverage).
