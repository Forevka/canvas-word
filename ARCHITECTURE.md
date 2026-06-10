# canvas-word — Architecture

A Word-like, page-accurate document editor rendered on `<canvas>`, using
[`@chenglou/pretext`](https://github.com/chenglou/pretext) as the text measurement /
line-breaking engine. Same architectural family as post-2021 Google Docs: the browser
never lays out document text — we do, deterministically.

```
┌─────────────────────────────────────────────────────────────┐
│  Input layer        mouse / keyboard / IME proxy / clipboard │
└───────────────┬─────────────────────────────────────────────┘
                │ commands
┌───────────────▼─────────────────────────────────────────────┐
│  Editor core        EditorState · transactions · undo       │
└───────────────┬─────────────────────────────────────────────┘
                │ ops mutate
┌───────────────▼─────────────────────────────────────────────┐
│  Document model     Document → Block → Paragraph → Run      │
└───────────────┬─────────────────────────────────────────────┘
                │ dirty blocks invalidate
┌───────────────▼─────────────────────────────────────────────┐
│  Layout engine      pretext line-breaking → LineBoxes →     │
│                     pagination → LayoutTree (Pages)         │
└───────────────┬─────────────────────────────────────────────┘
                │ placed fragments (absolute coords)
┌───────────────▼──────────────────────┬──────────────────────┐
│  Paint layer (canvas, virtualized)   │  A11y mirror (DOM)   │
└──────────────────────────────────────┴──────────────────────┘
```

**One-way data flow:** input → command → transaction → new model → incremental layout →
paint (next rAF). Selection/hit-testing *reads* the layout tree; nothing writes upward.

---

## 1. Document model (`shared/src/model/`)

The single source of truth. Pure data, zero DOM/canvas knowledge, fully serializable.
It lives in the `shared` workspace (not `frontend`) so the Node backend reuses the
exact same model, ops, and OT primitives the editor runs.

### Tree shape

```
Document
└─ Section            (page size, margins, headers/footers — one to start)
   └─ Block[]
      ├─ Paragraph    { runs: Run[], paraStyle }     ← 95% of blocks
      ├─ Table        { rows → cells → Block[] }     ← phase 4+
      └─ ImageBlock   { src, size, wrap-mode }
Run = { text: string, charStyle: CharStyle }          ← style-homogeneous span
```

- `CharStyle`: font family/size/weight/italic, color, underline, strikethrough, letter-spacing.
  These map 1:1 onto pretext's `RichInlineItem` font strings.
- `ParaStyle`: alignment, line-height, space before/after, indents, named style ref
  (Heading 1…) resolved through a `Stylesheet` cascade.
- Adjacent runs with equal styles are merged on every edit (normalization pass) — keeps
  run count low, which keeps pretext `prepareRichInline` inputs small.

### Addressing & operations

- `DocPosition = { blockId: string, offset: number }` — offsets are **UTF-16 code unit**
  indices into the paragraph's concatenated run text (same unit pretext/`Intl.Segmenter`
  uses, so no conversion layer).
- Blocks get stable string IDs; layout cache and undo reference IDs, never array indices.
- All mutation goes through **operations**: `insertText`, `deleteRange`, `applyCharStyle`,
  `setParaStyle`, `splitParagraph`, `mergeParagraphs`, `insertBlock`, `removeBlock`.
  Each op application returns its **inverse op** → undo is free, and later this is the
  exact seam where OT/CRDT collaboration slots in.
- Ops also report a **position-mapping function** so selection and any stored positions
  survive edits (insert 3 chars before offset 10 → offset 13).

## 2. Layout engine (`frontend/src/layout/`)

Turns the model into a `LayoutTree` of absolutely-positioned fragments. This is where
pretext lives — and the *only* place it's imported.

### Per-paragraph: pretext does the hard part

```
Paragraph.runs ──► RichInlineItem[] ──► prepareRichInline()      (cached)
                                          │
        maxWidth (content box, per page)  ▼
        ──────────────► layoutNextRichInlineLineRange() loop ──► LineBox[]
```

- **Prepare cache:** `Map<blockId, { revision, prepared: PreparedRichInline }>`.
  `prepareRichInline` does the expensive measurement once; it's invalidated only when a
  paragraph's text/styles change (revision counter bumped by ops). Width changes do
  **not** invalidate it — re-running the line-range loop is cheap arithmetic. This is
  pretext's whole value proposition.
- Each `LineBox` records: y-offset, ascent/height, and `InlineFragment[]` — runs of
  same-styled text with `x`, `width`, source `{blockId, startOffset, endOffset}`, and
  per-cluster advances (for caret math).
- Justified/centered/right alignment is applied here by distributing leftover width —
  pretext gives line content + natural width; alignment is our pass.

### Block flow + pagination

```
for each block:  lines = linesOf(block, contentWidth)
                 fill current page; when page is full → break
```

- A `Page` is `{ index, blocks: PlacedBlock[] }`; `PlacedBlock` holds the block's lines
  with page-relative coordinates.
- Break rules live here: never orphan a single line (widow/orphan control),
  `keep-with-next` for headings, unbreakable blocks (images) push to next page.
- **Line-level breaking** is what makes this Word-like: a 200-line paragraph spans pages
  mid-paragraph, which contenteditable-based editors struggle to do honestly.

### Incrementality

- Ops mark blocks dirty. Re-layout starts at the **first dirty block's page** and walks
  forward; stop early when a block lands at the same `(page, y)` as before (positions
  converged → everything after is unchanged).
- Typical keystroke: 1 paragraph re-broken (cache hit on all others), usually 0–1 pages
  repainted.

### Geometry queries (consumed by input + paint)

- `hitTest(pageIndex, x, y) → DocPosition` — line by y, then cluster by x using fragment
  advances.
- `caretRect(pos) → {page, x, y, height}` and `selectionRects(range) → Rect[]`
  (per-line rectangles; handles multi-page selections).

## 3. Paint layer (`frontend/src/paint/`)

Dumb and fast: takes the `LayoutTree` + `EditorState`, draws pixels. No measurement ever
happens here — that would defeat pretext.

- **Page virtualization:** the scroll container holds sized placeholder divs (one per
  page, exact page height — known without painting, thanks to layout). Only pages
  intersecting the viewport get a real `<canvas>`; ~2 pages of overscan. 500-page
  documents cost ~3 live canvases.
- **DPR-correct:** canvas backing store = CSS size × `devicePixelRatio`; re-rasterize on
  zoom (zoom = scale transform + crisper re-paint, *not* re-layout — layout is in CSS px).
- **Paint order per page:** page background/shadow → selection highlight rects →
  text fragments (`ctx.fillText` per fragment — one call per same-style run, not per
  glyph) → underline/strikethrough lines → embedded objects.
- **Caret is an overlay** (absolutely-positioned 1px DOM div or tiny top canvas): blinking
  must not trigger page repaints.
- **Damage tracking:** repaint only pages whose content or selection changed; within a
  page, optional dirty-rect clip for the common single-line edit.

## 4. Input layer (`frontend/src/input/`)

The hardest layer to get right, because we threw away contenteditable.

### The IME proxy (critical piece)

A visually-hidden but *focusable* `contenteditable` div (the "proxy") that is always
focused while the editor is active, and absolutely positioned at the caret's screen
location (so native IME candidate windows pop up in the right place).

- `beforeinput`/`keydown` → translated to commands; the proxy's own content is thrown away.
- **Composition** (`compositionstart/update/end`): during composition we render the
  composition string *inline at the caret* with underline decoration (a transient
  "phantom run" injected at layout time, not a model edit); on `compositionend` it becomes
  a real `insertText` op. This is the only honest way to do CJK/dead-key input on canvas.
- The proxy doubles as the screen-reader text surface (see a11y).

### Selection controller

- `mousedown → hitTest → anchor`; drag extends `focus`; auto-scroll near viewport edges.
- Double-click = word select (reuse `Intl.Segmenter` word granularity — same segmentation
  pretext uses, so selections align with line-break opportunities), triple-click =
  paragraph.
- Keyboard movement: arrows (cluster-wise, via fragment advances), Ctrl+arrows
  (word-wise), Home/End (line — needs layout, not just model), PageUp/Down,
  shift-variants extend. Up/Down preserve a "goal X" column across lines.

### Clipboard

- **Copy/cut:** serialize selection to `text/html` + `text/plain` via the async Clipboard
  API (HTML flavor is what makes paste into real Word/Docs work).
- **Paste:** sanitize incoming HTML → whitelist-map to runs/paragraph styles → insert as
  a model fragment. Plain-text fallback always works.

### Keymap / commands

Single dispatch table: `Ctrl+B → toggleCharStyle('bold')`, `Ctrl+Z → undo`, `Enter →
splitParagraph`, … Commands are pure `(state) → Transaction | null`, which makes toolbar
buttons and keys share one code path and makes the whole editing surface testable headless.

## 5. Editor core (`frontend/src/editor/`)

```
EditorState = { doc: Document, selection: DocSelection, stylesheet: Stylesheet }
Transaction = { ops: Op[], selectionAfter, inverseOps, origin }
```

- `applyTransaction(state, tr) → state'` then: invalidate layout for `tr`'s dirty blocks,
  schedule paint on next `requestAnimationFrame` (coalesces bursts of input).
- **Undo manager:** stack of inverse-op groups; consecutive typing coalesces into one
  undo step (broken by pauses >1s, selection moves, or style changes — Word behavior).
- State is immutable-by-convention (ops produce new run arrays); enables cheap dirty
  checking and future time-travel/collab.

## 6. Accessibility (`frontend/src/a11y/`)

Canvas text is invisible to assistive tech. Non-negotiable mitigation:

- The IME proxy is upgraded to an **a11y mirror**: it contains the plain text of the
  paragraph(s) around the caret with a real DOM selection kept in sync, so screen readers
  announce what's being read/edited. (Google Docs ships a variant of this.)
- ARIA live region for editor announcements ("Bold on", "Page 3 of 12").
- Full-document export to honest DOM remains available (print preview / reader mode).

---

## Repository layout

An npm-workspace monorepo. The editor layers above all live in the `frontend`
workspace; the model and OT primitives live in `shared` so the backend reuses them.

```
shared/    src/model/ (document, ops, position, stylesheet, text, tableGrid, lists)
           src/ (change, transform, replay, ids)   ← OT collaboration primitives
           src/persist/ (serialize, media)
frontend/  src/layout/  src/paint/  src/input/  src/editor/  src/a11y/
           src/import/docx/   ← .docx import (IMPORT.md)
           src/export/        ← .docx + PDF export, DOM-free measure host (EXPORT.md)
           src/sync/          ← live-collaboration client
           src/media/  src/fonts/  src/ui/  src/app/  src/model/ (sample/stress docs)
           src/wordcanvas.ts  ← @forevka/wordcanvas library entry
           src/editorApp.ts  src/main.ts  ← composition root + demo chrome
backend/   src/ (server, db, store/ChangeStore, export, import, admin, auth, webhooks, openapi)
dashboard/ admin panel (Vite)
examples/  embed-offline, embed-live   web/  Caddyfile (edge)
```

See the **Workspaces** and **Editor layer map** tables in [README.md](./README.md)
for the per-directory breakdown.

### Collaboration & backend (the OT seam, realized)

The "exact seam where OT/CRDT collaboration slots in" promised by the invertible-op
model (§1) is now built. `shared/` holds change/transform/replay primitives; the
editor's `sync/` records local edits and streams them over a WebSocket; the Node
`backend/` transforms and broadcasts them against a Postgres-backed `ChangeStore`,
and runs the same import/export pipelines server-side (so a doc can be rendered to
PDF/DOCX without a browser). Identity, presence, and attribution ride the same
channel; the `dashboard/` and integration tokens sit on top.

## Build order (delivered)

Milestones 1–6 below shipped in order; the editor then closed out the
[ROADMAP.md](./ROADMAP.md) feature plan and grew the collaboration backend,
export pipelines, and distribution layer described above. The full delivery log
is the **Implementation history** table in [README.md](./README.md).

1. **Read-only renderer** — model → pretext layout → single-page canvas paint. Proves the
   pretext integration and the prepare-cache before any editing complexity.
2. **Pagination + virtualized scroll** — multi-page, widow/orphan rules, 1000-page perf check.
3. **Selection + caret** — hit-testing, mouse/keyboard selection, caret overlay. Still read-only.
4. **Editing** — ops, transactions, undo, plain typing through the IME proxy (ASCII first,
   then composition/IME).
5. **Rich text** — char/para styles, toolbar, clipboard HTML in/out.
6. **Word furniture** — tables, images, headers/footers, find/replace, .docx import/export.

## Key risks & mitigations

| Risk | Mitigation |
|---|---|
| IME edge cases (CJK, dead keys, Android GBoard) | Proxy pattern from day one; test matrix early (milestone 4, not 6) |
| Font availability mismatch (layout vs paint) | `document.fonts.load()` before first layout; re-layout on `fonts.ready` |
| pretext API churn (young library) | All pretext calls confined to `frontend/src/layout/` — one adapter module |
| A11y debt | Mirror built in milestone 4 alongside typing, not bolted on |
| Float/decimal coordinate blur | Round paint coords to device pixels; keep layout in float CSS px |
