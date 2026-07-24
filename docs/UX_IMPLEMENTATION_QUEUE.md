# UX implementation queue

Working queue for the findings in [UX_CRITIQUE.md](./UX_CRITIQUE.md). One row =
one commit on the shared `feat/ux-overhaul` branch. Work **strictly top to
bottom**; each row is committed before the next begins.

Follow [SDLC.md](./SDLC.md) for everything not stated here.

---

## Rules of engagement

1. **One long-lived branch.** All 24 rows land on a single local branch,
   `feat/ux-overhaul`. Never branch per row, never switch off it. (Logical PR
   branches are sliced out of specific commits later, together.)
2. **One commit per row.** Work strictly top to bottom; each row is a single,
   self-contained commit with a Conventional Commits message.
3. **Every commit is green on its own.** `npm run typecheck`, `npm run test` and
   `npm run build` must pass at that exact SHA — no commit may depend on a later
   fix, because commits are sliced into logical branches after the fact.
4. **Local only.** Do not push, open a PR, or merge. Everything stays on the
   branch until the full set is reviewed together.
5. **No stopping between rows.** After a row is committed, post a short summary of
   what changed and what was verified, then immediately start the next row,
   straight through to row 24.
6. **Report honestly.** If a gate could not be run, say which and why. See the
   known harness constraints in SDLC.md — Playwright cannot drive pointer-capture
   drags, so verify those by op/command unit tests plus structural checks.
7. **Never commit** the scratchpad (`scratchpad/`, git-ignored) or the
   pre-existing `README.md` modification.

### Gates for every row

| Gate | Command |
|---|---|
| Types | `npm run typecheck` |
| Tests | `npm run test` |
| Editor build | `npm run build` |
| Library bundle | `npm run build:lib` (after any `shared/` or entry change) |
| Browser | Playwright against an `examples/` page — click/caret-driven UI paths |

**Every row in this queue is UI work, so every row must be browser-verified at
both pointer types and at 1512 / 1100 / 820 / 500 px.** Emulate `pointer: coarse`
explicitly — the top row of this queue exists because nobody did.

CHANGELOG entry under `## [Unreleased]` in the same commit as the code.

---

## P0 — correctness and trust

### 1. `fix/responsive-breakpoint` — the coarse-pointer bug
Critique: **R0, B1, B2**

- `frontend/src/ui/styles.ts:453` — `@media (pointer: coarse), (max-width: 760px)`
  has no width bound on the coarse arm, so the phone layout applies to every
  touchscreen laptop at any resolution.
- `styles.ts:477` — `.cw-float-drawer { width: 100% !important; }` then makes the
  Bookmarks/Activity drawer cover the entire viewport, opaque, swallowing all
  pointer events. The app is unusable until the user finds the `×`.
- Also fix the drawer's placement generally: it is `position: fixed` to the
  **window**, so it covers the ribbon (Editing-mode select + Review button become
  unreachable) and the status bar, and in an embed it paints over the host page.

**Acceptance:** on a coarse-pointer device at ≥1100 px the desktop layout renders
(group captions visible, ribbon does not overflow, Outline in-flow). Drawers are
scoped to the editor pane, below the ribbon, never window-fixed. Touch hit-target
sizing still applies on genuinely small screens — prefer `any-pointer: coarse` for
hit targets and a width/container query for layout. The existing ResizeObserver
`.compact` path is the right lever for layout; use it.

### 2. `feat/doc-identity-save-state`
Critique: **C2**

Document title in the chrome, live save state, dirty indicator, `Ctrl+S` bound
(today it fires the browser's save-page dialog). Wire to the existing `onSave`
hook; when no host handler is supplied, degrade honestly rather than claiming
"Saved".

**Acceptance:** the user can answer "what is this file and is my work safe?"
without opening a menu.

### 3. `feat/quick-access-undo-redo`
Critique: **C1** — undo/redo are buried in the File tab. Surface them where they
are always visible, plus a minimal quick-access cluster.

### 4. `feat/canvas-aria-mirror`
Critique: **A1** — the canvas exposes one empty `textbox "Document editor"`; a
screen-reader user cannot read a single word. Build an off-screen ARIA text
mirror of the laid-out pages, kept in sync with the model, with caret and
selection reflected (`aria-activedescendant`). Mirror the discipline already used
for the `F6` drawing-object layer.

**Acceptance:** NVDA or Narrator can read the document body, report the caret
position, and announce the selection. Largest row in P0 — split if needed.

### 5. `feat/surface-arbitration`
Critique: **L1** — dialogs, panels and the find bar have no manager. Repro:
Manage Styles → `Escape` (does not close) → `Ctrl+F` (find bar opens behind it) →
Review (opens fully occluded). Build one manager owning z-order, exclusivity and
`Escape`. `ui/contextToolbar.ts` already does this well for floating bars —
extend that model rather than inventing a second one.

### 6. `fix/ribbon-cleanup-contextual-tabs`
Critique: **C3, C5, C7**

- Styles gallery clips a half-card at 1366/1512 px.
- Remove the disabled `Sort` stub ("not supported by the engine yet").
- Replace the 15 permanently-disabled contextual buttons (13 shape, 2 image, 2
  content-control) and the always-present `Table` tab with real **contextual
  tabs** that appear on selection.

### 7. `feat/replace-native-prompts`
Critique: **2.8, B5** — native `prompt()`/`alert()` in 6+ places: drop-down
control items, hyperlink insert/edit from the context menu (the ribbon path
already has a styled popover — two UIs for one action), bookmark add/rename,
share failure, "Place the caret inside a content control first."

Bookmark names need real validation anyway (OOXML: no spaces, no leading digit,
≤40 chars, unique) — a `prompt()` can neither validate nor explain that.

### 8. `fix/ui-polish-sweep`
Critique: **V3, V4, V5, L6**

- Selection highlight contrast (currently too pale to locate).
- Locale decimal commas leaking into numeric inputs (`10,5`, `8,5`).
- Style-gallery previews paint late (flash of blank sample rows).
- Find bar: unlabelled input, `[W]` as a literal label, overlaps the status bar,
  no match counter in that state.

---

## P1 — the 2026 feel

### 9. `feat/outline-pane-upgrade`
Critique: **O2–O6** — collapse/expand sub-trees, filter box, real level styling
(depth is currently `padding-left` only — same size/weight/colour for H1 and H3),
page numbers, resizable panel, no unrecoverable truncation.

### 10. `feat/command-palette`
Critique: **Move 1** — `Ctrl+K` over the existing `commands` registry. Highest
leverage row in P1: it makes all ~145 commands reachable without ribbon real
estate, and gives shortcut discovery for free. Prerequisite for row 23.

### 11. `feat/dark-mode-shell`
Critique: **V2** — `prefers-color-scheme: dark` currently produces a
pixel-identical page. A `DARK_THEME` canvas preset and dark rules for context
bars exist; ribbon, dialogs, panels and status bar are hard-coded light.

### 12. `feat/fit-width-zoom-narrow-layout`
Critique: **R1–R4** — default to fit-width when the viewport cannot hold
fit-page; auto-collapsing panel rail below ~1100 px; tab-strip overflow menu
instead of silent truncation; handle the 760–1100 px band.

### 13. `fix/floating-chip-placement`
Critique: **L3–L5** — the TOC "Update table of contents" chip renders over the
heading it refers to and survives caret changes and `Escape`; the `+ Insert` chip
overlaps body text instead of sitting in the margin gutter; the context menu
overflows the viewport with no flip or scroll.

### 14. `chore/icon-system-a11y-names`
Critique: **C6, V6** — three buttons named `"A"`, two named `"ab"`; `AB`/`Ab`/`ab`
for all-caps/small-caps/double-strike; text-as-icon (`LTR`, `RTL`, `[W]`); no size
hierarchy (Bold is the same 26 px square as Double strikethrough). One metaphor
family, size tiers by frequency, unique accessible names.

### 15. `feat/shortcuts-cheatsheet`
Critique: **A3** — `Ctrl+/`. Generate from the keymap so it cannot drift.

---

## P2 — differentiation

### 16. `feat/outline-drag-reorder`
Critique: **O1** — highest-value single feature in this document. Drag a heading
to move its whole section, children included. Reuse the block-range move logic in
`ui/organizePages.ts`, which already reorders whole sections and "never splits
content". Needs row 9 landed first.

### 17. `feat/navigator-rail`
Critique: **S1, §2.4** — one Navigator with a 48 px icon rail replacing four
navigation surfaces currently in three paradigms: Outline (in-flow panel),
Bookmarks (fixed drawer), Organize-pages (modal), fields/content-controls (no
list at all). Tabs: `Headings · Pages · Objects · Marks`. Resizable, state
remembered. Move panel toggles out of `View ▸ Show` (**B8**) and leave `View` for
actual view settings.

### 18. `feat/bookmark-crossref-links`
Critique: **B7** — bookmarks currently have no consumer, which is why the drawer
feels pointless. Add "link to a place in this document" to the hyperlink UI and a
cross-reference insert (`REF` / `PAGEREF`). Full DoD applies: round-trip,
`sampleDoc.ts`, `DocumentBuilder`, C# bindings + `csharpParity`, C# showcase.

### 19. `feat/fidelity-panel`
Critique: **Move 3** — make the moat visible. On import, a passive badge:
`✓ Word-faithful` or `⚠ N features preserved but not editable`, expanding to a
plain-language list. Drive it from the existing coverage tracking
(`OOXML_COVERAGE.md`, the reverse-builder's uncovered-field report). This is the
one thing no browser competitor can show, because none of them keep the fidelity.

### 20. `feat/styles-panel`
Critique: **S2** — styles as a first-class panel, not a 24-field property sheet.
Hover to live-preview on the document, click to apply, right-click for *Select
all instances* / *Update to match selection* / *Rename everywhere*. Direct-
formatting-override chip when the caret sits on locally formatted text.

### 21. `feat/ai-selection-agent`
Critique: **S3** — the in-editor agent and WebMCP tooling already exist in this
codebase and are not surfaced in the build at all. Selection → `⌘K` → natural
language. It operates on the document *model*, not a DOM, so it can restructure
tables, rewrite fields and regenerate a TOC. Clearest leapfrog available.

### 22. `feat/input-rules-slash-menu`
Critique: **S4** — `##␣` → Heading 2, `-␣` → bullet, `1.␣` → numbered, `/` opens
the block inserter. The `+ Insert` chip is already the menu; this is the keyboard
door into it.

### 23. `feat/inspector-panel` — structural redesign
Critique: **Move 2** — one right-docked Inspector replaces twelve dialogs (Font,
Paragraph, Page Layout, Table Properties, Shape Size & Position, the three style
editors). Sections swap with the selection. **Live preview, no Apply button,
every change undoable.** This is most of the perceived jump from 2010 to 2026.

Expect to split: land the panel shell plus Font/Paragraph first, then Table, then
Page/Section, then Object, retiring each dialog as its section lands. Do not
delete a dialog until its Inspector section is at parity.

### 24. `feat/quiet-chrome-preset` — structural redesign
Critique: **Move 1** — demote the ribbon from architecture to preset. A quiet
48 px default bar (title + save state · undo/redo · style picker · six formatting
commands · insert `+` · overflow) with everything else reaching the user through
the contextual bar, the Inspector and the command palette. Ship the classic
ribbon as a switchable skin via `chrome: 'ribbon' | 'minimal'` so enterprise
migrations keep it.

Depends on rows 10 and 23. Last for a reason — do not start it early.

---

## Progress

Tick a row only after it is **committed** to `feat/ux-overhaul`.

- [x] 1 `fix/responsive-breakpoint`
- [x] 2 `feat/doc-identity-save-state`
- [x] 3 `feat/quick-access-undo-redo`
- [x] 4 `feat/canvas-aria-mirror`
- [x] 5 `feat/surface-arbitration`
- [x] 6 `fix/ribbon-cleanup-contextual-tabs`
- [x] 7 `feat/replace-native-prompts`
- [x] 8 `fix/ui-polish-sweep`
- [x] 9 `feat/outline-pane-upgrade`
- [x] 10 `feat/command-palette`
- [x] 11 `feat/dark-mode-shell`
- [x] 12 `feat/fit-width-zoom-narrow-layout`
- [ ] 13 `fix/floating-chip-placement`
- [ ] 14 `chore/icon-system-a11y-names`
- [ ] 15 `feat/shortcuts-cheatsheet`
- [ ] 16 `feat/outline-drag-reorder`
- [ ] 17 `feat/navigator-rail`
- [ ] 18 `feat/bookmark-crossref-links`
- [ ] 19 `feat/fidelity-panel`
- [ ] 20 `feat/styles-panel`
- [ ] 21 `feat/ai-selection-agent`
- [ ] 22 `feat/input-rules-slash-menu`
- [ ] 23 `feat/inspector-panel`
- [ ] 24 `feat/quiet-chrome-preset`
