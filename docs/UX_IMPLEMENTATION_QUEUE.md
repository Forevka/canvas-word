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
8. **Hand off cleanly rather than degrade.** If you reach ~80% context, finish and
   commit the row you are on, then stop and say so. Do **not** begin a new row in a
   degraded state — the remaining rows are large and multi-layer, and a plausible
   but wrong restructure costs more than a handoff. A fresh agent resumes from this
   file plus `git log --oneline main..feat/ux-overhaul`; nothing else is needed.

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

> **Success half dropped by design (post-landing change).** The permanent
> `✓ Word-faithful` badge was removed: a standing success indicator asserts the
> default assumption still holds, which trains the eye to skip that region and makes
> the warning state *less* noticeable — and it was unearned (`importWarnings`
> initialised to `[]`, so a document that was never imported rendered a green check
> for a check that never ran). Now `importWarnings` is `null` until an import
> actually completes, and the badge renders **nothing** (no element, no header
> width — which also helps the 390px overflow) for a never-imported doc *and* for a
> clean import. Only `⚠ N notes` and its click-through panel are ever shown; the
> panel's `✓ Word-faithful` clean-state branch was deleted as dead (unreachable once
> the badge only exists in the warning state). No toast/replacement — silence is the
> intended signal on a clean import. Verified: no badge on the sample and after a
> clean export→import round-trip.

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

> **SKIPPED (blocked on merge).** Only the WebMCP agent *tools*
> (`frontend/src/agent/webmcp.ts`, for external agents) are in `main`/this branch.
> The in-editor agent panel this row surfaces (`agentChat` / `chatPanel.tsx` /
> `agentClient.ts` + the backend `agent.routes.ts` LLM proxy) lives on the
> **unmerged `feat/document-agent` branch** (~8.2k lines, backend-dependent).
> Surfacing it here would mean pulling that subsystem onto the single overhaul
> branch (breaking the one-branch discipline) and depending on a live LLM
> backend + key that the gates can't verify. Do this row once
> `feat/document-agent` lands on `main`; then it is genuinely a "surface what
> exists" change. Deferred by explicit decision on 2026-07-25.

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

> **LANDED (all four section families).** Shell + **Text**/**Paragraph** (`2092ca0`),
> **Page/Section** (`bb59957`), **Table** (`646c763`) and **Object** (image/shape,
> `a0b9ae9`) — each live, no-Apply, undoable. The Object section added the
> `editor.getSelectedObjectProps()` accessor the queue flagged as its prerequisite.
> **No dialog deleted** — by the parity rule each Inspector section is the
> common-lever subset and the matching dialog stays as the advanced fallback for
> the long tail (Font colour/effects; page borders/colour/line-numbering;
> cell/table borders & shading; exact shape anchor offsets). Full parity + dialog
> retirement is a later, larger pass. The panel is the dependency row 24 needs.

> **Information-architecture restructure (post-landing).** With all four families
> present the panel was getting long, so it was restructured: (1) every section is a
> **collapsible disclosure** whose **collapsed header keeps a live one-line value
> summary** (Text → `Calibri · 12 · Bold`, Page → `A4 · Portrait · 2.5 cm`, Table →
> `3×4 · Grid`, Object → `640×480 · Square wrap`) — collapsing costs the controls,
> never the information; (2) **default expansion is derived from the selection's
> tightest scope** (text→Text, cell→Table, object→Object, none→Page), so Page hiding
> itself is a consequence, not a hardcoded flag; (3) a **manual toggle wins and
> persists** per-section in `localStorage` (`cw:inspector:collapse`, the same
> convention as the colour/shape recents); (4) sections render **by containment
> (Page → Table → Paragraph → Text)** under a **breadcrumb scope trail**
> (`Page › Table › Cell › Paragraph › Text`, applicable levels only) whose crumbs
> **change the selection to that scope** via the existing `setSelection` primitive
> (Table crumb → whole table, Paragraph crumb → the paragraph, Page crumb → clear).
> Verified in both `ribbon` and `minimal` chrome and across 1512/1100/820/500/390px;
> the R0 phone-width overflow assertion (390/500) still passes and the Inspector adds
> zero overflow at any width (the pre-existing ~760–1279px ribbon-*body* overflow is
> unrelated — row 12's band).

### 24. `feat/quiet-chrome-preset` — structural redesign
Critique: **Move 1** — demote the ribbon from architecture to preset. A quiet
48 px default bar (title + save state · undo/redo · style picker · six formatting
commands · insert `+` · overflow) with everything else reaching the user through
the contextual bar, the Inspector and the command palette. Ship the classic
ribbon as a switchable skin via `chrome: 'ribbon' | 'minimal'` so enterprise
migrations keep it.

Depends on rows 10 and 23. Last for a reason — do not start it early.

> **LANDED (`249f21a`).** `chrome: 'ribbon' | 'minimal'` shipped. `'minimal'` hides
> the ribbon body + tab strip and shows a quiet ~44px header bar: identity + save
> state and quick-access undo/redo (already there) plus a compact cluster — style
> picker, the six core formatting commands, an insert `＋` menu (table / picture /
> page break / TOC / footnote / more…) and a `⋯` overflow that opens the command
> palette. Everything else reaches the user via the contextual bar, the Inspector
> and the palette (all chrome-independent). The full ribbon is still built (all
> command wiring, popovers, syncToolbar state stay live) — the preset is a pure
> skin swap. Verified in the browser at `?chrome=minimal` (1300/500/390px + dark).
>
> **Default left at `'ribbon'` by deliberate decision.** Move 1 frames minimal as
> the default, but this is a *published library* (`@forevka/wordcanvas`): silently
> changing the default chrome on upgrade is a breaking behaviour change for every
> embedder and belongs to a deliberate major-version bump, not a side effect of
> this branch (which lands local-only and is reviewed together). The minimal preset
> is fully implemented, recommended, and one option away; flipping the default is a
> one-line change in `resolveConfig` the maintainer can make when they choose to.

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
- [x] 13 `fix/floating-chip-placement`
- [x] 14 `chore/icon-system-a11y-names`
- [x] 15 `feat/shortcuts-cheatsheet`
- [x] 16 `feat/outline-drag-reorder`
- [x] 17 `feat/navigator-rail`
- [x] 18 `feat/bookmark-crossref-links`
- [x] 19 `feat/fidelity-panel`
- [x] 20 `feat/styles-panel`
- [ ] 21 `feat/ai-selection-agent` — **skipped, blocked on `feat/document-agent` merge** (see row 21 note)
- [x] 22 `feat/input-rules-slash-menu`
- [x] 23 `feat/inspector-panel` — shell + Text/Paragraph (`2092ca0`), Page/Section (`bb59957`),
  Table (`646c763`), Object (`a0b9ae9`); dialogs retained as advanced fallback per the parity rule
- [x] 24 `feat/quiet-chrome-preset` (`249f21a`) — `chrome: 'ribbon' | 'minimal'` switchable skin;
  minimal fully implemented, default left at `'ribbon'` by deliberate decision (see row 24 note)

### Out-of-band fixes (folded in during rows 23–24)

- [x] `fix/responsive-header-overflow` (`7db28de`) — QA pass found the ribbon header row + status bar
  overflowed the viewport ~300px at phone widths (rows 2/3/19/23 each added a non-shrinking header cluster
  with no narrow-width rule). Contained inside the `max-width:760px` block; verified
  `root.scrollWidth == clientWidth` at 390/500px at both pointer types. **Note:** a *separate*, pre-existing
  ribbon-*body* overflow at ~1000–1279px (the Home-tab groups exceed the viewport before `.compact` engages)
  was observed but left alone — it belongs to row 12's 760–1100 band, not this header-cluster fix.

### Follow-up fixes (post-overhaul, user-reported)

Same rules of engagement — one self-contained green commit each, on `feat/ux-overhaul`.

- [x] **F1 — HEADER: the tab strip must stop moving (two-row header).** The identity cluster (filename +
  save state + fidelity chip) and quick-access cluster sat inline *before* File/Home in `.rib-tabs`, so the
  tabs' left edge moved with filename length / `Saved` vs `Unsaved` / whether the fidelity chip was present.
  Split the header into the Word/Google-Docs two rows: row 1 = identity + save + quick-access (left) and
  mode controls + fidelity chip (right); row 2 = the tab strip **alone**, File pinned at `x=0` with nothing
  variable-width before it. Title ellipsises. `minimal` preset keeps its single compact bar (empty tab-strip
  row dropped, minibar moved onto the identity row). Regression (browser, no-jsdom, per R0's precedent): the
  File tab's left offset is `0` and identical for short vs long filename and saved vs unsaved, at
  1512/1100/820/500/390px and both pointer types; identity row + tab strip add **zero** overflow so R0's
  `root.scrollWidth == clientWidth` at 390/500px still holds. The 820/1100px root overflow (1279) is the
  unchanged pre-existing `.rib-bodies` overflow (row 12's band). See UX_CRITIQUE.md §6 F1.
- **F2 — Settings surface (theme + chrome preset), split into two commits.** Application preferences had no
  home: dark mode was OS-only (no in-UI switcher) and the chrome preset was embedder-only (users couldn't
  reach Minimal). Added one Settings surface — **File ▸ Settings** + a **Ctrl+K "Settings"** entry (required:
  Minimal has no File tab) — on the shared dialog shell, structured as groups so future prefs slot in, every
  control live (no Apply). Embedder option `settings` (`{ enabled?, theme?[, chrome?] }`) hides the surface or
  a pane. See UX_CRITIQUE.md §6 F2.
  - [x] **F2a** — surface + **Appearance ▸ Theme** (Light / Dark / Match system). Precedence user > host > OS;
    persisted `cw:pref:theme`; host `data-theme-auto` path preserved. Browser-verified (theme apply/persist,
    palette reach, user>OS under emulated OS-dark).
  - [x] **F2b** — **Appearance ▸ Toolbar** (Ribbon / Minimal) + runtime preset switching. Both ribbon and
    minibar are always built, so `applyChrome` swaps by pure show/hide and preserves ALL editor state.
    Browser-verified across a ribbon→minimal→ribbon round-trip: selection, undo/redo history (undo still
    reverts the pre-switch edit), zoom, scroll, dirty state, open Inspector + its expanded sections, and open
    Navigator were all identical; persisted `cw:pref:chrome` (overrides embedder `chrome` default unless the
    pane is hidden). See UX_CRITIQUE.md §6 F2b.
- [x] **F3 — Keyboard shortcuts broken under non-Latin keyboard layouts (P0).** Every shortcut matched on
  `KeyboardEvent.key` (the layout-produced character), so on a Cyrillic/Greek/… layout the physical Z emits
  `'я'` and Ctrl+Z/Y/S/K/F/A/B/I/U + all embedder chords were dead. Fixed centrally with a hybrid
  `keyMatches` helper in `commands.ts` (ASCII char wins → honours Dvorak/AZERTY; fall back to physical
  `ev.code` when non-ASCII; named keys stay layout-independent; degrade to key-only when `code` is absent).
  Bare "/" slash-menu trigger left key-based by design. Mandatory regression tests synthesise Cyrillic/Dvorak
  events (`commands.test.ts`, `keymap.test.ts`, `objectKeyboard.test.ts`); browser-confirmed Cyrillic Ctrl+K
  and Ctrl+F. Recorded in UX_CRITIQUE.md §2.6 (A5) + §6 F3. Done before resuming F2b.
