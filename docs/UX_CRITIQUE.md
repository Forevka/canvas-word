# UX critique — canvas-word editor

Reviewed live at `https://doc-editor.forevka.dev/examples/offline/?devMode=true`
on 2026-07-24, at 1512×900, 1100×800, 820×800 and 500×800, plus a code-side
inventory of the ribbon, dialogs, context menus, floating toolbars, panels and
config surface.

---

## 0. Verdict

The engine is a genuine achievement. The chrome is the weakest part of the
product, and it is weak in a specific, fixable way:

> **You didn't clone Word's UI. You cloned Word's *information architecture* —
> and then dropped the affordances that made that architecture survivable.**

Word's ribbon is tolerable because it ships with group captions, three button
size tiers, contextual tabs that appear and disappear with the selection, a
Quick Access Toolbar for undo/save, a backstage view for file state, and thirty
years of muscle memory. canvas-word has the ribbon *shell* with almost none of
that. The result is the worst of both worlds: as unfamiliar as a new product,
as dense as a legacy one, with none of Word's escape hatches.

Second observation, and the more strategic one:

> **The moat is the fidelity, not the chrome.** Nobody will choose this over
> Google Docs because of the ribbon. They will choose it because a .docx goes
> in and comes out byte-faithful, page-accurate, in a browser, embeddable, with
> a real document model. **That is completely invisible in the UI today.**

So the two directions are: make the chrome quiet, and make the fidelity loud.

---

## 1. Credit where it's due

Not everything needs rework. These are good and should be *extended*, not replaced:

| Thing | Why it's good |
|---|---|
| `contextToolbar.ts` priority framework | Genuinely well-designed: single active bar, priority arbitration, viewport-aware, never overlaps the ribbon. This is the right primitive and it's under-used. |
| Page Layout dialog | Tabbed, live preview, unit switcher, draggable. Looks like 2020s software. |
| `+ Insert` chip on empty paragraphs | The one modern interaction in the product. It's a Notion block-inserter in embryo. |
| Contextual richness of the right-click menu | It knows about TOC fields, equations, shapes, content controls, header/footer bands, table cells. The *content* is excellent — only the presentation is wrong. |
| `customizeRibbon` / `contextToolbars` / `commands` embedder API | You already have the machinery to make the chrome a swappable preset. You just haven't used it on yourself. |
| Perf | 45 chars typed in 282 ms, 23-page relayout warm at ~3 ms. Nothing here is bottlenecked by rendering. |
| Panel empty states | Outline / Bookmarks / Review all have real, instructive empty copy. Better than most products. |

---

## 2. What's actually wrong

### 2.1 The command surface

**C1 — Undo and Redo are hidden inside the File tab.**
The two most-used commands in any editor are two clicks deep and invisible from
the default tab. There is no Quick Access Toolbar. Word solved this in 2007;
this is a regression against the thing being cloned.

**C2 — There is no Save concept at all, anywhere.**
No document title in the chrome. No dirty indicator. No autosave. No version
history. No `Ctrl+S` binding — which means `Ctrl+S` fires the *browser's* "save
page" dialog, an actively harmful outcome. The persistence model presented to
the user is two buttons labelled `PDF` and `DOCX` on a tab they have to go
looking for. A user cannot answer "is my work safe?", which is the single
highest-anxiety question in any document tool.

**C3 — The Styles gallery is clipped mid-card at 1512 px** ("Aa…" cut off at the
right edge with an internal scrollbar), on both pointer types. On coarse-pointer
devices the *whole ribbon* overflows and scrolls horizontally — see **R0**,
which is the root cause and a P0 bug.

**C4 — Tab weights are wildly unbalanced, and the ribbon changes height.**
`Layout` contains **two** icon buttons. `Insert` contains **thirty**, in one
undifferentiated row. `Home` is two rows tall (~140 px), `Insert` and `Layout`
are one row (~40 px) — so **the document jumps vertically every time you switch
tabs.** Layout instability on a pure navigation action.

**C5 — No contextual tabs; instead, a graveyard of disabled buttons.**
Contextual tabs (Table Tools, Picture Tools) are the one unambiguously good
ribbon invention, and you don't use it. Instead:
- a permanent top-level `Table` tab whose ten buttons sit there enabled even
  with the caret in a plain paragraph;
- an `Insert` tab carrying **13 permanently disabled** shape buttons whose
  tooltips read *"— select a shape first"*, plus two more for images and two for
  content controls.

So roughly half of `Insert` is dead space, and the enablement policy is
self-contradictory (shapes disable, table commands don't).

**C6 — Icon-only buttons with duplicate and meaningless accessible names.**
The a11y tree contains three buttons named `"A"` and two named `"ab"`. Small
caps / all caps / double-strikethrough render as `AB` / `Ab` / `ab` — three
near-identical text glyphs. `Find` and `Replace` are two separate buttons that
open the identical bar.

**C7 — A disabled stub shipped in the ribbon.**
Home ▸ Paragraph ▸ Sort, tooltip *"not supported by the engine yet"*. Never ship
the absence of a feature as a greyed control; it reads as broken, not as roadmap.

### 2.2 Layering — there is no surface arbitration

**L1 — Modal-ish surfaces stack and occlude each other.**
Live repro: open **Manage Styles** → press **Escape** (dialog does not close) →
press `Ctrl+F` (find bar opens *behind* it) → click **Review** (pane opens fully
occluded by the dialog). Three surfaces alive at once, mutually hiding, with the
document squeezed to a strip. `contextToolbar.ts` arbitrates floating bars
beautifully; nothing arbitrates dialogs, panels and the find bar.

**L2 — Floating dialogs cover the document they are previewing.**
Page Layout opens fixed at top-right over the page, is not resizable, does not
live-preview onto the real document, and renders a tiny abstract rectangle as
"PREVIEW" while the actual page sits right behind it, obscured.

**L3 — The TOC "Update table of contents" chip is sticky and misplaced.**
It renders *on top of* the "Table of Contents" heading it refers to, survives
caret changes and Escape, and at narrow widths floats over the Outline panel,
outside the editor pane entirely.

**L4 — The `+ Insert` chip overlaps body text** instead of sitting in the left
margin gutter, so the affordance hides the content it's offering to extend.

**L5 — The context menu overflows the viewport** with no flip or scroll; at
900 px height the last items are simply unreachable.

**L6 — The find bar hijacks the status bar** — it docks bottom-left over it,
has an unlabelled input, shows `[W]` as a literal text label for "whole word",
and displays no match counter in that state.

### 2.3 Responsive — and one P0 bug behind most of it

**R0 — The mobile layout is applied to full-size touchscreen laptops. This is a
P0 bug and it caused most of what follows.**

`frontend/src/ui/styles.ts:453`:

```css
@media (pointer: coarse), (max-width: 760px) { … }
```

The `(pointer: coarse)` arm has **no width bound**, so the phone layout activates
on *any* device whose primary pointer is coarse, at *any* resolution. That is
Surface devices, touchscreen ThinkPad / XPS / HP business laptops, iPads with a
keyboard, and touch Chromebooks — a large share of enterprise hardware, all
driving a mouse on a 1400–2500 px screen.

Measured on a coarse-pointer device at **1512 × 900**:

| | coarse-pointer (what those users get) | fine-pointer (intended) |
|---|---|---|
| Ribbon group captions | **all 26 hidden** (`.rib-label { display: none }`) | "Clipboard / Font / Paragraph / Styles / Editing" visible |
| Ribbon buttons | 40 × 40 touch targets | ~26 px |
| Ribbon width | `scrollWidth` 1604 > `clientWidth` 1512 → **horizontal scrollbar** | 1512 = 1512, fits |
| Editing group | icon-only | "Find / Replace / Select All" with text labels |
| Outline pane | `position: absolute; z-index: 30` — **overlays** the page | in-flow; page recentres |
| Bookmarks / Activity drawer | `width: 100% !important` — **covers the entire viewport** | 300 px right drawer |

The last row is the severe one. `styles.ts:477` sets
`.cw-float-drawer { width: 100% !important; }`, intended for a phone. On a
1512 px touchscreen laptop, opening **View ▸ Bookmarks** paints an opaque white
`position: fixed`, `z-index: 45` panel over the whole window. The ribbon, the
document and the status bar all disappear. It also intercepts every pointer
event — Playwright's click retry loop reports
`<button title="Close"> from <div class="cw-float-drawer"> subtree intercepts
pointer events` for 30 s straight. **The application is completely unusable
until the user finds the 12 px `×` in the top-left corner.**

Fix: bound the coarse arm — `@media (pointer: coarse) and (max-width: 1100px), (max-width: 760px)`
— or better, drive the whole thing off the existing ResizeObserver `.compact`
container query and use `(any-pointer: coarse)` only to enlarge hit targets.
Touch capability and screen size are different questions and this rule conflates
them.

**R1 — Below ~1100 px the layout genuinely runs out of room.** Even correctly
rendered, the 264 px in-flow Outline plus a 816 px Letter page at 100 % needs
~1100 px before rulers and scrollbars. There is no auto-collapse and no auto-fit.

**R2 — Zoom defaults to 100 % regardless of viewport.** It should default to
fit-width whenever the viewport can't hold fit-page.

**R3 — At 500 px the tab strip truncates with no overflow menu** — `Table`,
`View`, `Developer` vanish with no affordance.

**R4 — The 760–1100 px band is handled by neither layer** (given R0 fixed):
too wide for the phone breakpoint, too narrow for the desktop layout.

### 2.4 The navigation surfaces: Outline and Bookmarks

These deserve their own treatment, because together they show the IA problem in
miniature: **four navigation surfaces over the same document, in three different
UI paradigms, none of them finished.**

| Surface | Paradigm | Opened from |
|---|---|---|
| Outline | in-flow left panel, 264 px | View ▸ Show |
| Bookmarks | `position: fixed` overlay drawer, 300 px | View ▸ Show |
| Organize pages (thumbnails) | modal dialog | Layout ▸ Pages |
| Fields / content controls | no list at all — caret-only | — |

#### Outline

It is a **flat list of buttons**, confirmed in the live DOM:

```html
<button class="outline-item" style="padding-left: 26px">Tables</button>
```

`draggable: false`. No chevrons, no filter input, no context menu. So it is a
*jump list*, not a navigation pane.

**O1 — No drag-to-reorder. This is the big one.** Word's Navigation Pane and
Google Docs' outline both let you drag a heading to move that entire section,
with its children. It is the single most valuable thing a document outline can
do — and it is the one feature that *requires* exactly the structured model you
already have and a DOM-based editor struggles with.

You have already built the hard half of it: `organizePages.ts` drags page
thumbnails to reorder **whole sections** and explicitly "never splits content".
That block-range-move logic is most of what heading-drag needs. You wrote it and
then buried it in a modal behind a config flag.

**O2 — No collapse / expand.** You support Heading 1–9. Any real document's
outline is unusable past ~40 entries without collapsible sub-trees.

**O3 — No filter box.** Word has one. At 100+ headings, scrolling a flat list is
the only way to find anything.

**O4 — No visual hierarchy.** Depth is expressed *only* as `padding-left`
(12 / 26 / 40 px). Same font size, same weight, same colour for H1 and H3. You
cannot scan the structure of the document — which is the entire purpose of an
outline. Levels need weight/size/colour differentiation, or at minimum a hairline
indent guide.

**O5 — Truncation without recovery.** At 264 px, 5 of the 16 entries in the demo
document are ellipsised: *"Shape positioning — wrap, floa…"*, *"Miscellaneous
OOXML — symbols, …"*, *"International text — CJK & bidirectio…"*. The panel is
not resizable and entries do not wrap.

**O6 — No page numbers, no level badges, no current-section indicator** beyond a
row highlight.

**O7 — Closing it is one click; reopening requires knowing it lives under
View ▸ Show.** No edge affordance, no icon rail, no memory of state.

**O8 — Inconsistent default policy** — Outline opens by default, Review and
Bookmarks do not, with no stated rationale.

#### Bookmarks drawer

**B1 — It covers the ribbon and the status bar.** It is `position: fixed;
top: 0; right: 0; height: 100%; z-index: 45`, so at 1512 px it occludes the
top-right of the toolbar — including the **Editing-mode selector and the Review
button**, which become unreachable while it is open. Compare `contextToolbar.ts`,
which deliberately never places a bar over the ribbon. That discipline exists in
the codebase; this drawer ignores it.

**B2 — It is fixed to the viewport, not to the editor.** For an *embeddable*
editor this is simply wrong: dropped into a host page as a widget, the drawer
paints over the host's own chrome. Outline and Review are in-flow panels;
Bookmarks is a window-level overlay. Two paradigms, one product.

**B3 — On coarse-pointer devices it becomes a full-screen modal that blocks the
app.** See **R0**. This is the most severe single defect found in the review.

**B4 — The row layout stretches the name away from its actions.** The name
button is `flex: 1 1 auto` and the ✎ / 🗑 buttons are pinned right, so at 300 px
the actions sit ~230 px from the label they act on; in the broken full-bleed
case, ~1450 px. Put the actions next to the item, or reveal them on hover.

**B5 — Add and rename go through native `prompt()`.** Bookmark names have real
OOXML constraints (no spaces, no leading digit, ≤40 chars, unique) and a `prompt()`
can neither validate nor explain them.

**B6 — No context for any entry.** No page number, no surrounding text, no
creation date, no search, no sort, no count. A bookmark list without context is a
list of opaque identifiers.

**B7 — And the structural problem: bookmarks have no consumer in this UI.**
In Word, bookmarks exist principally to be *targets* — of cross-references
(`REF`), internal hyperlinks, and `PAGEREF`. canvas-word lets you create them,
but the hyperlink UI has no "link to a place in this document" option and there
is no cross-reference feature. So a top-level View toggle and a window-level
drawer are spent on a leaf feature with no payoff — while **undo and redo have
no home at all** (C1). That is the misallocation in one sentence.

**B8 — Wrong menu.** `View ▸ Show` currently mixes true view settings (ruler,
grid, formatting marks, snap) with content/navigation panels (Outline, Bookmarks,
Activity). Those are different categories and users look for them in different
places.

#### What I'd do instead

Collapse all four into **one Navigator panel** with a persistent 48 px icon rail
so it survives narrow viewports:

```
┌────┬───────────────────────────┐
│ ¶  │  Headings ▾   [ filter… ] │   ¶  Headings
│ ▤  │ ─────────────────────────  │   ▤  Pages   (thumbnails, already built)
│ ◇  │ ⠿ canvas-word          p1 │   ◇  Objects (images, shapes, tables, equations)
│ ⚑  │ ⠿   Fields             p1 │   ⚑  Marks   (bookmarks, fields, content controls)
│    │ ⠿   Content controls   p2 │
│    │ ⠿ ▸ Rich text…         p7 │   ⠿ = drag handle → moves the whole section
└────┴───────────────────────────┘        ▸ = collapse/expand
```

Concretely:

1. **Rail + tabs**, one panel, one paradigm, in-flow, resizable, state remembered.
2. **Drag-to-reorder headings**, reusing the `organizePages` section-move logic.
3. **Collapse/expand + filter box + page numbers + level styling.**
4. **Bookmarks become a tab, not a drawer** — and get their payoff: add
   *"Link to a place in this document"* to the hyperlink dialog and a
   cross-reference insert. Then the feature earns its list.
5. **Move panel toggles out of `View ▸ Show`** into the Navigator rail itself;
   leave `View` for actual view settings.

### 2.5 Visual language

**V1 — The palette is literally Word 2010.** `#2b579a` status bar, dark-blue
`File` tab, grey ribbon chrome. It is period-accurate, which is the problem.

**V2 — No dark mode in the shell.** Emulating `prefers-color-scheme: dark`
produces a pixel-identical page. A `DARK_THEME` canvas preset and dark rules for
context bars exist in code; the ribbon, dialogs, panels and status bar are
hard-coded light. In 2026 this reads as unfinished.

**V3 — Selection highlight is too pale** to locate at a glance in body text.

**V4 — Locale decimal commas leak into numeric fields**: font size shows `10,5`,
page width shows `8,5`. Inconsistent with the `10` and `12` shown elsewhere, and
a parsing hazard in a measurement-heavy tool.

**V5 — Style gallery previews paint late** — the Manage Styles list renders with
blank sample rows and an empty PREVIEW pane for a beat before filling in.

**V6 — Iconography is unsystematic**: mixed outline/filled metaphors,
text-as-icon (`AB`, `Ab`, `LTR`, `RTL`, `[W]`), and no size hierarchy — a
20-times-a-session command (Bold) is the same 26 px square as a
twice-a-year command (Double strikethrough).

### 2.6 Accessibility — the serious one

**A1 — The document is invisible to assistive technology.**
The canvas exposes exactly one node: `textbox "Document editor"`, with no
content. A screen-reader user cannot read a single word of the document. This is
inherent to canvas rendering and the standard fix is an off-screen ARIA text
mirror of the visible pages, kept in sync with the model, with caret/selection
reflected via `aria-activedescendant`. There is an `F6` path into the *drawing
object* layer, which suggests the problem was recognised for shapes but not for
text.

For an editor being embedded in enterprise products, this is a procurement
blocker (WCAG 2.1 AA / EN 301 549 / VPAT), not a nice-to-have.

**A2 — Duplicate accessible names** across ribbon buttons (see C6).

**A3 — No keyboard shortcut discovery surface.** No `Ctrl+/` cheat sheet, no
shortcuts dialog, and the tooltips are the only documentation.

**A4 — The ribbon self-scrolls on focus**, moving controls under the pointer.

### 2.7 Missing table stakes for 2026

No command palette. No `Ctrl+P` / print. No `Ctrl+K` for links. No spell check.
No paste-options chip. No templates or new-document experience. No persistent
page thumbnails (only a modal "Organize pages"). No shortcuts help.

And the most striking omission: **the in-editor AI agent and WebMCP tooling that
already exist in this codebase are not surfaced in this build at all.** You built
an agent that operates on a real document *model* — the thing that makes an
LLM dramatically better at documents than it is against a DOM — and the UI
doesn't mention it.

### 2.8 Native `prompt()` / `alert()` still in the product

Drop-down content-control items, hyperlink insert/edit *from the context menu*
(while the ribbon path has a styled popover — two different UIs for one action),
bookmark add and rename, share failure, and *"Place the caret inside a content
control first."* Six-plus places where a browser modal breaks the illusion.

---

## 3. The reimagining

Three structural moves. Each is independently shippable.

### Move 1 — Demote the ribbon from *architecture* to *preset*

Make the default surface a quiet 48 px bar, and keep the classic ribbon as a
switchable skin for enterprises migrating off Word. You already have
`customizeRibbon()`, `contextToolbars`, and a `commands` registry — you have the
machinery to do this without forking the app.

```
┌──────────────────────────────────────────────────────────────────────────────┐
│ ⌂  Quarterly report.docx  ·  Saved      ↶ ↷ │ Normal ▾ │ B I U │ ≡▾ │ ＋ │   │
│                                          … more     [Editing ▾] [Share] [◐] [✦]│
└──────────────────────────────────────────────────────────────────────────────┘
│                                                                              │
│   [ ruler ]                                                        [Inspector]
│   ┌──────────────────────────────┐                                 ┌─────────┐
│   │                              │                                 │ Text    │
│   │        page canvas           │                                 │ Para    │
│   │                              │                                 │ Page    │
└──────────────────────────────────────────────────────────────────────────────┘
  Page 3 of 23 · 4,113 words        Print layout ▾              — ●———— + 100 %
```

- **Left**: document title (currently absent entirely) + live save state.
- **Centre**: undo/redo, style picker, the six formatting commands people
  actually use, an insert `＋`, and `…` for the long tail.
- **Right**: mode, Share, comments, theme, AI.
- Everything else reaches the user through the floating contextual bar (already
  built), the Inspector (Move 2), or the command palette.

**Add a command palette** (`Ctrl+K`) over the existing command registry. It is
the single highest leverage item on this list: it makes all ~145 commands
reachable without any of them needing ribbon real estate, it gives you shortcut
discovery for free, and it's the thing that lets you delete ribbon rows without
losing capability.

**If you keep a ribbon at all**, fix the three cheap things: restore group
captions, add contextual tabs (Table / Picture / Shape / Header-Footer appear on
selection and the permanent `Table` tab and the 13 dead shape buttons go away),
and add a QAT with undo/redo/save.

### Move 2 — One Inspector replaces twelve dialogs

Font, Paragraph, Page Layout, Table Properties, Shape Size & Position, and the
three Style editors are all the same thing: a property sheet for the current
selection. Today each is a separate floating window that covers the document,
requires an explicit **Apply**, and offers no live preview.

Collapse them into a single right-docked **Inspector** whose sections swap with
the selection:

| Selection | Inspector sections |
|---|---|
| caret in text | Text · Paragraph · Style |
| table cell | Cell · Table · Borders · Style |
| image / shape | Arrange · Size & position · Fill & outline · Wrap |
| nothing / margin | Page · Section · Columns · Background |

Rules: **live preview, no Apply button, every change undoable.** That single
change is most of the perceived jump from "2010" to "2026" — it is how Figma,
Framer, Notion and current Word Online all behave.

It also fixes L1 and L2 for free: fewer floating windows, less occlusion, less
arbitration to build.

> **Update (implementation): information architecture.** With all four section
> families present the panel can get long, so it is structured rather than stacked:
>
> - **Every section is a collapsible disclosure**, and a **collapsed header still
>   shows a live one-line value summary** of its own state (Text → `Calibri · 12 ·
>   Bold`, Paragraph → `Left · Single · 0/0 pt`, Page → `A4 · Portrait · 2.5 cm`,
>   Table → `3×4 · Grid`, Object → `640×480 · Square wrap`). Collapsing costs the
>   user the *controls*, never the *information* — a title-only accordion would be a
>   regression. Summaries stay live as the selection and document change.
> - **Default expansion follows the selection's tightest scope** (caret in text →
>   Text; caret in a cell → Table; image/shape → Object; nothing → Page), so "Page
>   hidden by default" falls out as a consequence, not a hardcoded rule. A **manual
>   toggle wins and persists** (localStorage) so the panel never fights the user.
> - Sections render **by containment, outer→inner (Page → Table → Paragraph →
>   Text)**, headed by a **breadcrumb scope trail** (`Page › Table › Cell ›
>   Paragraph › Text`, only the levels that apply). **Clicking a crumb changes the
>   selection to that scope** (click *Table* → whole table selected; *Paragraph* →
>   the paragraph), which then re-derives the expansion.

### Move 3 — Make the fidelity visible

This is the differentiator and it is currently a secret. You already track
element-by-element coverage (`OOXML_COVERAGE.md`) and you already have a
reverse-builder that reports uncovered fields. Turn that into a user-facing
trust signal:

- **On import**, a passive badge — but only in its *warning* half:
  `⚠ 3 features preserved but not editable`, expanding to a plain-language list
  (*"SmartArt diagram on page 4 — kept exactly as-is, can't be edited here"*).
- **A round-trip guarantee statement** in that panel: what is preserved
  byte-faithfully vs. what is remodelled.
- **Optional overlay** highlighting the regions carrying unmodelled content.

No competitor in the browser can show this, because none of them keep the
fidelity. It converts your hardest engineering work into the reason someone
picks you.

> **Update (implementation): the `✓ Word-faithful` success half was deliberately
> dropped.** A permanent success indicator is chrome asserting that the default
> assumption still holds — it trains the eye to ignore that region, which makes the
> *warning* state less noticeable, not more; and it was unearned (a never-imported
> document showed a green check for a check that never ran). The badge now renders
> **nothing** — no element, so it occupies no header width — for a never-imported
> document and for a clean import; it appears **only** as `⚠ N notes` when an import
> actually preserved-but-adapted something. Silence is the intended signal on a
> clean import — no toast, no replacement confirmation. (The `⚠ N` state and its
> click-through panel are unchanged.)

### Supporting moves

**S1 — Merge four navigation surfaces into one Navigator.**
You have Outline, Bookmarks, Organize-pages thumbnails, and content-control /
field lists as four separate UIs. One left panel, four tabs:
`Headings · Pages · Objects · Fields`. Collapsible to a 48 px icon rail — which
also fixes R1, because a rail can survive an 820 px viewport.

**S2 — Styles as a first-class panel, not a property sheet.**
Hover a style → live preview on the document. Click → apply. Right-click →
*Select all instances*, *Update to match selection*, *Rename everywhere*. Show
"direct formatting overrides" as a dismissible chip when the caret sits on
locally-formatted text. For a tool whose reason to exist is structured
documents, styles deserve better than a 24-field form.

**S3 — Surface the AI you already built.**
Selection → `⌘K` → *"make this a table"*, *"tighten this paragraph"*,
*"match the Heading 2 style"*. Because it operates on the document model rather
than a DOM, it can do things Word's own assistant can't — restructure tables,
rewrite fields, regenerate a TOC. This is the clearest leapfrog available.

**S4 — Input rules on the empty-paragraph path.**
`##␣` → Heading 2, `-␣` → bullet, `1.␣` → numbered, `|||␣` → table, `/` → the
`＋ Insert` menu. You already have the block inserter; this is the keyboard door
into it, and it's what makes fast writers stop reaching for the toolbar.

**S5 — Give the header/footer band, TOC, fields and content controls a shared
"structural object" visual language** — same subtle tint, same hover chip, same
edit affordance. Right now each is bespoke.

---

## 4. Prioritised backlog

**P0 — trust and correctness. Ship first.**

| # | Item | Notes |
|---|---|---|
| 0 | **Bound the `(pointer: coarse)` media query.** One-line CSS fix; today it makes the app unusable on touchscreen laptops the moment Bookmarks is opened | R0 — highest severity found |
| 1 | Scope the Bookmarks/Activity drawer to the editor pane, below the ribbon, not `position: fixed` over the window | B1, B2 |
| 2 | Document title + save state in the chrome; bind `Ctrl+S`; dirty indicator | C2 |
| 3 | Undo/Redo out of the File tab into the always-visible bar | C1 |
| 4 | ARIA text mirror of the canvas for screen readers | A1 — enterprise blocker |
| 5 | Surface arbitration: one manager for dialogs / panels / find bar; Escape closes the top surface | L1 |
| 6 | Styles gallery must not clip a half-card at 1366 / 1512 px | C3 |
| 7 | Remove the disabled `Sort` stub; replace 15 disabled contextual buttons with contextual tabs | C5, C7 |
| 8 | Replace all `prompt()` / `alert()` with real UI (bookmark names need validation anyway) | 2.8, B5 |

**P1 — the 2026 feel.**

| # | Item | Notes |
|---|---|---|
| 9 | Outline: collapse/expand, filter box, level styling, page numbers, resizable | O2–O6 |
| 10 | Command palette (`Ctrl+K`) over the existing command registry | Move 1 |
| 11 | Inspector panel; migrate Font, Paragraph, Table Properties first | Move 2 |
| 12 | Dark mode across the shell | V2 |
| 13 | Auto fit-width zoom + auto-collapsing Navigator rail below 1100 px | R1–R4 |
| 14 | Fix chip placement: TOC chip, `+ Insert` gutter, context-menu overflow | L3–L5 |
| 15 | Icon system pass: size tiers, one metaphor family, no text-as-icon, unique a11y names | C6, V6 |
| 16 | Shortcuts cheat sheet (`Ctrl+/`) | A3 |

**P2 — differentiation.**

| # | Item | Notes |
|---|---|---|
| 17 | **Drag-to-reorder headings in the Outline**, reusing `organizePages` section-move logic | O1 — highest-value single feature on this list |
| 18 | Unified Navigator rail (Headings · Pages · Objects · Marks) | S1, §2.4 |
| 19 | Give bookmarks a consumer: link-to-place-in-document + cross-references | B7 |
| 20 | Fidelity / compatibility panel | Move 3 |
| 21 | Styles panel with hover preview + select-all-instances | S2 |
| 22 | AI on selection + agent panel | S3 |
| 23 | Markdown input rules and `/` command | S4 |
| 24 | Quiet-chrome default with `chrome: 'ribbon' \| 'minimal'` embedder option | Move 1 |

---

## 5. The one-line version

Stop treating "looks like Word" as the design goal and start treating
"round-trips like Word" as the design goal. The first one is a liability you
inherited; the second one is the only thing here that nobody else can copy.

## 6. Follow-up fixes (post-overhaul)

Fixes reported by the user after the 24-row overhaul landed. Each is one
self-contained commit on `feat/ux-overhaul`, held to the same rules.

### F1 — The tab strip must stop moving (two-row header)

**Problem.** Rows 2, 3 and 19 each pinned a cluster to the *left* of the tab
strip, inline inside `.rib-tabs`: the document-identity cluster (filename +
`Saved`/`Unsaved changes` + the fidelity `⚠ N notes` chip) followed by the
quick-access undo/redo cluster, all before the `File`/`Home` tabs. Because those
clusters are variable-width — the filename length, the save-state wording, and
whether the fidelity chip is present all change their width — **the tabs' x-
position moved with document state.** `File` and `Home` are muscle-memory
targets; the document was deciding where they lived. This is the same left-of-
tabs pile-up that R0 had to *contain* at phone widths; F1 removes the cause.

**Fix.** Adopt the Word / Google-Docs two-row header. Row 1 (the *identity row*,
~32px) holds title + save state and the quick-access cluster on the left and the
mode controls (Editing/Suggesting/Viewing, Inspector, Review, collapse chevron)
plus the fidelity chip on the right, right-aligned by the existing
`margin-left:auto` on `.cw-header-review`. Row 2 is the tab strip **alone**:
`.rib-tabs` now contains only the scrolling tabs + their overflow `⋯`, with
`padding-left: 0` so `File` sits at the true left edge (`x = 0`) and **nothing
variable-width precedes it, ever.** The title still ellipsises (it is
`flex: 0 1 auto` inside a `min-width: 0` row) rather than pushing anything.

**Why two rows and not just "reserve a fixed width for the identity cluster."**
A fixed reservation would keep the tabs still but would either clip a long
filename or waste space on a short one, and it would still place document state
in the tabs' own row — one careless future insertion re-introduces the drift. A
separate row makes the invariant *structural*: the tab strip's left offset is a
function of layout only, never of content, so it cannot regress by adding another
header widget.

**`minimal` preset.** `chrome: 'minimal'` (row 24) has no tab strip, so its row 2
is empty; it is dropped (`display:none`) and the minibar rides the identity row,
leaving the single quiet ~44px bar row 24 shipped. Unchanged for the user.

**Regression assertion (browser, no-jsdom — matching R0's precedent).** Measured
on the running dev server with Playwright, at 1512/1100/820/500/390px and both
`pointer: fine` and `pointer: coarse` (CDP `Emulation.setEmulatedMedia`): the
`File` tab's `getBoundingClientRect().left` is `0` and **identical** for a short
vs a long filename and for `Saved` vs `Unsaved changes` (`fileLeftInvariant:
true` everywhere); the identity row and the tab strip each report **zero**
horizontal overflow, so `wordcanvas-root.scrollWidth == clientWidth` at 390/500px
holds (R0 preserved). The residual root overflow at 820/1100px (`scrollWidth`
1279) is entirely `.rib-bodies` — the pre-existing Home-tab-groups overflow in
row 12's 760–1100 band, unchanged by this fix (the identity row and tab strip
overflow by 0 at those widths). Reproduce: measure `.rib-tab.file` left with the
default title, then set `.cw-doc-title`/`.cw-save-text` to long strings, force a
reflow, and re-measure — the two must be equal.
