# UX Analysis — Drawing-Shapes Suite & Editor Interaction

**Scope:** the drawing-shapes feature shipped across PRs #223–#238 (insert-shape gallery, fill/outline, wrap/anchor/z-order, read-only + editable text boxes, custom geometry, grouped shapes, the rotate handle, image rotation) plus the surrounding editor interaction model.
**Method:** the dev server was driven live with Playwright — ribbon, gallery, selection, floating toolbar, fill/outline popovers, and insert behavior were exercised and screenshotted; pointer-capture gestures (resize / rotate / drag-move / drag-select) can't be scripted, so those were assessed structurally (handle geometry, cursors, commit logic) and by reading the code and its tests. Findings cite `file:line`.
**Date:** 2026-07-22.

The shapes work is genuinely strong under the hood: lossless DrawingML round-trip, shapes reuse the image anchor/wrap/z-order model verbatim (`commands.ts:2611` — *"Shapes reuse ImageBlock's anchor shape verbatim"*), aspect-ratio corner resize, Shift-to-15° rotation snapping, single-step undo per op, and excellent rendering (gradients, rotation, text clipping). The gaps are almost entirely in the **interaction surface** — the last-mile polish that PR #9 in `docs/SHAPES_PLAN.md:222` was chartered to deliver ("complete right-click menu; keyboard … arrow-nudge; a11y labels; audit that **every** shape feature is reachable via the floating popup **and** the ribbon **and** the menu"). That audit has not fully landed, and this report is largely a map of what's still open against it.

---

## Executive summary

- **Inserting a shape gives almost no feedback.** The new shape is dropped at the caret, **not selected**, and the view doesn't move to it — so nothing appears to happen, and none of the sizing/formatting affordances are armed. Insert into a field (e.g. the TOC) even splits and corrupts the field silently. This is the single biggest first-run problem.
- **The three control surfaces disagree.** Fill and Outline live only on the ribbon and floating toolbar and are **absent from the right-click menu**; the right-click menu is the *only* place that exposes "Behind / In Front of Text." A user who works via one surface can't reach features that exist on another — the exact thing PR #9 was meant to prevent.
- **Shapes are effectively mouse-only and invisible to assistive tech.** There is no keyboard way to select, insert, move, resize, rotate, or format a shape; arrow-key nudge (an explicit PR #9 goal) doesn't exist; the canvas exposes as a single opaque `textbox` with no ARIA representation of shapes at all.
- **No precision authoring.** No numeric size/position/rotation entry and no angle readout while rotating — you can only eyeball angles (or Shift-snap to 15°). There is no Format-Shape panel.
- **Off-page loss risk.** Dragging an anchored shape has no page-bounds clamp; a shape can be moved fully off-page, and because it isn't keyboard-selectable there's no way to get it back.

**Overall:** the feature is production-quality at the model/render layer and demo-quality at the interaction layer. None of the findings below are broken *round-trips* — they're interaction, discoverability, and accessibility gaps.

## Top 5 quick wins

1. **Select (and reveal) the shape you just inserted.** After `editor.dispatch(insertShape(...))` in `editorApp.ts:1302`, select the new block and scroll it into view. One change fixes the "nothing happened" problem and arms every handle/toolbar affordance for free. *(→ F1)*
2. **Add Fill and Outline to the right-click menu.** The popovers already exist (`editorApp.ts:1053–1146`); wire two entries into `buildContextEntries` at `index.ts:3172`. Closes the sharpest cross-surface inconsistency. *(→ C1)*
3. **Show a live angle readout while rotating.** `applyRotatePreview` (`objectController.ts:418`) already computes `deg`; render it as a small badge near the handle and snap-highlight at 15° multiples. *(→ B1)*
4. **Make the gallery keyboard-operable.** The gallery cells are plain `<div>`s with no role/tabindex (`editorApp.ts:1294`). Give them `role="button"`, `tabindex`, arrow-key roving focus, and Enter/Space to insert. *(→ E1)*
5. **Add arrow-key nudge for a selected shape (Shift = larger step) and clamp moves to the page.** Explicitly on the PR #9 task list (`SHAPES_PLAN.md:223`) and still missing; also removes the off-page-loss trap. *(→ E2, F2)*

---

## Severity legend

- **Critical** — blocks a core task for a class of users, or silently corrupts/loses content.
- **Major** — significant friction on a common task, or a whole surface/audience is excluded.
- **Minor** — noticeable rough edge with an easy workaround.
- **Nit** — polish / consistency detail.

---

## A. Discoverability & insertion

### A1 — Inserting a shape produces no visible feedback · **Major**
Verified live: with a caret in the body, Insert → shapes → Rectangle inserts the block but leaves `getSelectedShape() === null` — no handles, no floating toolbar, and the viewport does not move to the shape. To a user it reads as "nothing happened," and none of the resize/format affordances are armed until they *find* the shape and click it. Word selects a freshly inserted shape and shows its handles immediately.
**Fix:** on insert, select the new block and `revealBlock` it (`editorApp.ts:1302`; `insertShape` at `commands.ts:1860`). See Quick Win #1.

### A2 — Insert-at-caret can land inside a field and corrupt it · **Major**
`insertShape` splits the caret paragraph and drops the shape inline wherever the caret is (`commands.ts:1860`). Live test: with the caret in the Table of Contents, the inserted rectangle split the TOC and pushed entries apart — a broken result with no warning. There is no guard against inserting into a TOC/field/other structured region.
**Fix:** block (or redirect to the nearest body paragraph) shape insertion when the caret is inside a field/TOC/SDT, or insert as an anchored/floating object rather than splitting the flow.

### A3 — Gallery is flat and thin vs Word/PowerPoint · **Minor**
The gallery offers 8 presets + Text Box in one flat 4-column grid with no categories and no search (`editorApp.ts:1283–1291`): Rectangle, Rounded rectangle, Ellipse, Triangle, Diamond, Right arrow, Left arrow, Line. Word/PPT group Lines, Basic Shapes, Block Arrows, Flowchart, Callouts, Stars. Common shapes users will reach for (star, callout/speech bubble, connector/elbow line, pentagon/chevron, oval-callout) are absent. For 8 items a flat grid is fine; the gap is **coverage**, and the flat layout won't scale if presets grow.
**Fix:** as the preset set grows, add category headers; keep the live hover-name label (`editorApp.ts:1299`), which is a nice touch. Consider a small "recently used" row.

### A4 — Only one entry point; no inline "+" and no shortcut · **Minor**
Insertion is exclusively the ribbon Insert → shapes popover. The empty-paragraph "＋ Insert" menu (`ui/insertMenuToolbar.ts:1`) deliberately covers only "heading, list, table, page break, footnote" — shapes are absent — and there is no keyboard shortcut for shapes anywhere in `keymap.ts`. Users who live in the inline menu never discover shapes.
**Fix:** add a "Shape ▸" entry to the inline insert menu; optionally a shortcut.

---

## B. Affordances & feedback

### B1 — No angle readout while rotating · **Major**
The rotate handle is well built (arc-arrow glyph, `cursor:grab`→`grabbing`, Escape cancels, CSS-preview then a single commit on pointer-up — `objectController.ts:209–231, 418–463`), but there is **no numeric degree feedback** during the drag. `applyRotatePreview` only sets `transform: rotate(${deg}deg)` (`objectController.ts:421`). Snapping is Shift → 15° only (`snapRotation`, `objectController.ts:22`), with no visual indication that a snap occurred. Hitting a specific angle (say 45°) means eyeballing it.
**Fix:** render `deg` as a badge near the handle during the drag and highlight when a 15° snap engages. See Quick Win #3.

### B2 — No numeric size / position / rotation entry (no Format-Shape panel) · **Major**
Everything is drag-only. There is no field to type a width, height, X/Y offset, or exact rotation — no Format-Shape / Layout dialog analogue. Precise or repeatable layouts are impossible, and it compounds A1 (can't nudge into place) and B1 (can't set an exact angle). The floating toolbar (`ui/shapeContextToolbar.ts:53–68`) is fill / outline / add-text / wrap / align / z-order / delete only; rotate isn't even on it (handle-only).
**Fix:** add a "Size & position…" dialog (reuse the `pageLayout.ts` dialog-shell pattern) with numeric width/height/rotation and offset fields.

### B3 — Rotate handle can collide with the floating toolbar and the `e` resize handle · **Minor**
The rotate handle is absolutely pinned at `left: calc(100% + 8px); top: calc(50% - 9px)` (`objectController.ts:218`) regardless of shape size or viewport position, and the floating toolbar renders above the shape. On a shape near the right/top edge, or a very small shape, the rotate button overlaps the `e` midpoint handle and/or the toolbar. Observed: with a shape near the top of the page the toolbar rendered over the body text just under the ribbon.
**Fix:** flip the handle to the left/top when space is tight; clamp the toolbar within the page and offset it from the handle.

### B4 — Selection handles don't adapt to tiny shapes · **Minor**
Eight 8px resize handles plus the 18px rotate button are positioned without collision avoidance; resize is clamped to `MIN_SIZE_PX = 16` (`objectController.ts:169`) but `insertShape`/import have no lower bound, so a sub-20px shape ends up smothered by its own handles.
**Fix:** below a threshold, drop midpoint handles (corners only) and/or render handles outside the box.

---

## C. Consistency across surfaces (ribbon vs floating toolbar vs right-click)

PR #9's explicit acceptance test is that *every* shape feature is reachable from the floating popup **and** the ribbon **and** the menu (`SHAPES_PLAN.md:224`). It isn't. A per-surface matrix:

| Action | Ribbon "Shape" group | Floating toolbar | Right-click menu |
|---|:--:|:--:|:--:|
| Fill | ✅ | ✅ | ❌ |
| Outline (colour/width/dash) | ✅ | ✅ | ❌ |
| Add / edit text | ❌ | ✅ | ✅ |
| Wrap: in-line / square | ✅ | ✅ | ✅ |
| Behind / In front of **text** | ❌ | ❌ | ✅ (only here) |
| Align L/C/R | ❌ | ✅ | ✅ |
| Bring to front / send to back | ✅ | ✅ | ✅ |
| Rotate | ❌ | ❌ | ❌ (handle only) |
| Delete | ✅ | ✅ | ✅ |

### C1 — Fill & Outline missing from the right-click menu · **Major**
`buildContextEntries` (`index.ts:3144–3182`) builds Wrap, Align, Add/Edit Text, z-order and Delete for a selected shape but **not** Fill or Outline. A user who recolors via right-click (the natural instinct) can't — the feature is invisible to them. The popovers already exist and are shared (`editorApp.ts:1053–1146`).
**Fix:** add Fill and Outline entries at `index.ts:3172`. See Quick Win #2.

### C2 — Two overlapping "depth" concepts, split unevenly across surfaces · **Major**
There are two distinct ordering controls with near-identical names: **layer vs text** ("Behind Text" / "In Front of Text" → `setShapeLayer`) and **z-order among drawings** ("Bring to Front" / "Send to Back" → `bringShapeToFront`/`sendShapeToBack`). The right-click menu exposes both (`index.ts:3148–3176`); the ribbon and floating toolbar expose **only** z-order. So the entire "behind/in front of text" concept is discoverable *only* by right-clicking, and the two similarly-worded controls sitting side by side in that menu invite confusion.
**Fix:** expose "behind/in front of text" on the floating toolbar too, and relabel to disambiguate (e.g. "Move behind text" vs "Bring forward (among shapes)"), matching Word's separate "Wrap Text ▸ Behind/In Front" and "Bring Forward/Send Backward" groupings.

### C3 — "Add text" absent from the ribbon Shape group · **Minor**
Add/edit-text is on the floating toolbar and right-click menu but not the ribbon Shape group (`editorApp.ts` Shape group buttons). Minor, but it's another cell in the reachability matrix that fails the PR #9 audit.

### C4 — Delete-tooltip and toolbar-styling drift · **Nit**
The shape and image bars append the shortcut hint ("Delete shape (Del)"), but the table bar's delete tooltips don't ("Delete rows"/"Delete columns", `ui/tableContextToolbar.ts:73`). The shape bar reuses the image bar's CSS class `cw-img-toolbar` (good, visually identical) while the table bar uses a different class `cw-tablebar` (`ui/tableContextToolbar.ts:49`) — a subtle visual inconsistency across the three floating bars.
**Fix:** standardize the "(Del)" hint and converge the three bars on one base class.

---

## D. Fill / outline depth

### D1 — Fill is solid-only; the showcase shows gradients the UI can't create · **Minor**
The fill popover is the shared 18-swatch palette + "No fill" + "More colours…" (`editorApp.ts:1053–1067`), solid colour only. Yet the "Drawing shapes" section's hero rectangle is a blue→purple **gradient** (it round-trips from import, but there is no UI to author or edit a gradient, picture fill, theme colour, recent-colour row, or eyedropper). A user who sees the gradient in the sample can't reproduce it.
**Fix:** if gradient/picture fill stays out of scope, that's defensible — but consider theme colours and a recent-colours row for parity with the rest of the editor's colour pickers. The outline popover is nicely complete by contrast (colour + width `SHAPE_STROKE_WIDTHS` + dash `SHAPE_DASHES`, `editorApp.ts:1068–1146`).

### D2 — Fill/outline width & dash selects drift from the swatch styling · **Nit**
The width/dash `<select>`s are inline-styled (`min-width:52px;font-size:12px;color:#605e5c`, `editorApp.ts:1097`) rather than using the palette's class system, so they look slightly foreign inside the same popover.

---

## E. Keyboard & accessibility

This is the weakest area and the one furthest from the PR #9 charter. `a11y/mirror.ts:1` is candid: *"Canvas text is invisible to assistive tech… A full mirror … is milestone 6; this is the minimum honest version."* For shapes specifically the state is:

### E1 — Shape gallery is not keyboard-operable · **Major**
Gallery cells are plain `<div class="cell">` with click/mouseenter listeners and no `role`/`tabindex` (`editorApp.ts:1294, 1310`). A keyboard user can open the popover (the ribbon control is a real `<button>`) but cannot Tab to or activate a preset. Confirmed in the a11y tree: the presets surface as `generic` nodes, not buttons.
**Fix:** `role="button"` + roving `tabindex` + arrow-key navigation + Enter/Space. See Quick Win #4.

### E2 — No arrow-key nudge; shapes can't be moved by keyboard at all · **Major**
The arrow handlers operate purely on the text caret and bail when there's no text selection (`selectionController.ts:781`), so a selected *object* never reaches a move path. Arrow-nudge for anchored shapes is an explicit PR #9 task (`SHAPES_PLAN.md:223`) and is not implemented. Combined with E3 there is **no** keyboard path to position a shape.
**Fix:** on a selected shape, arrows nudge by 1px (Shift = 10px / grid step), routing through `moveAnchoredShape`. See Quick Win #5.

### E3 — No Tab order to select a shape · **Major**
`Tab` handles table-cell navigation only (`selectionController.ts:769`); there is no object focus ring. A keyboard user cannot *select* a shape in the first place, which strands E2/B-series entirely. A keyboard user can only Delete or Enter-to-edit a shape that is *already* selected by mouse (`selectionController.ts:745–767`).
**Fix:** include drawing objects in a document Tab/F6 cycle so they can be reached and then nudged/deleted.

### E4 — Shapes have zero screen-reader representation · **Major** (systemic canvas a11y is **Critical**)
`mirror.ts` mirrors only the paragraph text under the caret (`mirror.ts:29–34`); selecting, moving, or rotating a shape announces nothing, and there is no accessible name/role/description for any shape. The rotate handle does carry `role="button"` + `aria-label="Rotate"` (`objectController.ts:216`) — good, but it's an island. The broader "canvas is one opaque `textbox`" problem is a known milestone-6 item; for shapes it means the feature is unusable with a screen reader.
**Fix:** as the a11y mirror lands, emit a live-region announcement on shape select/move/rotate and give each shape an accessible name (preset + any text).

---

## F. Edge cases & failure modes

### F1 — (see A1) new shape not selected — repeated here because it's also the entry to every other affordance.

### F2 — Anchored shape can be dragged fully off-page and lost · **Major**
`moveAnchoredShape` writes raw `offsetXPx/offsetYPx` with no page-bounds clamp (`commands.ts:2725`). A shape dragged past the page edge disappears, and because it isn't keyboard-selectable (E3) there's no way to retrieve it short of undo.
**Fix:** clamp offsets to the page (with a small overhang) on commit, or add an "off-page objects" affordance. Pairs with Quick Win #5.

### F3 — Text overflow in a shape is silently clipped · **Minor**
Shape text is vertically centered and hard-clipped to the box (`engine.ts:2746`, `renderer.ts:1249 ctx.clip()`) with no shrink-to-fit, autofit, or box auto-grow. Type more than fits and it just vanishes below the fold with no indication.
**Fix:** at minimum an overflow indicator; ideally Word's "Resize shape to fit text" / "Shrink text on overflow" options.

### F4 — Shape text inside a table cell is silently read-only · **Minor**
`shapeHasEditableText` returns false for non-top-level shapes (`index.ts:1340`), and the right-click "Add/Edit Text" entry is omitted for them (`index.ts:3172`). Double-clicking such a shape does nothing, with no explanation.
**Fix:** either support it or show a disabled/explanatory state ("Text editing isn't available for shapes inside table cells").

### F5 — Groups render but can't be authored or edited · **Major**
Grouped shapes (`wpg:wgp`) import, render, rotate, and round-trip, but there is **no** `groupShapes`/`ungroup` command, no multi-select, no duplicate (Ctrl+D), and no align/distribute-multiple (grep for `groupShapes`/`ungroup` returns nothing). A user cannot create a group, enter a group to edit a child, or ungroup one — features that arrive only via import or the builder. Group enter/exit was also a named PR #9 keyboard task.
**Fix:** add multi-select → group/ungroup and group child-entry; even a minimal "Group / Ungroup" pair would close a large capability gap.

### F6 — Stale "read-only text box body" comments · **Nit** (report-only, not user-facing)
`engine.ts:2739` and `renderer.ts:1232` still say *"Read-only text box body"* although #235 made top-level text boxes editable (only cell-nested ones are read-only). Harmless but misleading to the next maintainer.

---

## G. Word / PowerPoint convention alignment

**Matches (good):** aspect-ratio corner resize is explicitly "Word behavior" (`objectController.ts:378`); Shift → 15° rotation snap; double-click / Enter to edit text; text-box body insets match Word's `wps:bodyPr` defaults (`engine.ts:2704`); rotated text rotates with the geometry ("Word rotates both", `shapeRotatedText.test.ts:1`); `line` preset treated as the OOXML box diagonal.

**Diverges (by design or by gap):**
- **Click-to-place, not drag-to-draw.** A preset drops a fixed 180×120px shape at the caret (`editorApp.ts:1300`); Word/PPT let you drag-draw to size. Acceptable as a default, but should at least select-after-insert (A1).
- **No Format-Shape panel, no shape effects** (shadow/glow/3-D), no shape-style presets, no gradient/picture fill, no "Edit Points" for the custom-geometry path (custom geometry is render/round-trip only).
- **No group/ungroup, multi-select, duplicate, or align/distribute** (F5).
- **Two wrap styles surfaced** (in-line, square) vs Word's seven; tight/through/top-and-bottom aren't exposed, though behind/in-front exist via the layer control (C2).

None of these break round-trip; they're the difference between "edits DrawingML faithfully" and "authors like Word."

---

## What works well (so the report isn't all deficits)

- Lossless DrawingML round-trip is the foundation and it's solid; shapes deliberately reuse the image anchor/wrap/z-order model, so positioning behaves consistently with images and shares one stacking space.
- Rendering quality is excellent — gradients, rotation about center, text clipping, and grouped-child transforms all paint correctly on canvas and in PDF export.
- The rotate handle, resize handles, ghost-preview-then-single-commit protocol, Escape-to-cancel, and single-step undo per op are all well-engineered; the rotate handle even has a proper ARIA label.
- Fill/outline popovers reuse the editor's existing colour UI, so there's no bespoke colour widget to learn; the outline popover (colour + width + dash) is genuinely complete.
- Insert-a-text-box drops the caret inside for immediate typing — a nice Word-parity detail.

## Method notes & limitations

Live-driven with Playwright against the local dev build (`npm run dev`, feature-showcase document, 27 shapes present). Click-driven paths (ribbon, gallery, selection, floating toolbar, popovers, insert) were exercised directly; the "not selected after insert" and "inserts into the TOC" findings were reproduced live via the editor's `getSelectedShape()` API. Pointer-capture gestures (resize / rotate / drag-move / drag-select) can't be scripted and were assessed from the handle geometry, cursors, commit logic, and unit tests, per the constraint in `SHAPES_PLAN.md:272`. Severity reflects user impact; several "Major"s are individually small changes but collectively define whether shapes are usable by keyboard/AT users at all.
