# Drawing Shapes — implementation plan (issue #206)

A proper, **full-OOXML-round-trip** `ShapeBlock` feature (the lossy `CustomBlock`
alternative is explicitly rejected). Structured as a sequence of independently
shippable, DoD-complete PRs — each one a vertical slice that survives
model → `.docx` → import → model and validates against the OOXML schema.

Follows [SDLC.md](./SDLC.md): implement in dependency order (**model → op →
command → layout → UI → import/export → builder → C# parity → tests → docs**),
and land in **every** DoD surface a model change touches.

Every shape feature must be reachable from **both** the floating context popup
**and** the ribbon (and, where an image analog exists, the right-click menu). The
default in-editor document (`sampleDoc.ts`, loaded when no `docId` is passed) must
demonstrate **full coverage** of the feature — every preset, fill/stroke variant,
text box, wrap mode, anchor mode, and z-order overlap.

> **SDLC.md itself must change.** Its Definition of Done currently lists
> "non-image drawing shapes" under *Intentionally out of scope*. PR 0 removes that
> line — this feature reverses that decision.

---

## The parallel we mirror: `ImageBlock`

Shapes are "like an image, but the picture is a vector geometry we draw." Almost
every layer already has the exact shape of code we need, keyed to `ImageBlock`.
The table below is the master mirror map (verified against the current tree).

| Layer | Image reference (mirror this) | New shape symbol |
| --- | --- | --- |
| Model (authoritative) | `shared/src/model/document.ts` — `interface ImageBlock` (L468–529), `Block` union (L777) | `interface ShapeBlock` + `\| ShapeBlock` |
| Model (public d.ts) | `frontend/types/model.d.ts` — `ImageBlock` (L121–134), `Block` (L252) | `ShapeBlock` + union |
| Op patch + reducer | `shared/src/model/ops.ts` — `ImagePropsPatch` (L43–51), `setImageProps` op + reducer (L766–813, inverse-symmetric) | `ShapePropsPatch`, `setShapeProps` |
| Insert/remove | `insertBlock`/`removeBlock` (already generic — reused as-is) | — |
| Commands | `frontend/src/editor/commands.ts` — `locateImage`/`ImageLocation` (L414–430), `insertImage` (L1815), `setImageProps` (L2468), `bringImageToFront`/`sendImageToBack` (L2539/2549), `moveAnchoredImage` (L2561), `deleteImage` (L2606) | `locateShape`, `insertShape`, `setShapeProps`, `bringShapeToFront`/`sendShapeToBack`, `moveAnchoredShape`, `deleteShape` |
| Layout | `frontend/src/layout/engine.ts` — measured union (L1210), `placeImage`/`placeAnchoredImage` (L1605–1653), z-order (L2248); `layoutTree.ts` `PlacedImage` | `placeShape`/`placeAnchoredShape`, `PlacedShape` |
| Paint (canvas) | `frontend/src/paint/renderer.ts` — image draw branch (L1106–1132) | shape draw branch (geometry paths) |
| Paint (PDF) | `frontend/src/export/pdf/paintBlock.ts` — image mirror | shape mirror |
| Hit-test / object rect | `frontend/src/layout/geometry.ts` — `objectRect` (L861), `hitTestSelectableObject` (L842), `scanImages` | shape scan added to both |
| Selected-object state | `frontend/src/index.ts` — `selectedObject` (L1146), `selectObject` (L1243), `selectedIsImage` (L1233), `getSelectedObjectRect` (L3720) | `selectedIsShape`, `getSelectedShapeRect`, `CurrentFormat.shapeSelected` (L209 area) |
| Resize frame | `frontend/src/input/objectController.ts` — `createObjectFrame` / 8 `HANDLES` (L114) / `onResizeCommit` (already generic) | resize commit → `setShapeProps({widthPx,heightPx})` |
| Import | `frontend/src/import/docx/documentParser.ts` — `parseDrawing` (L694), `images-skipped` warning (L705), VML warn (L801); `types.ts` IR; `mapToModel.ts` `mapImage` (L729) | parse `wps:wsp`; IR `kind:"shape"`; `mapShape` |
| Export | `frontend/src/export/docx/documentXml.ts` — `imageParagraphXml` (L888), `drawingIdAttrs` (L433), `wp:inline`/`wp:anchor` wrappers | `shapeParagraphXml` (reuses the same wrappers) |
| Builder (JS) | `frontend/src/builder/storyBuilder.ts` — `ImageOptions` (L14–30), `.image()` (L106); `blockFactory.ts` `image()` (L125); `frontend/types/builder.d.ts` | `ShapeOptions`, `.shape()`, `blockFactory.shape()` |
| C# binding | `dotnet/.../Builder/StoryBuilder.cs` — `Image()` (L87); `Specs.cs` `ImageOptions` record + `ToJs` (L515), `enum ImageWrap` (L39) | `Shape()`, `ShapeOptions`, `enum ShapePreset` |
| Parity guard | `frontend/src/builder/csharpParity.test.ts` (enforces JS↔C# method parity) | `Shape()` must appear both sides |
| Floating popup | `frontend/src/ui/imageContextToolbar.ts` (`priority:30`); framework `ui/contextToolbar.ts` (`pickActive`, `createContextToolbarManager`); registered `editorApp.ts:3278` | `ui/shapeContextToolbar.ts` (`priority:31`) |
| Ribbon insert | `editorApp.ts` — Illustrations group (L1705), `pickAndInsertImage`+button (L1710–1745); grid-picker pattern `tableGridPopover` (L1120) triggered by caret button (L1703) | Shapes gallery button + `insertShape` |
| Ribbon contextual | `editorApp.ts` — "Picture" group (L1746) with `enable(btn, f => f.imageSelected)` (L877 / sync L2808) | "Shape" group with `enable(btn, f => f.shapeSelected)` |
| Right-click menu | `frontend/src/index.ts` — image menu (L2938–2969): Wrap / Align / Bring-to-Front / Send-to-Back / Delete | shape submenu |
| Sample / showcase | `frontend/src/model/sampleDoc.ts` — `image()` helper (L71), images section (L476–557); `dotnet/examples/.../Showcase/Program.cs` (L246–341) | `shape()` helper + "Drawing shapes" section |
| Coverage doc | `docs/OOXML_COVERAGE.md` | drawings: dropped → supported |

---

## Design decisions (forks resolved up front)

**Model — `ShapeBlock`.** Block-level entry in the block list (inserted by
splitting a paragraph, exactly like an image), with the image positioning surface
reused verbatim:

```ts
interface ShapeBlock {
  kind: "shape";
  id: string;
  revision: number;
  geometry: { preset: ShapePreset; adjust?: Record<string, number> }; // a:prstGeom prst + a:avLst
  fill?: { color: string } | { none: true };        // hex; undefined ⇒ theme-neutral default
  stroke?: { color: string; widthPt: number; dash?: ShapeDash; none?: true };
  widthPx: number;                                    // the single layout lever (box), like image
  heightPx: number;
  align: "left" | "center" | "right";
  wrap?: "block" | "square";                          // mirror ImageBlock exactly
  anchor?: ImageBlock["anchor"];                       // reuse the wp:anchor float+z shape verbatim
  rotation?: number;                                   // a:xfrm@rot degrees (PR 2)
  text?: ShapeTextBody;                                // wps:txbx body (PR 3)
  drawingId?: { anchorId?: string; editId?: string };  // wp14 identity (reuse image plumbing)
  fieldId?: string; sdtPath?: string[];                // membership, mirror image
}
type ShapePreset = "rect" | "roundRect" | "ellipse" | "triangle" | "diamond"
                 | "rightArrow" | "leftArrow" | "line";                 // ⊇ grows in PR 2
```

The public `frontend/types/model.d.ts` `ShapeBlock` mirrors `ImageBlock`'s slimmer
public surface (geometry/fill/stroke/size/align/wrap — omit the internal
`drawingId`/`sdtPath`), matching the existing image asymmetry.

**Single layout lever** = the `widthPx × heightPx` box; the `preset` selects which
vector path is drawn inside it. No per-vertex model — freeform `a:custGeom` is a
deferred follow-up.

**OOXML representation (export).** Reuse the image `wp:inline` / `wp:anchor`
wrappers + `drawingIdAttrs`, swap the graphic payload:

```
wp:inline|wp:anchor
 └ a:graphic / a:graphicData uri="…/wordprocessingShape"
    └ wps:wsp
       ├ wps:cNvSpPr
       ├ wps:spPr
       │  ├ a:xfrm (a:off, a:ext[, rot])
       │  ├ a:prstGeom prst=…  ( + a:avLst )
       │  ├ a:solidFill | a:noFill                 ← fill
       │  └ a:ln w=…  ( a:solidFill + a:prstDash )  ← stroke
       ├ wps:txbx / w:txbxContent (paragraphs)      ← text body (PR 3)
       └ wps:bodyPr
```

Export emits **bare DrawingML `wps`** (no `mc:AlternateContent`/VML) — modern Word
reads it and it validates against the OOXML schema (the `ooxml-validate` CI gate).
VML is **import-only** (legacy read), see PR 5.

**Import.** Extend `parseDrawing` to branch on `a:graphicData@uri === wps` → parse
`wps:wsp` into an IR `{ kind:"shape", … }`; `mapShape` converts EMU→px and reuses
the image anchor/wrap mapping. The `images-skipped` warning (L705) is narrowed so
recognized shapes no longer trip it; genuinely unknown drawings still warn.

**Text body.** `ShapeTextBody = { blocks: Paragraph[] }` — a nested paragraph flow
like a table cell, so `wps:txbx ↔ w:txbxContent` maps 1:1 and layout can reuse
cell-style fixed-width sub-flow. Read-only render + round-trip lands in PR 3;
**editable** caret-in-box editing is its own PR (PR 6) because it needs the
selection/input layer to route a caret into a nested sub-flow.

**Fill/outline pickers** reuse the existing color-popover already used by font
color / highlight (no new color UI).

**Selection.** `selectedObject` is already a generic block id — add `locateShape` +
`selectedIsShape()`, a `getSelectedShapeRect()` alongside the image/equation
variants, and a `CurrentFormat.shapeSelected` flag so ribbon `enable(...)`
predicates can gate the contextual Shape group.

---

## PR sequence

Each PR passes all [verification gates](./SDLC.md#verification-gates) before its PR
is opened, extends `sampleDoc.ts` + the C# showcase for what it adds (DoD), and
updates `OOXML_COVERAGE.md` + `CHANGELOG.md` in the same PR.

### PR 0 — docs + sub-issues (straight to `main`, no branch/PR)
- `docs/SHAPES_PLAN.md` (this file).
- Edit `docs/SDLC.md` DoD: remove "and non-image drawing shapes" from *Intentionally
  out of scope*.
- Create the **phase sub-issues (PR 1–9)** under parent #206, each carrying its own
  self-contained implementation brief (the mirror map for that phase) so a subagent
  can pick it up cold. Add a task-list to #206 linking them. **Nothing is deferred —
  the full feature (incl. editable text boxes, freeform geometry, grouped shapes) is
  in scope.**

### PR 1 — Foundation vertical slice `feat/shape-block`
The "new `Block` kind" PR — inherently the largest because DoD requires the whole
stack for anything touching the model. Kept narrow on breadth: presets **rect,
ellipse, line**; **solid fill + solid stroke**; **in-flow (block) only**.
- Model (`document.ts` + `model.d.ts`), op (`setShapeProps` + reducer + inverse),
  commands (`locateShape`/`insertShape`/`setShapeProps`/`deleteShape`).
- Layout `placeShape` (block flow) + `PlacedShape`; canvas painter (rect/ellipse/line
  paths, fill+stroke) + PDF mirror.
- Selection + drag-resize wired through `objectController` → `setShapeProps`.
- Import (`wps:wsp` rect/ellipse/line) + export (`shapeParagraphXml`), narrowing the
  `images-skipped` warning.
- Builder `.shape()` + `ShapeOptions` + `builder.d.ts`; C# `Shape()` + `ShapeOptions`
  + `enum ShapePreset` + `csharpParity`.
- **Both surfaces:** ribbon Insert→Illustrations **Shapes** gallery (grid picker,
  `tableGridPopover` pattern) → `insertShape`; contextual **Shape** ribbon group
  (Delete) + floating `shapeContextToolbar` (Delete) — skeletons that PR 2 fills in.
- `sampleDoc.ts` "Drawing shapes" section (rect/ellipse/line) + showcase.
- Round-trip test + `ooxml-validate` + `OOXML_COVERAGE.md` + `CHANGELOG.md`.

### PR 2 — Geometry & style breadth + fill/outline UI `feat/shape-styles`
- Presets: roundRect, triangle, diamond, right/left arrow (+ `a:avLst` adjust
  passthrough); `rotation` (`a:xfrm@rot`).
- Stroke: width, dash styles, `none`; fill: `none`/solid. Painter path table +
  import/export prst/dash maps; builder + C# enums grow.
- **Both surfaces:** Fill + Outline controls in the floating popup **and** the
  contextual ribbon Shape group (reusing the existing color popover); shape-preset
  gallery filled out.
- `sampleDoc.ts` gallery of every preset + stroke/fill variants; showcase mirror.

### PR 3 — Text boxes (read-only) `feat/shape-textbox`
- `ShapeTextBody` model; import `wps:txbx`→`w:txbxContent`, export same; layout
  sub-flow inside the box; render text.
- Builder/C# `text` option; sample text box + showcase.
- Editable caret-in-box deferred to its filed follow-up issue.

### PR 4 — Positioning parity `feat/shape-anchor-wrap`
- Square wrap + absolute `wp:anchor` float + z-order, reusing the image anchor model
  and `anchorZRange`/`anchorFor`.
- Commands `bringShapeToFront`/`sendShapeToBack`/`moveAnchoredShape`; `placeAnchoredShape`.
- **All three surfaces:** wrap + order in floating popup, ribbon Shape group, and
  right-click shape menu (mirroring the image menu).
- `sampleDoc.ts` overlap/wrap/behind-vs-front demo; showcase mirror.

### PR 5 — VML fallback import `feat/shape-vml-import`
- Parse legacy `w:pict` / `v:rect` / `v:shape` / `v:textbox` → IR shape (read-only;
  export stays DrawingML). Relax the legacy VML warning (L801) for recognized shapes.
- Round-trip fixtures from a real legacy `.docx`.

### PR 6 — Editable text boxes `feat/shape-textbox-edit`
Promotes PR 3's read-only text body to a fully editable nested flow.
- Route the caret into the `ShapeTextBody` sub-flow: extend `selectionController` /
  hit-testing so a click inside a text box lands a caret in the nested paragraphs;
  editing ops target the nested block path (mirror how table-cell content is edited).
- Enter to double-click-enter the box; Escape to pop back to object selection.
- Round-trip: edited text persists through `wps:txbx`. Sample text box becomes
  editable; showcase mirror.

### PR 7 — Freeform / custom geometry `feat/shape-custgeom`
- Model `geometry` gains a `custom` variant carrying a normalized path
  (`a:custGeom` → `a:path` move/line/cubic/close in a 0–1 or EMU coord space).
- Painter renders the path; import parses `a:custGeom`; export emits it. Preset
  shapes remain the common case; custom is the escape hatch.
- Builder/C# `.shape()` accepts a path; sample adds a freeform shape; showcase mirror.

### PR 8 — Grouped shapes `feat/shape-group`
- New `ShapeGroupBlock` (or `geometry: { group }` container) mapping `wpg:wgp` —
  child shapes with a group transform (`a:xfrm` + `a:chOff`/`a:chExt` child coord
  space). Selection selects the group; children move/scale with it.
- Layout composes children under the group transform; import parses `wpg:wgp`, export
  emits it. Depends on PR 4 (anchoring) for group positioning. Sample + showcase.

### PR 9 — Full-surface polish & default-doc completeness `feat/shape-polish`
- Complete right-click shape menu; keyboard (Delete / Escape / arrow-nudge for
  anchored; group enter/exit); a11y labels; audit that **every** shape feature is
  reachable via the floating popup **and** the ribbon **and** the menu.
- Finalize the `sampleDoc.ts` "Drawing shapes" section to **full coverage** (every
  preset + a custom-geometry shape, fill/stroke variants, an editable text box, all
  wrap+anchor modes, a z-order overlap, and a grouped shape) and the C# showcase.
- Docs pass: `OOXML_COVERAGE.md`, `README.md`, `docs/RIBBON.md` shapes entry.

---

## Dependency graph & integration order

```
                 ┌─► PR2 (styles) ─────► PR7 (custGeom) ─┐
PR1 (foundation) ┼─► PR3 (textbox RO) ─► PR6 (textbox ed)┤
                 ├─► PR4 (anchor/wrap) ─► PR8 (group) ────┼─► PR9 (polish)
                 └─► PR5 (VML import, after PR2+PR3) ─────┘
```

- **PR 1 gates everything** — it introduces the `ShapeBlock` kind; nothing else can
  round-trip until it lands. Build and merge it first, alone.
- **After PR 1 merges**, PR 2 / PR 3 / PR 4 fan out in parallel (one worktree each).
- **Second wave**: PR 5 (needs PR 2 geometry + PR 3 text body for VML textboxes),
  PR 6 (needs PR 3), PR 7 (needs PR 2), PR 8 (needs PR 4).
- **PR 9 last**, after all merge.

**Parallel development, serialized integration.** Every PR touches a few shared
files (`document.ts`, `ops.ts`, `documentXml.ts`, `documentParser.ts`,
`commands.ts`, `sampleDoc.ts`). Agents work in **isolated git worktrees** so they
never collide on disk, but the **orchestrator merges one at a time** and each open
branch rebases onto `main` after the prior merge. This keeps CI honest (every merge
is validated on the real integrated tree) and confines conflict resolution to the
rebasing agent.

## Orchestration (herdr)

- One **orchestrator** pane (this one) owns the plan, the sub-issues, the merge
  queue, and the rebases.
- Each phase runs as a **claude agent in its own herdr pane + git worktree**
  (`git worktree add ../cw-shape-<slug> -b feat/shape-<slug>`), started in **auto
  mode**, handed only its sub-issue number and told to follow SDLC.md end-to-end
  (branch → implement in dependency order → all gates green → PR with `Closes #NNN`).
- The orchestrator monitors via `herdr wait agent-status … --status done` +
  `pane read`, verifies each PR's gates, squash-merges in dependency order, then
  releases the next wave.

## Testing & known constraints
- **Round-trip** unit tests per PR (model → export → import → model equality on the
  new fields).
- **Playwright** cannot drive pointer-capture drags (resize/drag-move commit) — verify
  those by the op/command unit tests + structural checks (handles appear;
  `setDocument` with the target state renders). Click-driven paths — the **ribbon
  buttons, gallery picker, and floating-popup buttons** — *are* drivable and get
  browser-verified.
- **`ooxml-validate`** (Open XML SDK) is the real export gate — the lenient importer
  is not enough; every PR that emits new XML extends the C# showcase so CI validates it.

## Scope
Full DrawingML shape support with lossless `.docx` round-trip. **Nothing is
deferred** — preset + custom geometry, solid/no fill, stroke styles, rotation,
editable text boxes, wrap/anchor/z-order positioning, grouped shapes, and legacy VML
import are all in scope across PRs 1–9. The lossy `CustomBlock` path is explicitly
rejected.
