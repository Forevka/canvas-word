# SDLC — how a change ships in canvas-word

The repeatable workflow every feature/fix follows here, plus the **Definition of
Done** that decides when it's actually finished. If you're an agent or a new
contributor picking up a task, this is the loop to run. It exists because the
codebase has a wide, layered surface (TS model ↔ layout ↔ import/export ↔ npm
builder ↔ C# bindings) and a "done" change usually has to land in **all** the
layers a feature touches — not just the one where the work started.

Companion docs: [ARCHITECTURE.md](./ARCHITECTURE.md) (how the layers fit),
[OOXML_COVERAGE.md](./OOXML_COVERAGE.md) (element-by-element import/export map),
[BUILDER.md](./BUILDER.md), [IMPORT.md](./IMPORT.md), [EXPORT.md](./EXPORT.md),
[REVIEW.md](./REVIEW.md).

---

## The loop

1. **Understand** — read the issue / request. Restate the scope; note anything
   deferred to a separate issue (file those first, with a real description).
2. **Explore before editing.** Find the *parallel* that already exists — almost
   every feature is "like an existing one." Use CodeGraph (`codegraph_explore`)
   for "how does the code connect" and `fff` (`mcp__fff__grep`) for "where does
   this text live." Read the analogous op / command / bar / round-trip and mirror
   its shape. A new op mirrors a sibling op; a new context bar mirrors an existing
   bar; a new round-tripped field mirrors a sibling field's export + import + IR.
3. **Plan.** Resolve the design forks with sensible defaults (don't ask the user
   about anything the code or a convention already answers). Decide the model
   shape, the single layout lever, and the OOXML representation up front. List the
   DoD surfaces the change must touch (below).
4. **Branch.** `feat/<slug>` (or `fix/`, `chore/`, `docs/`, `ci/`). Never commit
   feature work straight to `main`.
5. **Implement in dependency order** so each layer compiles on the one below:
   **model → op (+ inverse) → command → layout → UI/editor wiring →
   import/export → builder → C# parity → tests → docs.** (Skip layers a change
   genuinely doesn't touch — an editor-only bar needs no OOXML work.)
6. **Verify** — see the gates below. All green before the PR, not after.
7. **Document** — CHANGELOG + any coverage/README/example that the change makes
   stale, in the *same* commit as the code.
8. **PR to `main`** with a detailed body (what/why/how, tests, verification,
   `Closes #NNN`).
9. **Review loop** — let CodeRabbit review; address every actionable finding, and
   apply the worthwhile nitpicks (normalization, dedup) too. Push fixes as
   follow-up commits.
10. **Merge** — squash-merge once CI is green and findings are resolved; delete
    the branch; `git checkout main && git pull --ff-only`.
11. **Remember** — update the memory file for the subsystem and close/annotate the
    tracking issue.

---

## Definition of Done (DoD)

A feature isn't done when it works in the editor — it's done when it survives a
round-trip and is reachable from **every** surface it belongs to. For a change to
the document model, that means landing in all of:

- **Frontend round-trip** — model → `.docx` export → import → model, asserted by a
  test. New OOXML mappings update **[OOXML_COVERAGE.md](./OOXML_COVERAGE.md)** in
  the same PR.
- **`sampleDoc.ts`** — the flagship in-editor document demonstrates the feature so
  it's visible by default (and exercised by the census/round-trip tests).
- **`DocumentBuilder`** — the fluent TS composer (`frontend/src/builder/`) can
  author it; add options to the relevant `*Options` and the hand-written
  `frontend/types/*.d.ts` surface.
- **C# bindings** — the ClearScript wrapper
  (`dotnet/src/WordCanvas.ClearScript/Builder/`) mirrors the JS builder. The
  `csharpParity` test guards method-name drift; add the option to the C# `*Options`
  record + its `ToJs`.
- **C# showcase** — `dotnet/examples/.../Program.cs` exercises it, so the CI OOXML
  schema validation covers the re-exported document.

**Intentionally out of scope** (don't retrofit these unless asked): Comments /
tracked changes as in-stream OOXML (they live in the review overlay — see
[REVIEW.md](./REVIEW.md)), OLE objects, and non-image drawing shapes.

---

## Verification gates

Run from the affected package (`frontend/` and/or `shared/`). All must pass before
opening the PR:

| Gate | Command | Covers |
| --- | --- | --- |
| Types | `npm run typecheck` | `tsc --noEmit`, both packages |
| Unit / integration | `npm run test` (or `npx vitest run <file>` while iterating) | model ops, commands, layout, round-trips |
| Editor build | `npm run build` | Vite app build |
| Library bundle | `npm run build:lib` | the published `@forevka/wordcanvas` bundle — rebuild after any `shared/` or entry change |
| C# | `dotnet build -c Release` (ClearScript lib + showcase) | binding + showcase compile |
| OOXML schema | CI `ooxml-validate` (`.github/workflows/ci.yml`) | re-exported showcase validates against the OOXML schema with the Open XML SDK — the lenient importer is **not** enough |
| Browser | Playwright against an `examples/` page | selection/caret/click-driven UI paths |

**Author and verify are separate passes.** Don't self-approve in the same breath
as writing — the CodeRabbit review is the approval lane, and CI is the evidence.

---

## Conventions

- **Windows shell.** This is a Windows machine — use the **PowerShell** tool for
  anything that touches a path (`git`, `dotnet`, `npm`). Git Bash treats `\` as an
  escape, so `C:\...` paths silently break there. Read/Write/Edit/Glob/Grep take
  Windows paths natively.
- **No AI attribution.** No `Co-Authored-By: Claude`, no "Generated with…" lines in
  commits, PRs, issues, comments, or code. Write as the author.
- **CHANGELOG.** Every user-visible change gets an `### Added` / `### Changed` /
  `### Fixed` entry under `## [Unreleased]`, in the same PR.
- **Commit style.** Conventional prefixes (`feat(editor):`, `fix(export):`,
  `refactor(...)`, `chore(...)`, `docs(...)`). Follow-up commits addressing review
  say so (`refactor(...): address CodeRabbit — …`).
- **PR body.** What + why + how, a tests/verification section, and `Closes #NNN`.
  Be honest about what was and wasn't verified (see constraints below).
- **Squash-merge** to `main`; the merge commit carries the `(#NNN)` suffix.
- **Docs-only changes commit straight to `main`** — no branch, no PR, no review. The
  branch → PR → CodeRabbit loop is for code; pure prose (`*.md`) doesn't need it. A
  change that mixes code + docs still goes through a PR (the code drives that).
- **Deferred work is a real issue.** Split from-scratch sub-features into their own
  tracked issues with a proper description before shipping the part that's ready.

---

## Known harness constraints (state these honestly in PRs)

- **Playwright can't drive the editor's drags.** The canvas uses pointer-capture
  drag selection and resize handles use `setPointerCapture`, which rejects
  synthetic pointer IDs — so **drag-select** and **drag-to-resize commit** paths
  can't be driven from scripted events. Verify these by their commit *logic* (unit
  tests on the op/command) and by structural/visual checks (handles appear;
  `setDocument` with the target state renders correctly). Click- and caret-driven
  UI (context bars, menus, buttons) *is* drivable and should be browser-verified.
- **CodeRabbit incremental re-review is rate-limited.** After a fix push it often
  reports "Review rate limited." The *first* full review still counts: if it
  completed, its findings are addressed, all real CI (test / ooxml-validate /
  GitGuardian) is green, and no threads are open, merging is fine — note it in the
  PR.

---

## Where each layer lives

| Layer | Path |
| --- | --- |
| Document model + ops | `shared/src/model/` (`document.ts`, `ops.ts`, `query.ts`) |
| Editor commands | `frontend/src/editor/commands.ts` |
| Editor host / wiring | `frontend/src/index.ts`, `frontend/src/editorApp.ts` |
| Layout / geometry | `frontend/src/layout/` (`engine.ts`, `geometry.ts`, `math/`) |
| UI (bars, dialogs, menus) | `frontend/src/ui/` |
| `.docx` import | `frontend/src/import/docx/` (`documentParser.ts` → IR → `mapToModel.ts`) |
| `.docx` / PDF export | `frontend/src/export/` (`docx/documentXml.ts`, `pdf/`) |
| Builder | `frontend/src/builder/` + hand-written `frontend/types/*.d.ts` |
| C# bindings | `dotnet/src/WordCanvas.ClearScript/` |
| Sample / showcase | `frontend/src/model/sampleDoc.ts`, `dotnet/examples/.../Program.cs` |
