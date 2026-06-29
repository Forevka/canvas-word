# Task: implement GitHub issue #54 end-to-end

Autonomous implementation agent in a dedicated git worktree, already on branch
`feat/issue-54-contextual-spacing`. Repo: Forevka/canvas-word (TS frontend + shared model + .NET ClearScript
bindings). Work only on THIS issue.

## 0. Setup
- Run `npm install` at the worktree root first (fresh worktree has no node_modules).

## 1. Read the issue
- `gh issue view 54` — implement exactly what it specifies, including its Definition-of-done checklist.

## 2. Definition of done (ALL required)
- Frontend round-trip: model (`shared/src/model/document.ts`), import (`frontend/src/import/docx/*`),
  export (`frontend/src/export/docx/*`), layout (`frontend/src/layout/engine.ts`), paint
  (`frontend/src/paint/renderer.ts` + PDF `frontend/src/export/pdf/paintBlock.ts`) as the feature needs.
- Default document: demonstrate the feature in `frontend/src/model/sampleDoc.ts` (the "no docId" showcase).
- TS DocumentBuilder: authorable via `frontend/src/builder/*`.
- C# bindings: mirror new builder methods in `dotnet/src/WordCanvas.ClearScript/Builder/`
  (`DocumentBuilder.cs`, `StoryBuilder.cs`, `Specs.cs`). `frontend/src/builder/csharpParity.test.ts` MUST pass.
- C# showcase: update `dotnet/examples/WordCanvas.Example.Showcase/Program.cs` to mirror the new `sampleDoc.ts`.
- Tests: docx round-trip (import -> export -> re-import) + layout/paint test where visual.
- CHANGELOG.md: entry under `[Unreleased]`.

## 3. Verify — all must pass (never finish red)
- `cd frontend && npx tsc --noEmit`
- `cd frontend && npx vitest run src/builder/csharpParity.test.ts src/export/docx src/import/docx src/layout` (+ files you add)
- `cd dotnet && dotnet build -c Release`

## 4. Keep merges easy (siblings edit the SAME shared files in parallel)
- Make ADDITIVE changes: new builder methods, a new clearly-labeled section appended in `sampleDoc.ts`
  and `Program.cs` near related content — do NOT restructure existing code. Keep the diff to issue #54 only.

## 5. Ship
- Commit with Conventional Commits. NO AI/Claude attribution anywhere.
- `git push -u origin HEAD`; `gh pr create` with body ending in `Closes #54`.

## 6. CodeRabbit review (required before you are done)
- After opening the PR, wait a few minutes for CodeRabbit's automated review.
- Fetch it: `gh pr view <PRNUM> --comments` and `gh api repos/Forevka/canvas-word/pulls/<PRNUM>/comments`.
- Address every actionable CodeRabbit finding (fix, commit, push); iterate until none remain (resolve or reply to each).
- Do NOT merge the PR yourself. When CodeRabbit feedback is resolved, print `READY-FOR-MERGE #<PRNUM>` and stop.
