# MergeReport — assemble a report from parts (C#)

The **"render parts, then fold them"** pattern: several sections are produced
independently and merged into one report, with explicit `MergeOptions` per part and
a per-section footer. Run it:

```sh
dotnet run -c Release --project dotnet/examples/WordCanvas.Example.MergeReport
# → out/merged-report.docx + out/merged-report.pdf
```

## What it shows

- **Generate parts in a loop.** Each section is built with the `DocumentBuilder`
  (in a real pipeline each might instead be rendered elsewhere and brought in with
  `engine.ImportDocx`). Every part deliberately reuses the same `"Heading 1"` style
  name and an `intro` bookmark, so the merge has real collisions to reconcile.
- **Fold with `MergeOptions`.** The parts are folded incrementally with
  `WordDocument.Append(part, options)`, where each part chooses its own
  `SectionBreak` (`NextPage` vs `Continuous`), styles reconcile via
  `StyleMergeMode.UseDestination`, and `RenameBookmarksOnCollision` keeps the
  duplicate `intro` bookmarks (`intro`, `intro__2`, …) instead of dropping them:

  ```csharp
  var report = parts[0];
  for (var i = 1; i < parts.Count; i++)
      report = report.Append(parts[i], new MergeOptions
      {
          Styles = StyleMergeMode.UseDestination,
          SectionBreak = specs[i].Break,          // per-part
          RenameBookmarksOnCollision = true,
      });
  ```

  …or fold them all at once: `engine.Merge(parts, new MergeOptions { … })`.
- **Per-section footers** via `SetSectionFooter`: a blank footer on the cover, and a
  branded content footer (a borderless `[ logo | address ]` table + a centered
  `Page X of Y`) on every other section.
- **`MergeOptions.Styles` side by side**: it prints the named-style count for
  `UseDestination` (duplicates collapsed by name) vs `KeepSource` (each part's styles
  kept, ids renamed).

> The ClearScript JS bundle is generated + gitignored — if you edit `entry.ts`/the
> bridge, run `node frontend/scripts/build-clearscript.mjs` before `dotnet run`.
