using WordCanvas.ClearScript;
using WordCanvas.ClearScript.Builder;

// =============================================================================
// Report assembly with merge — the "render parts, then fold them" pattern.
//
// Several sections are produced INDEPENDENTLY (here, built with the DocumentBuilder
// in a for-loop; in a real pipeline each might be rendered elsewhere and downloaded
// as a .docx, then brought in with engine.ImportDocx). They are then folded into one
// report with explicit MergeOptions per part, and given per-section footers.
//
// It exercises the whole merge surface: WordCanvas.Merge / WordDocument.Append,
// every MergeOptions field (styles mode, section-break kind, bookmark-collision
// handling), and SetSectionFooter for a per-section footer pass.
// =============================================================================

var outDir = Path.Combine(AppContext.BaseDirectory, "out");
Directory.CreateDirectory(outDir);

// A 1×1 blue PNG, scaled to a bar in the footer (kept inline so the example needs no assets).
byte[] logo = Convert.FromBase64String(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");

using var engine = new WordCanvasEngine();

// 1. The parts to assemble: a title + body, how each joins the previous section, and
//    whether it gets the branded footer (the cover page stays blank, like a real report).
var specs = new[]
{
    new SectionSpec("Cover", "Prepared for ACME Corp. — assembled from independently rendered parts.", MergeSectionBreak.NextPage, Footer: false),
    new SectionSpec("Executive Summary", "The subject is a well-located property. Lorem ipsum dolor sit amet.", MergeSectionBreak.NextPage, Footer: true),
    new SectionSpec("Comparable Analysis", "Three comparable sales support the concluded value.", MergeSectionBreak.NextPage, Footer: true),
    new SectionSpec("Appendix", "Supporting tables and definitions.", MergeSectionBreak.Continuous, Footer: true),
};

// 2. Render each part independently (the for-loop). Every part defines a style named
//    "Heading 1" and an `intro` bookmark ON PURPOSE — so the merge has to reconcile
//    the style names and rename the colliding bookmarks.
var parts = new List<WordDocument>();
foreach (var s in specs)
{
    var part = engine.NewBuilder(new CreateOptions { PageSize = PageSizeName.Letter })
        .Style(new NamedStyle { Id = "Heading1", Name = "Heading 1", Char = new CharStyle { Bold = true, FontSizePx = 24, Color = "#1a1a2e" }, Para = new ParaStylePatch { SpaceBeforePx = 16, SpaceAfterPx = 6 } })
        .Paragraph(s.Title, p => p.WithStyle("Heading1"))
        .Paragraph(s.Body)
        .Paragraph(p => p.Text("See ").Bookmark("intro", "the introduction").Text(" for context."))
        .Build();
    parts.Add(part);
    Console.WriteLine($"  rendered part '{s.Title}' ({part.BlockCount} blocks)");
}

// 3a. Fold them incrementally with WordDocument.Append, using per-part MergeOptions
//     (each part chooses its own section-break kind; same-named styles collapse onto
//     the destination's; colliding bookmarks are renamed rather than dropped).
var report = parts[0];
for (var i = 1; i < parts.Count; i++)
{
    var options = new MergeOptions
    {
        Styles = StyleMergeMode.UseDestination,        // "Heading 1" from each part → the destination's
        SectionBreak = specs[i].Break,                 // per-part: NextPage or Continuous
        RenameBookmarksOnCollision = true,             // intro → intro__2 → intro__3 …
    };
    report = report.Append(parts[i], options);
}

// 3b. …or fold them all at once with a single options set (equivalent when every part
//     shares the same break kind):
//        var report = engine.Merge(parts, new MergeOptions { SectionBreak = MergeSectionBreak.NextPage });

Console.WriteLine();
Console.WriteLine($"Merged report: {report.BlockCount} blocks, {report.GetSections().Count} sections, {report.MediaCount} images.");
foreach (var w in report.Warnings) Console.WriteLine($"  merge warning: {w.Code} — {w.Detail}");

// 4. Per-section footers: a blank footer on the cover, the branded content footer
//    (a borderless [logo | address] table + a centered "Page X of Y") everywhere else.
var editor = report.Edit();
var sectionCount = report.GetSections().Count;
for (var i = 0; i < sectionCount; i++)
{
    if (i == 0) editor.SetSectionFooter(i, f => f.Paragraph()); // cover: intentionally blank
    else editor.SetSectionFooter(i, f => BuildContentFooter(f, logo));
}
report = editor.ToDocument();

File.WriteAllBytes(Path.Combine(outDir, "merged-report.docx"), report.ExportDocx());
File.WriteAllBytes(Path.Combine(outDir, "merged-report.pdf"), report.ExportPdf());

// 5. MergeOptions.Styles side-by-side: useDestination collapses the duplicate "Heading 1"
//    onto one style; keepSource keeps each part's own definition (renaming the id).
var useDest = parts[0].Append(parts[1], new MergeOptions { Styles = StyleMergeMode.UseDestination });
var keepSrc = parts[0].Append(parts[1], new MergeOptions { Styles = StyleMergeMode.KeepSource });
Console.WriteLine();
Console.WriteLine("MergeOptions.Styles:");
Console.WriteLine($"  useDestination : {useDest.GetStyles().Count} named styles (duplicates collapsed by name)");
Console.WriteLine($"  keepSource     : {keepSrc.GetStyles().Count} named styles (source styles kept, ids renamed)");

// 6. Templating: open a template with a block-level content control, find it by tag,
//    and replace its whole content with a freshly rendered section — media, tables,
//    styles and nested controls all reconciled.
var template = engine.NewBuilder(new CreateOptions { PageSize = PageSizeName.Letter })
    .Style(new NamedStyle { Id = "Heading1", Name = "Heading 1", Char = new CharStyle { Bold = true, FontSizePx = 24 } })
    .Paragraph("Valuation Report", p => p.WithStyle("Heading1"))
    .ContentControlRange(cc => cc.Paragraph("[ BODY — replaced at fill time ]"), new SdtOptions { Tag = "BODY", Alias = "Body" })
    .Paragraph("— end of template —")
    .Build();

var freshSection = engine.NewBuilder(new CreateOptions { PageSize = PageSizeName.Letter })
    .Style(new NamedStyle { Id = "Heading1", Name = "Heading 1", Char = new CharStyle { Bold = true, FontSizePx = 24 } })
    .Paragraph("Filled Section", p => p.WithStyle("Heading1"))
    .Paragraph("This content was rendered separately and merged into the BODY control.")
    .Build();

var filled = template;
foreach (var sdt in template.GetSdtsByTag("BODY"))
    filled = filled.ReplaceSdtContent(sdt.Id, freshSection, new MergeOptions { Styles = StyleMergeMode.UseDestination });

File.WriteAllBytes(Path.Combine(outDir, "template-filled.docx"), filled.ExportDocx());
var filledBack = engine.ImportDocx(filled.ExportDocx());
Console.WriteLine();
Console.WriteLine("Template SDT fill (ReplaceSdtContent):");
Console.WriteLine($"  control 'BODY' filled → placeholder gone={filled.FindText("replaced at fill time").Count == 0}, section present={filledBack.FindText("Filled Section").Count > 0}");

Console.WriteLine();
Console.WriteLine($"Output written to: {outDir}");
return;

// A reusable branded content footer: a borderless [ logo | right-aligned address ]
// table over a centered "Page X of Y". The address wraps on its own and the page
// number is a live field resolved at layout (no placeholder / update-fields step).
static void BuildContentFooter(StoryBuilder f, byte[] logoPng)
{
    const string Font = "Times New Roman";
    const double SizePx = 16; // ~12pt
    var flush = new SpacingOptions { Before = 0, After = 0 };
    var borderless = new CellBorders();       // present, no edges ⇒ no lines
    var tightPad = new CellMargin(0, 2, 0, 2);

    f.Paragraph(p => p.Spacing(flush));       // leading spacer
    f.Table(t =>
    {
        t.ColFractions(0.28, 0.72);
        t.Row(r =>
        {
            r.Cell(c => c.Image(logoPng, "image/png", new ImageOptions { WidthPx = 64, HeightPx = 24, Align = TextAlign.Left }),
                new CellOptions { Borders = borderless, Margin = tightPad, VAlign = CellVAlign.Center, PreferredWidth = new PreferredWidth(28, WidthType.Pct) });
            r.Cell(c => c
                    .Paragraph("APR-2026-00042", p => p.Align(TextAlign.Right).Font(Font).FontSize(SizePx).Spacing(flush))
                    .Paragraph("123 Main St, Springfield, IL 62704", p => p.Align(TextAlign.Right).Font(Font).FontSize(SizePx).Spacing(flush)),
                new CellOptions { Borders = borderless, Margin = tightPad, VAlign = CellVAlign.Center, PreferredWidth = new PreferredWidth(72, WidthType.Pct) });
        });
    }, new TableOptions { WidthMode = TableWidthMode.AutofitWindow });
    f.Paragraph(p => p
        .Align(TextAlign.Center).Font(Font).FontSize(SizePx).Spacing(new SpacingOptions { Before = 7, After = 0 })
        .Text("Page ").PageField().Text(" of ").NumPagesField());
}

// A part to render + fold into the report.
internal readonly record struct SectionSpec(string Title, string Body, MergeSectionBreak Break, bool Footer);
