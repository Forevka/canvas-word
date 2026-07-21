using System.IO.Compression;
using WordCanvas.ClearScript;
using WordCanvas.ClearScript.Builder;

// =============================================================================
// Showcase: rebuild the editor's DEFAULT document (the "no docId" sample) with the
// typed C# DocumentBuilder, then export it to PDF + DOCX. It exercises most of the
// builder surface — headings + TOC, inline fields (DATE/PAGE/NUMPAGES/IF), every
// content-control kind, merged / tall (paginating) / field-in-cell tables, images,
// multilevel + bullet lists, footnotes, bookmarks, hidden text, hyperlinks, sub/
// superscript, CJK text, display + inline equations (LaTeX & MathML), right-to-left
// paragraphs, a registered real table style, and header/footer page fields.
// =============================================================================

const string Ink = "#1a1a2e";
const string Lorem = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod " +
    "tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud " +
    "exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ";

var blue = SolidPng(360, 110, 0x1a, 0x73, 0xe8);
var purple = SolidPng(150, 110, 0x9c, 0x27, 0xb0);
// A non-uniform tile so the image-crop (a:srcRect) demo shows a visible difference.
var quad = QuadrantPng(150, 110);

using var engine = new WordCanvasEngine();

var builder = engine.NewBuilder(new CreateOptions { PageSize = PageSizeName.Letter })
    // Heading styles so withStyle() resolves AND the TOC detects them by name.
    .Style(new NamedStyle { Id = "Heading1", Name = "Heading 1", Char = new CharStyle { Bold = true, FontSizePx = 24, Color = Ink }, Para = new ParaStylePatch { SpaceBeforePx = 18, SpaceAfterPx = 8 } })
    .Style(new NamedStyle { Id = "Heading2", Name = "Heading 2", Char = new CharStyle { Bold = true, FontSizePx = 19, Color = Ink }, Para = new ParaStylePatch { SpaceBeforePx = 14, SpaceAfterPx = 6 } })
    // A shaded paragraph style — used below to demonstrate an explicit "No Color" clear (issue #147).
    .Style(new NamedStyle { Id = "Callout", Name = "Callout", BasedOn = "Normal", Para = new ParaStylePatch { Shading = "#d9ead3", SpaceBeforePx = 8, SpaceAfterPx = 8 } })

    // A REAL table style (conditional bands) — registered once, referenced by id,
    // baked onto the cells AND emitted as w:tblStyle so it survives the docx round-trip.
    .TableStyle(new TableStyleDef
    {
        Id = "ShowcaseGrid",
        Name = "Showcase Grid",
        Conds = new Dictionary<TableCond, TableCondProps>
        {
            [TableCond.WholeTable] = new() { Borders = CellBorders.All(new CellBorder { Color = "#c8ccd0", WidthPx = 1 }) },
            [TableCond.FirstRow] = new() { Shading = Ink, Char = new CharStyle { Bold = true, Color = "#ffffff" } },
            [TableCond.Band2Horz] = new() { Shading = "#f1f3f4" },
        },
    })

    // Document-level default tab interval (w:defaultTabStop) — 0.75in instead of the
    // engine's 0.5in fallback (#63); the tab-stop demo below visibly honors it.
    .DefaultTabStop(72)

    // Title + subtitle + Table of Contents.
    .Paragraph("canvas-word", p => p.Align(TextAlign.Center).Font("Arial, sans-serif").FontSize(32).Bold().Color(Ink))
    .Paragraph("a canvas-rendered, page-accurate Word editor — rebuilt by the C# builder", p => p.Align(TextAlign.Center).Italic().Color("#5f6368"))
    .Paragraph("Table of Contents", p => p.Bold().FontSize(20).Color(Ink))
    .TableOfContents(new TocOptions { MaxLevel = 2 })

    // ---- Fields ----
    .Paragraph("Fields", p => p.WithStyle("Heading1"))
    .Paragraph(p => p
        .Text("Fields are first-class objects you can insert and edit. Today is ")
        .DateField("MMMM d, yyyy")
        .Text(", this is page ").PageField()
        .Text(". A conditional (IF) field can branch: ")
        .IfField("2", ">", "1", "the condition held", "it did not")
        .Text(". Each recomputes on Update Field."))
    .Paragraph(p => p
        .Text("Footnotes", new CharStyle { Bold = true }).Text(" are supported")
        .Footnote("Footnotes lay out at the bottom of their page, with a separator rule — just like Word.")
        .Text(". So is ").Text("hidden metadata", new CharStyle { Hidden = true })
        .Bookmark("sample", "bookmarked text")
        .Text(" and inline formatting: ")
        .Text("bold", new CharStyle { Bold = true }).Text(", ")
        .Text("italic", new CharStyle { Italic = true }).Text(", ")
        .Text("underline", new CharStyle { Underline = true }).Text(", ")
        .Text("strike", new CharStyle { Strikethrough = true }).Text(", ")
        .Text("highlight", new CharStyle { HighlightColor = "#fff3a3" }).Text(", x")
        .Text("2", new CharStyle { VerticalAlign = "super", FontSizePx = 11 })
        .Text(", and a ")
        .Text("hyperlink", new CharStyle { Link = "https://forevka.dev", Color = "#0b57d0", Underline = true })
        .Text("."))
    // Underline styles + colors (w:u val + color) — double/dotted/dashed/wave/thick.
    .Paragraph(p => p
        .Text("Underlines carry a ").Text("style", new CharStyle { Italic = true })
        .Text(" and an optional ").Text("color", new CharStyle { Italic = true }).Text(": ")
        .Text("double", new CharStyle { Underline = true, UnderlineStyle = "double" }).Text(", ")
        .Text("dotted", new CharStyle { Underline = true, UnderlineStyle = "dotted" }).Text(", ")
        .Text("dashed", new CharStyle { Underline = true, UnderlineStyle = "dash" }).Text(", ")
        .Text("dot-dash", new CharStyle { Underline = true, UnderlineStyle = "dotDash" }).Text(", ")
        .Text("thick", new CharStyle { Underline = true, UnderlineStyle = "thick" }).Text(", ")
        .Text("a red wavy", new CharStyle { Underline = true, UnderlineStyle = "wave", UnderlineColor = "#d93025" }).Text(", and ")
        .Text("a blue double", new CharStyle { Underline = true, UnderlineStyle = "double", UnderlineColor = "#1a73e8" })
        .Text(" — each round-trips through Word's w:u (style + color)."))
    // Case transforms (w:caps / w:smallCaps): the model text stays as typed; only the
    // rendered glyphs are uppercased (small caps additionally shrinks the lowercase ones).
    .Paragraph(p => p
        .Text("Case transforms: ")
        .Text("all caps", new CharStyle { Caps = true }).Text(" (w:caps) and ")
        .Text("Small Caps", new CharStyle { SmallCaps = true }).Text(" (w:smallCaps) — both render UPPERCASED, ")
        .Text("while the underlying text stays exactly as authored."))
    // Endnotes (w:endnoteReference) — like footnotes, but collected at the document end.
    .Paragraph(p => p
        .Text("Endnotes", new CharStyle { Bold = true }).Text(" round-trip too")
        .Endnote("Endnotes collect at the very end of the document, under their own separator rule — Word's “end of document” placement.")
        .Text(" — like footnotes, but parked at the document end."))

    // Minor run typography & effects (w:rPr extras) — via the Effects() binding.
    .Paragraph(p => p
        .Text("Double-struck text (w:dstrike) round-trips and paints two rules.")
        .Effects(new RunEffectsOptions { DoubleStrikethrough = true }))
    .Paragraph(p => p
        .Text("Raised baseline via w:position — distinct from superscript (same size).")
        .Effects(new RunEffectsOptions { PositionPx = 4 }))
    .Paragraph(p => p
        .Text("W I D E   character width scaling (w:w) stretches each glyph horizontally.")
        .Effects(new RunEffectsOptions { WidthScalePct = 150 }))
    .Paragraph(p => p
        .Text("Kerning, emphasis marks, outline, grid snapping & a run border (w:kern/w:em/w:outline/w:snapToGrid/w:bdr).")
        .Effects(new RunEffectsOptions
        {
            KerningMinPx = 12,
            EmphasisMark = EmphasisMark.Dot,
            Outline = true,
            SnapToGrid = false,
            Border = new CellBorder { Color = "#1a73e8", WidthPx = 1 },
        }))
    .Paragraph(p => p
        .Text("This run carries a French proofing language tag (w:lang) — round-tripped with no layout effect.")
        .Lang(new RunLangOptions { Val = "fr-FR" }))

    // ---- Content controls ----
    .Paragraph("Content controls", p => p.WithStyle("Heading1"))
    .Paragraph(p => p
        .Text("Every Word content-control kind round-trips: ")
        .RichTextControl("rich text", new SdtOptions { Alias = "Rich text" }).Text(", ")
        .PlainTextControl("plain text", new SdtOptions { Alias = "Plain text" }).Text(", a dropdown ")
        .DropDown("One", new[] { new SdtListItem("One", "1"), new SdtListItem("Two", "2") }, new SdtOptions { Alias = "Choice" }).Text(", a combo box ")
        .ComboBox("Alpha", new[] { new SdtListItem("Alpha", "a"), new SdtListItem("Beta", "b"), new SdtListItem("Gamma", "g") }, new SdtOptions { Alias = "Combo" }).Text(", a date picker ")
        .DateControl("6/16/2026", "M/d/yyyy", new SdtOptions { Alias = "Pick a date" }).Text(", and a checkbox ")
        .Checkbox(true).Text(", plus one with custom glyphs ")
        // w14:checkedState/uncheckedState — pin the exact checked/unchecked marks.
        .Checkbox(true, checkedSymbol: new CheckboxSymbol("MS Gothic", "2612"), uncheckedSymbol: new CheckboxSymbol("MS Gothic", "2610")).Text("."))

    // ---- Tables ----
    .Paragraph("Tables", p => p.WithStyle("Heading1"))
    .Paragraph("Merged cells (column- and row-spanning), with shading:")
    .Table(t => t
        .Row(r => r.Cell("Merged header spanning all three columns",
            new CellOptions { ColSpan = 3, Shading = "#1a73e8", Style = new CharStyle { Bold = true, Color = "#ffffff" } }))
        .Row(r => r
            .Cell("Spans 2 rows", new CellOptions { RowSpan = 2, Shading = "#e8f0fe", Style = new CharStyle { Bold = true } })
            .Cell("B1").Cell("C1"))
        .Row(r => r.Cell("B2").Cell("C2", new CellOptions { Style = new CharStyle { Color = "#188038" } })))
    .Paragraph("Cell vertical alignment (w:vAlign) — short labels sit top, centered and bottom within a tall row:")
    .Table(t => t
        .ColFractions(0.4, 0.2, 0.2, 0.2)
        .Row(r => r
            .Cell("vAlign", new CellOptions { Shading = "#1a73e8", Style = new CharStyle { Bold = true, Color = "#ffffff" } })
            .Cell("top", new CellOptions { Shading = "#1a73e8", Style = new CharStyle { Bold = true, Color = "#ffffff" } })
            .Cell("center", new CellOptions { Shading = "#1a73e8", Style = new CharStyle { Bold = true, Color = "#ffffff" } })
            .Cell("bottom", new CellOptions { Shading = "#1a73e8", Style = new CharStyle { Bold = true, Color = "#ffffff" } }))
        .Row(r => r
            .Cell("This tall cell holds several lines of text so the row grows well past the height of a single line — making the vertical position of its short neighbours visible.\nLine two.\nLine three.\nLine four.")
            .Cell("top", new CellOptions { VAlign = CellVAlign.Top, Style = new CharStyle { Color = "#188038" } })
            .Cell("center", new CellOptions { VAlign = CellVAlign.Center, Style = new CharStyle { Color = "#188038" } })
            .Cell("bottom", new CellOptions { VAlign = CellVAlign.Bottom, Style = new CharStyle { Color = "#188038" } })))
    .Paragraph("Minor & advanced table props (w:tblInd indent, w:bidiVisual right-to-left columns, w:textDirection / w:noWrap cells, plus caption/description alt text):")
    .Table(t => t
        .Row(r => r
            .Cell("Standings", new CellOptions { Shading = "#9c27b0", TextDirection = CellTextDirection.TbRl, Style = new CharStyle { Bold = true, Color = "#ffffff" } })
            .Cell("Rank 1", new CellOptions { Shading = "#9c27b0", NoWrap = true, Style = new CharStyle { Bold = true, Color = "#ffffff" } })
            .Cell("Rank 2", new CellOptions { Shading = "#9c27b0", NoWrap = true, Style = new CharStyle { Bold = true, Color = "#ffffff" } })
            .Cell("Rank 3", new CellOptions { Shading = "#9c27b0", NoWrap = true, Style = new CharStyle { Bold = true, Color = "#ffffff" } }))
        .Row(r => r
            .Cell("Points", new CellOptions { Style = new CharStyle { Bold = true } })
            .Cell("980", new CellOptions { HideMark = true, Style = new CharStyle { Color = "#188038" } })
            .Cell("845", new CellOptions { HideMark = true, Style = new CharStyle { Color = "#188038" } })
            .Cell("712", new CellOptions { HideMark = true, Style = new CharStyle { Color = "#188038" } })),
        new TableOptions
        {
            ColFractions = new[] { 0.4, 0.2, 0.2, 0.2 },
            Indent = 36,
            BidiVisual = true,
            Overlap = TableOverlap.Never,
            Caption = "Leaderboard",
            Description = "A right-to-left, indented table demonstrating issue #61 table properties.",
        })
    .Paragraph("Fields work inside table cells too:")
    .Table(t => t
        .Row(r => r.Cell("Metric", Bold()).Cell("Value", Bold()))
        .Row(r => r.Cell("Rendered on").Cell(s => s.Paragraph(p => p.DateField("yyyy-MM-dd"))))
        .Row(r => r.Cell("Page").Cell(s => s.Paragraph(p => p.Text("p. ").PageField()))))
    .Paragraph("And a table tall enough to paginate across pages — rows break cleanly:")
    .Table(t =>
    {
        t.Row(r => r.Cell("#", Bold()).Cell("Feature", Bold()).Cell("Status", Bold()));
        for (var i = 1; i <= 22; i++)
        {
            var n = i;
            t.Row(r => r.Cell(n.ToString())
                .Cell($"Demonstrated capability number {n} that keeps the table running past the bottom of the page")
                .Cell("supported", new CellOptions { Style = new CharStyle { Color = "#188038" } }));
        }
    })
    .Paragraph("Table-level defaults — borders, a shading fill and cell padding set once on the table (w:tblBorders / w:shd / w:tblCellMar) and round-tripped at that level instead of baked onto every cell:")
    .Table(new CellContent[][]
    {
        new CellContent[] { "Table-level default", "OOXML carrier (tblPr)" },
        new CellContent[] { "Borders (outer + interior)", "w:tblBorders" },
        new CellContent[] { "Shading fill", "w:shd" },
        new CellContent[] { "Cell margins / padding", "w:tblCellMar" },
        // A cell that explicitly CLEARS the table default shading (issue #150) — Word's
        // "No Color": stays unshaded among its shaded siblings and round-trips the clear.
        new CellContent[] { "Explicit clear (\"No Color\")", new CellSpec { Text = "w:shd val=clear fill=auto", ShadingCleared = true } },
    }, new TableOptions
    {
        HeaderRow = true,
        Borders = TableBorders.All(new CellBorder { Color = "#1a73e8", WidthPx = 1 }),
        Shading = "#eef5ff",
        CellMargin = new CellMargin(6, 10, 6, 10),
    })
    .Paragraph("Row properties (w:trPr) — a repeating header row (re-drawn atop each page), an exact-height row, and cant-split data rows kept whole:")
    .Table(t =>
    {
        t.Row(
            r => r.Cell("#", HeadCell()).Cell("Row property", HeadCell()).Cell("Effect", HeadCell()),
            new RowOptions { Header = true });
        t.Row(
            r => r.Cell("0").Cell("trHeight — exact 44px")
                .Cell("forced to exactly 44px tall", new CellOptions { Style = new CharStyle { Color = "#188038" } }),
            new RowOptions { Height = 44, HeightRule = RowHeightRule.Exact });
        for (var i = 1; i <= 20; i++)
        {
            var n = i;
            t.Row(
                r => r.Cell(n.ToString())
                    .Cell($"cantSplit data row {n} — kept whole across a page break")
                    .Cell("supported", new CellOptions { Style = new CharStyle { Color = "#188038" } }),
                new RowOptions { CantSplit = true });
        }
    }, new TableOptions { ColFractions = new double[] { 0.1, 0.6, 0.3 } })

    // ---- Rich text, lists & images ----
    .Paragraph("Rich text, lists & images", p => p.WithStyle("Heading1"))
    .Paragraph("An inline block image (a builder-supplied PNG, embedded into the PDF/DOCX):")
    .Image(blue, "image/png", new ImageOptions { WidthPx = 360, HeightPx = 110, Align = TextAlign.Center, Wrap = ImageWrap.Block })
    .Paragraph("A square-wrapped image floats and text flows around it. " + Repeat(Lorem, 3))
    .Image(purple, "image/png", new ImageOptions { WidthPx = 150, HeightPx = 110, Align = TextAlign.Left, Wrap = ImageWrap.Square })
    .Paragraph("A rotated image (18°) — the editor's rotate handle spins it and a:xfrm@rot round-trips:")
    .Image(blue, "image/png", new ImageOptions { WidthPx = 150, HeightPx = 110, Align = TextAlign.Center, Wrap = ImageWrap.Block, Rotation = 18 })
    .Paragraph("A linked (\"Link to File\") image: its bytes stay OUTSIDE the document — export re-emits an a:blip r:link + a TargetMode=\"External\" relationship instead of packing bytes:")
    .Image("https://raw.githubusercontent.com/git/git-scm.com/main/public/images/logos/downloads/Git-Icon-1788C.png",
        new ImageOptions { WidthPx = 96, HeightPx = 96, Align = TextAlign.Left, Linked = true })
    .Paragraph("Drawing shapes", p => p.WithStyle("Heading2"))
    .Paragraph("Vector preset shapes (OOXML wps:wsp / a:prstGeom) round-trip losslessly to .docx. The preset gallery — rectangle, rounded rectangle (with an adjust handle), ellipse, triangle, diamond and right/left arrows:")
    .Shape(ShapePreset.Rect, new ShapeOptions { WidthPx = 150, HeightPx = 80, Align = TextAlign.Left, Fill = ShapeFill.Solid("#bcd6ef"), Stroke = ShapeStroke.Solid("#41719c", 1.5) })
    .Shape(ShapePreset.RoundRect, new ShapeOptions { WidthPx = 150, HeightPx = 80, Align = TextAlign.Left, Fill = ShapeFill.Solid("#d9ead3"), Stroke = ShapeStroke.Solid("#38761d", 1.5), Adjust = new Dictionary<string, double> { ["adj"] = 18000 } })
    .Shape(ShapePreset.Ellipse, new ShapeOptions { WidthPx = 150, HeightPx = 90, Align = TextAlign.Left, Fill = ShapeFill.Solid("#f4cccc"), Stroke = ShapeStroke.Solid("#cc4125", 1) })
    .Shape(ShapePreset.Triangle, new ShapeOptions { WidthPx = 130, HeightPx = 90, Align = TextAlign.Center, Fill = ShapeFill.Solid("#fce5cd"), Stroke = ShapeStroke.Solid("#e69138", 1.5) })
    .Shape(ShapePreset.Diamond, new ShapeOptions { WidthPx = 130, HeightPx = 90, Align = TextAlign.Center, Fill = ShapeFill.Solid("#d9d2e9"), Stroke = ShapeStroke.Solid("#674ea7", 1.5) })
    .Shape(ShapePreset.RightArrow, new ShapeOptions { WidthPx = 180, HeightPx = 70, Align = TextAlign.Right, Fill = ShapeFill.Solid("#c9daf8"), Stroke = ShapeStroke.Solid("#3d6ea5", 1.5) })
    .Shape(ShapePreset.LeftArrow, new ShapeOptions { WidthPx = 180, HeightPx = 70, Align = TextAlign.Right, Fill = ShapeFill.Solid("#ead1dc"), Stroke = ShapeStroke.Solid("#a64d79", 1.5) })
    .Paragraph("Outline styles (a:prstDash) — solid, dash, dot, dash-dot and a thick long-dash:")
    .Shape(ShapePreset.Rect, new ShapeOptions { WidthPx = 150, HeightPx = 56, Align = TextAlign.Left, Fill = ShapeFill.Solid("#eeeeee"), Stroke = ShapeStroke.Solid("#333333", 1.5) })
    .Shape(ShapePreset.Rect, new ShapeOptions { WidthPx = 150, HeightPx = 56, Align = TextAlign.Left, Fill = ShapeFill.Solid("#eeeeee"), Stroke = ShapeStroke.Dashed("#333333", ShapeDash.Dash, 1.5) })
    .Shape(ShapePreset.Rect, new ShapeOptions { WidthPx = 150, HeightPx = 56, Align = TextAlign.Left, Fill = ShapeFill.Solid("#eeeeee"), Stroke = ShapeStroke.Dashed("#333333", ShapeDash.Dot, 1.5) })
    .Shape(ShapePreset.Rect, new ShapeOptions { WidthPx = 150, HeightPx = 56, Align = TextAlign.Left, Fill = ShapeFill.Solid("#eeeeee"), Stroke = ShapeStroke.Dashed("#333333", ShapeDash.DashDot, 1.5) })
    .Shape(ShapePreset.Rect, new ShapeOptions { WidthPx = 150, HeightPx = 56, Align = TextAlign.Left, Fill = ShapeFill.Solid("#eeeeee"), Stroke = ShapeStroke.Dashed("#333333", ShapeDash.LgDash, 3) })
    .Paragraph("Fill variants — a stroke-only diagonal line (no fill) and an outline-only rectangle (no fill):")
    .Shape(ShapePreset.Line, new ShapeOptions { WidthPx = 200, HeightPx = 60, Align = TextAlign.Left, Fill = ShapeFill.NoFill, Stroke = ShapeStroke.Solid("#38761d", 2) })
    .Shape(ShapePreset.Rect, new ShapeOptions { WidthPx = 160, HeightPx = 70, Align = TextAlign.Left, Fill = ShapeFill.NoFill, Stroke = ShapeStroke.Solid("#674ea7", 1) })
    .Paragraph("A rotated rectangle (20°) and a rotated arrow (−15°) — a:xfrm@rot round-trips:")
    .Shape(ShapePreset.Rect, new ShapeOptions { WidthPx = 150, HeightPx = 90, Align = TextAlign.Center, Fill = ShapeFill.Solid("#fff2cc"), Stroke = ShapeStroke.Solid("#bf9000", 1.5), Rotation = 20 })
    .Shape(ShapePreset.RightArrow, new ShapeOptions { WidthPx = 170, HeightPx = 80, Align = TextAlign.Center, Fill = ShapeFill.Solid("#d0e0e3"), Stroke = ShapeStroke.Solid("#134f5c", 1.5), Rotation = -15 })
    .Paragraph("A text box — a shape carrying a body of text (OOXML wps:txbx), editable in Word and the editor (double-click to type inside); edits round-trip losslessly:")
    .Shape(ShapePreset.Rect, new ShapeOptions
    {
        WidthPx = 300,
        HeightPx = 96,
        Align = TextAlign.Center,
        Fill = ShapeFill.Solid("#fff2cc"),
        Stroke = ShapeStroke.Solid("#bf9000", 1),
        Text = new[] { "Editable text box", "Double-click to edit this text, then click away to commit." },
    })
    .Paragraph("A freeform custom geometry (OOXML a:custGeom) — a five-pointed star traced as a path of line segments rather than a preset, round-tripping losslessly:")
    .Shape(ShapePreset.Rect, new ShapeOptions
    {
        WidthPx = 120,
        HeightPx = 120,
        Align = TextAlign.Center,
        Fill = ShapeFill.Solid("#ffe599"),
        Stroke = ShapeStroke.Solid("#bf9000", 1.5),
        Path = new[]
        {
            ShapePathSegment.MoveTo(0.5, 0.0),
            ShapePathSegment.LineTo(0.6176, 0.3382),
            ShapePathSegment.LineTo(0.9755, 0.3455),
            ShapePathSegment.LineTo(0.6902, 0.5618),
            ShapePathSegment.LineTo(0.7939, 0.9045),
            ShapePathSegment.LineTo(0.5, 0.7),
            ShapePathSegment.LineTo(0.2061, 0.9045),
            ShapePathSegment.LineTo(0.3098, 0.5618),
            ShapePathSegment.LineTo(0.0245, 0.3455),
            ShapePathSegment.LineTo(0.3824, 0.3382),
            ShapePathSegment.Close,
        },
    })
    .Paragraph("Shape positioning — wrap, float & z-order", p => p.WithStyle("Heading3"))
    .Paragraph("A square-wrapped shape floats at the margin and the paragraph text flows around it, exactly like a square-wrapped image (issue #217):")
    .Shape(ShapePreset.Ellipse, new ShapeOptions { WidthPx = 130, HeightPx = 100, Align = TextAlign.Left, Wrap = ImageWrap.Square, Fill = ShapeFill.Solid("#d9ead3"), Stroke = ShapeStroke.Solid("#38761d", 1.25) })
    .Paragraph("Square wrap lifts the shape out of the block flow and registers a float so this paragraph re-breaks beside it. " + Lorem + Lorem)
    .Paragraph("Below, two absolutely-anchored shapes overlap: the second has a higher z-order so it paints on top. A third shape sits BEHIND this text (a soft background tint) while text stays selectable, and a fourth sits IN FRONT of it.")
    // Overlapping anchored shapes — same layer, different z (higher paints on top).
    .Shape(ShapePreset.Rect, new ShapeOptions { WidthPx = 150, HeightPx = 90, Align = TextAlign.Left, Fill = ShapeFill.Solid("#c9daf8"), Stroke = ShapeStroke.Solid("#3d85c6", 1), Anchor = new ShapeAnchor { Behind = false, OffsetXPx = 20, OffsetYPx = 0, Z = 1 } })
    .Shape(ShapePreset.Ellipse, new ShapeOptions { WidthPx = 150, HeightPx = 90, Align = TextAlign.Left, Fill = ShapeFill.Solid("#fce5cd"), Stroke = ShapeStroke.Solid("#e69138", 1), Anchor = new ShapeAnchor { Behind = false, OffsetXPx = 90, OffsetYPx = 30, Z = 2 } })
    // Behind-text tint (selectable text stays on top) + an in-front-of-text accent.
    .Shape(ShapePreset.Rect, new ShapeOptions { WidthPx = 260, HeightPx = 70, Align = TextAlign.Left, Fill = ShapeFill.Solid("#fff2cc"), Stroke = ShapeStroke.NoOutline, Anchor = new ShapeAnchor { Behind = true, OffsetXPx = 0, OffsetYPx = 120, Z = -1 } })
    .Shape(ShapePreset.Ellipse, new ShapeOptions { WidthPx = 80, HeightPx = 80, Align = TextAlign.Left, Fill = ShapeFill.Solid("#ead1dc"), Stroke = ShapeStroke.Solid("#a64d79", 1), Anchor = new ShapeAnchor { Behind = false, OffsetXPx = 230, OffsetYPx = 110, Z = 3 } })
    .Paragraph("Anchored shapes do not occupy vertical flow space, so this line follows the previous paragraph directly while the shapes float over/under it. " + Lorem)
    .Paragraph("A grouped drawing (OOXML wpg:wgp) bundles several shapes into one object that moves and scales together — here a small flow diagram of two labelled boxes joined by an arrow:")
    .ShapeGroup(new ShapeGroupOptions
    {
        WidthPx = 360,
        HeightPx = 96,
        Align = TextAlign.Center,
        Children = new[]
        {
            new ShapeGroupChild { Preset = ShapePreset.RoundRect, XPx = 0, YPx = 20, WidthPx = 130, HeightPx = 56, Fill = ShapeFill.Solid("#d9ead3"), Stroke = ShapeStroke.Solid("#38761d", 1.5), Text = new[] { "Input" } },
            new ShapeGroupChild { Preset = ShapePreset.RightArrow, XPx = 140, YPx = 36, WidthPx = 80, HeightPx = 24, Fill = ShapeFill.Solid("#fff2cc"), Stroke = ShapeStroke.Solid("#bf9000", 1) },
            new ShapeGroupChild { Preset = ShapePreset.RoundRect, XPx = 230, YPx = 20, WidthPx = 130, HeightPx = 56, Fill = ShapeFill.Solid("#c9daf8"), Stroke = ShapeStroke.Solid("#3d85c6", 1.5), Text = new[] { "Output" } },
        },
    })
    .Paragraph("Multilevel numbered list:")
    .List(new[]
    {
        new ListItem("Model — invertible ops"),
        new ListItem("Layout — pretext pagination"),
        new ListItem("line caches keyed by (revision, width)", Level: 1),
        new ListItem("Paint — one fillText per fragment"),
    }, new ListOptions { Kind = ListKind.Number })
    .Paragraph("Bulleted list:")
    .List(new[]
    {
        new ListItem("markers are paint-only"),
        new ListItem("so caches survive renumbering", Level: 1),
    }, new ListOptions { Kind = ListKind.Bullet })
    // Explicit list opt-out (issue #152): a paragraph whose style would apply a list can
    // remove it with w:numId=0 (Word's "opt out"), so the opt-out round-trips instead of
    // the style's list silently returning on the next save.
    .Paragraph("This paragraph explicitly opts out of any list its style would apply (w:numId=0).",
        p => p.ClearList())
    // Per-paragraph list membership (ParagraphBuilder.ListItem): a single paragraph joins a
    // list at a given level — the counterpart to List() for an item that carries its own
    // paragraph formatting.
    .Paragraph("A standalone paragraph joined to the bullet list via ListItem.",
        p => p.ListItem("bullets"))
    .Paragraph("A fully justified, multi-page paragraph exercises line-level pagination. " + Repeat(Lorem, 12),
        p => p.Align(TextAlign.Justify))

    // ---- Paragraph borders & shading (w:pBdr / paragraph w:shd) ----
    .Paragraph("Paragraph borders & shading", p => p.WithStyle("Heading1"))
    .Paragraph("A whole paragraph can carry a border box and a background fill — Word's w:pBdr and paragraph-level w:shd. The box hugs the paragraph between its indents and round-trips to .docx and PDF.",
        p => p.Spacing(new SpacingOptions { Before = 6 })
              .Borders(ParaBorders.All(new CellBorder { Color = "#1a73e8", WidthPx = 1 }))
              .Shading("#eef4ff"))
    .Paragraph("Borders and shading are independent: this paragraph is shaded with no border, and a double-ruled box can sit on a plain background.",
        p => p.Spacing(new SpacingOptions { Before = 6 })
              .Indent(new IndentOptions { Left = 24, Right = 24 })
              .Shading("#fff3e0")
              .Borders(new ParaBorders { Left = new CellBorder { Color = "#e8710a", WidthPx = 3, Style = BorderStyle.Double } }))
    // Explicit "No Color" over a shaded named style (issue #147): the first paragraph inherits
    // the Callout style's green fill; the second uses the same style but CLEARS it, so the fill
    // is overridden and stays cleared through a .docx round-trip instead of silently returning.
    .Paragraph("This paragraph uses the Callout style and inherits its green shading fill.",
        p => p.WithStyle("Callout"))
    .Paragraph("This paragraph also uses Callout but explicitly clears the shading (No Color).",
        p => p.WithStyle("Callout").ClearShading())
    // Explicit border-box clear (issue #153): a paragraph whose style carries a w:pBdr box can
    // remove it with an empty w:pBdr, so the "no box" round-trips instead of the style's box returning.
    .Paragraph("This paragraph explicitly clears any border box its style would apply (empty w:pBdr).",
        p => p.ClearBorders())
    // ---- Contextual spacing: same-style runs sit tight (w:contextualSpacing) ----
    .Paragraph("Contextual spacing — each verse line below carries 12px after-spacing, yet w:contextualSpacing collapses the gaps between adjacent same-style paragraphs (Word's list-style default); only the run's outer edges keep their space:",
        p => p.Spacing(new SpacingOptions { Before = 10, After = 4 }))
    .Paragraph("Roses are red,", p => p.Spacing(new SpacingOptions { After = 12 }).ContextualSpacing())
    .Paragraph("violets are blue,", p => p.Spacing(new SpacingOptions { After = 12 }).ContextualSpacing())
    .Paragraph("contextual spacing keeps these lines tight,", p => p.Spacing(new SpacingOptions { After = 12 }).ContextualSpacing())
    .Paragraph("the way Word's list paragraphs do.", p => p.Spacing(new SpacingOptions { After = 12 }).ContextualSpacing())
    // Automatic paragraph spacing (issue #160): Word's w:beforeAutospacing/@w:afterAutospacing
    // (the HTML-<p> behavior). It bakes an approximate auto value and round-trips the attributes.
    .Paragraph("This paragraph uses Word's automatic before/after spacing (w:beforeAutospacing / w:afterAutospacing) instead of an explicit value.",
        p => p.Spacing(new SpacingOptions { BeforeAuto = true, AfterAuto = true }))

    // ---- Minor paragraph properties (w:widowControl / w:suppressLineNumbers / w:mirrorIndents / w:adjustRightInd / w:textAlignment) ----
    .Paragraph("Minor paragraph properties", p => p.WithStyle("Heading1"))
    .Paragraph("Lower-frequency w:pPr settings round-trip too: this paragraph disables widow/orphan control (w:widowControl), is excluded from line numbering (w:suppressLineNumbers), and carries mirrored indents (w:mirrorIndents) with right-indent adjustment (w:adjustRightInd) — each preserved through .docx. It also carries the unmodeled East-Asian / hyphenation toggles (w:snapToGrid, w:suppressAutoHyphens, w:kinsoku, w:overflowPunct, w:wordWrap, w:topLinePunct, w:autoSpaceDE, w:autoSpaceDN), round-tripped with no layout behavior.",
        p => p.Spacing(new SpacingOptions { Before = 6 })
              .WidowControl(false)
              .SuppressLineNumbers()
              .MirrorIndents()
              .AdjustRightInd()
              .SnapToGrid(false)
              .SuppressAutoHyphens()
              .Kinsoku(false)
              .OverflowPunct()
              .WordWrap(false)
              .TopLinePunct()
              .AutoSpaceDE(false)
              .AutoSpaceDN(false))
    .Paragraph("With extra line spacing, bottom vertical line alignment (w:textAlignment) drops the text onto the lower edge of each tall line box.",
        p => p.Spacing(new SpacingOptions { Before = 6, LineHeight = 2 })
              .TextAlignment(LineVAlign.Bottom))

    // ---- Miscellaneous OOXML round-trip (#63): symbols, image crop, tab stops ----
    .Paragraph("Miscellaneous OOXML — symbols, image crop & tab stops", p => p.WithStyle("Heading1"))
    .Paragraph(p => p
        .Text("Symbol-font glyphs (Word's w:sym) carry their font and code point, so they survive a round-trip: ")
        .Symbol("Wingdings", "F04A").Text("  ").Symbol("Wingdings", "F0FC").Text("  ").Symbol("Wingdings", "F0E0")
        .Text("  ").Symbol("Webdings", "F069")
        .Text(" — each is a glyph from a symbol font, not text."))
    .Paragraph("Image cropping (a:srcRect) trims the source to a window — here the same tile is shown whole, then center-cropped:")
    .Image(quad, "image/png", new ImageOptions { WidthPx = 150, HeightPx = 110, Align = TextAlign.Left, Wrap = ImageWrap.Block })
    .Image(quad, "image/png", new ImageOptions { WidthPx = 150, HeightPx = 110, Align = TextAlign.Left, Wrap = ImageWrap.Block, Crop = new ImageCrop(0.25, 0.2, 0.25, 0.2) })
    .Paragraph("Tab stops honor the document's default interval (w:defaultTabStop): columns\tline up\tat\teach default tab.")
    // Explicit tab clear (issue #154): a paragraph can REMOVE a tab an inherited style
    // provides with w:val="clear" — the removal round-trips instead of the style tab reappearing.
    .Paragraph("This paragraph sets a left tab at 1.5\" and explicitly clears the tab at 1\" (w:val=clear).",
        p => p.TabStops(new[] { new TabStop { PosPx = 96, Cleared = true }, new TabStop { PosPx = 144, Align = TabAlign.Left } }))

    // ---- International text (CJK) ----
    .Paragraph("International text — CJK", p => p.WithStyle("Heading1"))
    .Paragraph("East-Asian scripts lay out the way Word does — measured on canvas, with Unicode line-breaking and kinsoku, not the browser's contenteditable.")
    .Paragraph("日本語 — CJK line-breaking & kinsoku", p => p.Bold().Color(Ink))
    .Paragraph("日本語の文章は単語の間にスペースを入れません。それでもエンジンは文字単位で行を折り返し、句読点が行頭に来ないように禁則処理（kinsoku）を行います。「角括弧」のような約物も正しく扱われ、長い段落でもページをまたいで自然に流れます。")

    // ---- Run-level typography (tracking, theme tints, CS/EA font slots) ----
    // These all survive the .docx round-trip. The complex-script (w:cs) slot is set
    // on Latin text so it stays legible in the headless PDF (it only changes glyph
    // selection for complex scripts), while the East-Asian (w:eastAsia) slot rides
    // along with CJK text Word renders directly.
    .Paragraph("Run-level typography", p => p.WithStyle("Heading1"))
    .Paragraph(p => p
        .Text("Character tracking (OOXML w:spacing) widens ")
        .Text("w i d e", new CharStyle { LetterSpacingPx = 3 })
        .Text(" or tightens ")
        .Text("tight", new CharStyle { LetterSpacingPx = -0.5 })
        .Text(" inter-letter spacing; on import a theme color's tint/shade resolves to its actual shade — here the resolved ")
        .Text("60% tint", new CharStyle { Color = "#8faadc", Bold = true })
        .Text(" and ")
        .Text("50% shade", new CharStyle { Color = "#223962", Bold = true })
        .Text(" of the Office accent blue #4472C4."))
    .Paragraph(p => p
        .Text("Runs preserve their East-Asian (w:eastAsia) and ")
        .Text("complex-script", new CharStyle { FontFamilyComplexScript = "Scheherazade New, serif" })
        .Text(" (w:cs) font slots independently of the Latin face — ")
        .Text("日本語", new CharStyle { FontFamilyEastAsia = "Yu Mincho, serif" })
        .Text(" carries its CJK typeface through the round-trip."))
    // Explicit highlight clear (issue #155): a highlighted run, then a run that CLEARS the
    // highlight (w:highlight=none) so it overrides an inherited character-style highlight.
    .Paragraph(p => p
        .Text("A highlighted run ", new CharStyle { HighlightColor = "#ffff00" })
        .Text("and a run that explicitly clears highlight", new CharStyle { HighlightCleared = true })
        .Text(" — the clear round-trips as w:highlight=none."))

    // ---- Equations (LaTeX + MathML, typeset on canvas, round-tripped to OMML) ----
    .Paragraph("Equations", p => p.WithStyle("Heading1"))
    .Paragraph(p => p
        .Text("Equations are first-class — the identity ")
        .InlineEquation("e^{i\\pi}+1=0")
        .Text(" flows inline with the text, while display equations get their own centered line:"))
    .Equation("\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}")
    .Paragraph("The same equation can be authored as MathML, here right-aligned:")
    .EquationMathml(
        "<math><mrow><msup><mi>a</mi><mn>2</mn></msup><mo>+</mo><msup><mi>b</mi><mn>2</mn></msup><mo>=</mo><msup><mi>c</mi><mn>2</mn></msup></mrow></math>",
        new EquationOptions { Align = EquationAlign.Right })
    .Paragraph("Equations can be drag-resized in the editor; the scale round-trips to .docx:")
    .Equation("E = mc^2", new EquationOptions { Scale = 1.6 })

    // ---- Right-to-left / bidirectional ----
    // NOTE: .Direction(Rtl) drives a true bidi (OOXML w:bidi) layout. We demo it with
    // Latin text because the HEADLESS PDF export bundles only Latin + STIX-math + CJK
    // fallback fonts — Arabic/Hebrew glyphs would render as tofu here (the .docx still
    // carries the text and Word renders it). The base-direction effect (right
    // alignment + mirrored start/end indents) is visible regardless of script.
    .Paragraph("Right-to-left & bidirectional", p => p.WithStyle("Heading1"))
    .Paragraph("Setting a paragraph's base direction to RTL lays it out right-to-left (OOXML w:bidi): it starts at the right edge and start/end alignment + indents mirror. Arabic and Hebrew flow through this exact path; bracketed numbers like [2024] keep their bidi order.",
        p => p.Direction(Direction.Rtl))

    // ---- A table using the registered real table style ----
    .Paragraph("Table styles", p => p.WithStyle("Heading1"))
    .Paragraph("This table references the \"Showcase Grid\" style registered above — its header band, borders, and row banding are baked onto the cells and emitted as w:tblStyle:")
    .Table(new CellContent[][]
    {
        new CellContent[] { "Layer", "Role" },
        new CellContent[] { "Model", "invertible ops + undo" },
        new CellContent[] { "Layout", "pretext pagination" },
        new CellContent[] { "Paint", "one fillText per fragment" },
    }, new TableOptions { StyleId = "ShowcaseGrid" })

    // ---- Fixed line spacing (w:lineRule exact/atLeast) ----
    .Paragraph("Fixed line spacing", p => p.WithStyle("Heading1"))
    .Paragraph("Beyond a line-height multiplier, Word supports fixed point spacing via w:lineRule: \"exact\" pins every line to a height (taller glyphs clip), \"atLeast\" floors it but grows for a taller line. Both round-trip through .docx and drive pagination.")
    .Paragraph("This paragraph uses EXACT 28px line spacing — every line box is exactly 28px tall, so the lines sit at a constant pitch no matter how the text wraps across the page.",
        p => p.Spacing(new SpacingOptions { LineRule = LineRule.Exact, LineHeightPx = 28 }))
    .Paragraph(p => p
        .Text("This paragraph uses AT-LEAST 24px line spacing — lines are at least 24px tall, but a line with a larger glyph, ")
        .Text("like this 30px word", new CharStyle { FontSizePx = 30 })
        .Text(", grows to fit it.")
        .Spacing(new SpacingOptions { LineRule = LineRule.AtLeast, LineHeightPx = 24 }))
    // ---- Section breaks (odd-page parity) + line numbering ----
    .Paragraph("Section breaks & line numbering", p => p.WithStyle("Heading1"))
    .Paragraph("Word sections can force their first page onto an odd or even page number (a blank filler page is inserted when the running page count has the wrong parity) and print a number beside every line in the margin (w:lnNumType). The next section demonstrates both.")
    .SectionBreak() // plain Next Page break: ends the main flow as its own section
    .Paragraph("A line-numbered section, starting on an odd page", p => p.WithStyle("Heading2"))
    .Paragraph("Every line in this section carries a small margin line number; counting restarts on each page (Word's default) and round-trips to .docx and PDF. " + Repeat(Lorem, 6), p => p.Align(TextAlign.Justify))
    // This break ENDS the line-numbered section: its w:type is oddPage and its
    // sectPr carries the w:lnNumType that numbered every line above.
    .SectionBreak(new SectionBreakOptions { BreakType = SectionBreakType.OddPage, LineNumbering = new LineNumbering { CountBy = 1, Restart = LineNumberRestart.NewPage } })
    .Paragraph("Line numbering is a per-section property, so this closing line — in the final, unnumbered section — has no margin numbers.")

    // ---- Header + footer with page fields ----
    .Header(h => h.Paragraph(p => p.Font("Arial, sans-serif").FontSize(11)
        .Text("canvas-word", new CharStyle { Bold = true, Color = "#5f6368" })
        .Text("  ·  feature showcase", new CharStyle { Color = "#9aa0a6" })))
    .Footer(f => f.Paragraph(p => p.Align(TextAlign.Center).Font("Arial, sans-serif").FontSize(11).Color("#9aa0a6")
        .Text("Page ").PageField().Text(" of ").NumPagesField()));

// ---- Build + export -------------------------------------------------------
var doc = builder.Build();
foreach (var w in builder.Warnings) Console.WriteLine($"  builder warning: {w.Code} {w.Detail}");

var outDir = AppContext.BaseDirectory;
var pdf = doc.ExportPdf();
var docx = doc.ExportDocx();
File.WriteAllBytes(Path.Combine(outDir, "showcase.pdf"), pdf);
File.WriteAllBytes(Path.Combine(outDir, "showcase.docx"), docx);

var re = engine.ImportDocx(docx);
Console.WriteLine($"Built {doc.BlockCount} blocks with the typed builder.");
Console.WriteLine($"  showcase.pdf  : {pdf.Length / 1024} KB");
Console.WriteLine($"  showcase.docx : {docx.Length / 1024} KB");
Console.WriteLine($"  round-trip    : re-imported to {re.BlockCount} blocks, {re.MediaCount} images, {re.Warnings.Count} warnings");

// ---- Query the re-imported document (WordprocessingDocument-style access) --
var paragraphs = re.GetParagraphs();
var sections = re.GetSections();
var pages = re.GetPages();
Console.WriteLine("Query:");
Console.WriteLine($"  paragraphs    : {paragraphs.Count} (first: \"{Trunc(paragraphs.Count > 0 ? paragraphs[0].Text : "")}\")");
Console.WriteLine($"  sections      : {sections.Count}");
Console.WriteLine($"  pages         : {pages.Count}");
var hits = re.FindText("Showcase");
Console.WriteLine($"  find 'Showcase': {hits.Count} paragraph(s)");
foreach (var h in hits.Take(3))
{
    var pageNo = pages.FirstOrDefault(p => p.BlockIds.Contains(h.Id))?.Number;
    Console.WriteLine($"      [{h.Container}] {(pageNo is { } n ? $"p{n} " : "")}\"{Trunc(h.Text)}\"");
}

// ---- Edit the document in place (undo/redo), then round-trip --------------
var firstId = paragraphs[0].Id;
var editor = re.Edit()
    .SetParagraphText(firstId, "EDITED HEADING")
    .InsertParagraphAfter(firstId, "Inserted by the C# editor.");
Console.WriteLine("Edit:");
Console.WriteLine($"  after 2 edits : first=\"{editor.ToDocument().GetParagraphs()[0].Text}\", blocks={editor.ToDocument().BlockCount}, canUndo={editor.CanUndo}");
editor.Undo(); // undo the insert
editor.Undo(); // undo the text change
Console.WriteLine($"  after 2 undos : first=\"{editor.ToDocument().GetParagraphs()[0].Text}\", blocks={editor.ToDocument().BlockCount}, canRedo={editor.CanRedo}");
editor.Redo(); // redo the text change
var editedDocx = editor.ToDocument().ExportDocx();
var reEdited = engine.ImportDocx(editedDocx);
Console.WriteLine($"  round-trip    : re-imported edited docx → find 'EDITED HEADING' = {reEdited.FindText("EDITED HEADING").Count}");

// ---- Content controls (SDTs) — the templating surface ---------------------
var sdts = re.GetSdts();
Console.WriteLine("Content controls (SDTs):");
Console.WriteLine($"  total         : {sdts.Count} ({re.GetSdtRoots(sdts).Count} root)"); // reuse the fetched list
foreach (var s in sdts.Take(4))
    Console.WriteLine($"      [{s.SdtType}] {(s.Alias is { } a ? $"\"{a}\" " : "")}depth={s.Depth} text=\"{Trunc(s.Text)}\"");

var plain = sdts.FirstOrDefault(s => s.Alias == "Plain text");
var check = sdts.FirstOrDefault(s => s.SdtType == "checkbox");
var choice = sdts.FirstOrDefault(s => s.SdtType == "dropDown");
if (plain is not null && check is not null && choice is not null)
{
    var sdtEditor = re.Edit()
        .SetSdtText(plain.Id, "FILLED BY C#")
        .SetSdtProps(plain.Id, new SdtPropsPatch { Tag = "FilledField" })
        .SetCheckbox(check.Id, false)
        .SetSdtValue(choice.Id, "2"); // select the dropdown option whose value is "2"
    var after = sdtEditor.ToDocument().GetSdts();
    var plainAfter = after.First(s => s.Id == plain.Id);
    var checkAfter = after.First(s => s.Id == check.Id);
    var choiceAfter = after.First(s => s.Id == choice.Id);
    Console.WriteLine($"  after fill    : plain text=\"{plainAfter.Text}\" tag={plainAfter.Tag}, checkbox checked={checkAfter.Checked}, dropdown=\"{choiceAfter.Text}\"");
}

// ---- More query getters (fields / styles / bookmarks / notes / text) -------
Console.WriteLine("More query:");
Console.WriteLine($"  fields        : {re.GetFields().Count}, styles: {re.GetStyles().Count}, bookmarks: {re.GetBookmarks().Count}, footnotes: {re.GetFootnotes().Count}");
var pos = re.PositionOfText("Showcase");
Console.WriteLine($"  positionOf 'Showcase': {(pos is { } pp ? $"{Trunc(pp.BlockId)}@{pp.Offset}" : "not found")}");
if (pos is { } p2)
    Console.WriteLine($"  rangeText     : \"{Trunc(re.RangeText(p2.BlockId, p2.Offset, p2.BlockId, p2.Offset + 8))}\"");
if (choice is not null) // reuse the dropdown SDT already fetched above
{
    var val = re.GetSdtValue(choice.Id);
    Console.WriteLine($"  dropdown value: text=\"{val?.Text}\" selected={val?.Selected}");
}
Console.WriteLine($"  blockPath[0]  : container={re.GetBlockPath(paragraphs[0].Id)?.Container}"); // reuse `paragraphs`

// ---- Ergonomic bulk / structural edits (find/replace, move, table) ---------
var ergo = re.Edit();
var firstBlock = re.GetParagraphs()[0].Id;
ergo.ReplaceAllText("canvas-word", "CANVAS-WORD");          // literal
var literalHits = ergo.ToDocument().FindText("CANVAS-WORD").Count;
ergo.ReplaceAllText("CANVAS-\\w+", "cw", "g");              // regex overload → collapse to "cw"
ergo.MoveBlock(firstBlock, 2);
var ergoDoc = ergo.ToDocument();
Console.WriteLine("Ergonomic edits:");
Console.WriteLine($"  replaceAllText: literal hits={literalHits}, after regex 'cw' hits={ergoDoc.FindText("cw").Count}");
Console.WriteLine($"  moveBlock     : first block moved away from the top = {ergoDoc.GetParagraphs()[0].Id != firstBlock}");

// setParagraphStyle: patch the first paragraph (align + outline level + spacing).
var styled = re.Edit();
styled.SetParagraphStyle(firstBlock, new ParaStylePatch { Align = TextAlign.Center, OutlineLevel = 0, SpaceBeforePx = 12 });
Console.WriteLine($"  setParaStyle  : first block outlineLevel now = {styled.ToDocument().GetParagraphs()[0].OutlineLevel}");

// ---- Merge / append documents ----------------------------------------------
// Build two small extra parts and fold them after the showcase document, then
// export the combined report. Each seam starts a new page; every id space
// (styles, lists, controls, bookmarks, notes) is reconciled automatically.
var appendix = engine.NewBuilder(new CreateOptions { PageSize = PageSizeName.Letter })
    .Paragraph("Appendix A", p => p.WithStyle("Heading1"))
    .Paragraph("Appended as a separate part via merge.")
    .Build();
var glossary = engine.NewBuilder(new CreateOptions { PageSize = PageSizeName.A4 }) // distinct geometry → a real section
    .Paragraph("Glossary", p => p.WithStyle("Heading1"))
    .Paragraph("Merge — folding documents into one, rebasing every id space.")
    .Build();

var mergedReport = engine.Merge(doc, appendix, glossary); // == doc.Append(appendix).Append(glossary)
var mergedDocx = mergedReport.ExportDocx();
File.WriteAllBytes(Path.Combine(outDir, "showcase-merged.docx"), mergedDocx);
var mergedBack = engine.ImportDocx(mergedDocx);
Console.WriteLine("Merge:");
Console.WriteLine($"  parts         : showcase({doc.BlockCount}) + appendix({appendix.BlockCount}) + glossary({glossary.BlockCount})");
Console.WriteLine($"  merged        : {mergedReport.BlockCount} blocks, {mergedReport.MediaCount} images → showcase-merged.docx ({mergedDocx.Length / 1024} KB)");
Console.WriteLine($"  round-trip    : re-imported to {mergedBack.BlockCount} blocks, sections={mergedBack.GetSections().Count}, appendix found={mergedBack.FindText("Appendix A").Count > 0}, glossary found={mergedBack.FindText("Glossary").Count > 0}");

// ---- Per-section footer via SetSectionFooter (table + logo + PAGE field) ----
// Author a branded content footer (borderless 2-col table: logo | right-aligned
// address, then a centered "Page X of Y") and apply it to the merged report's body
// section. The page number is a live field resolved at layout — no placeholder.
var footerLogo = SolidPng(64, 24, 0x1a, 0x73, 0xe8);
var footerEditor = mergedReport.Edit();
var lastSection = mergedReport.GetSections().Count - 1;
footerEditor.SetSectionFooter(lastSection, f => BuildContentFooter(f, footerLogo));
var footerDocx = footerEditor.ToDocument().ExportDocx();
File.WriteAllBytes(Path.Combine(outDir, "showcase-footer.docx"), footerDocx);
Console.WriteLine("Section footer:");
Console.WriteLine($"  applied        : content footer (table + logo + PAGE) to body section {lastSection} → showcase-footer.docx ({footerDocx.Length / 1024} KB)");

Console.WriteLine($"Output written to: {outDir}");

static string Trunc(string s) => s.Length <= 60 ? s : s[..57] + "...";

// A reusable branded content footer: a borderless [ logo | address ] table over a
// centered "Page X of Y". The address wraps on its own (no manual line-splitting)
// and PageField/NumPagesField resolve at layout (no "update fields" step).
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
    }, new TableOptions { WidthMode = TableWidthMode.AutofitWindow, Borders = new TableBorders() });
    f.Paragraph(p => p
        .Align(TextAlign.Center).Font(Font).FontSize(SizePx).Spacing(new SpacingOptions { Before = 7, After = 0 })
        .Text("Page ").PageField().Text(" of ").NumPagesField());
}
return;

static CellOptions Bold() => new() { Style = new CharStyle { Bold = true } };

// A header-band cell: bold white text on the showcase blue (matches the sample doc).
static CellOptions HeadCell() => new() { Shading = "#1a73e8", Style = new CharStyle { Bold = true, Color = "#ffffff" } };

static string Repeat(string s, int n) => string.Concat(Enumerable.Repeat(s, n));

// Minimal solid-color PNG encoder (RGB, no deps) so the builder has real image bytes
// to embed — pdfkit can't embed the SVG data URLs the in-browser sample uses.
// A non-uniform tile: four colored quadrants over a contrasting center square, so a
// center crop (a:srcRect) is visibly different from the full image in the showcase.
static byte[] QuadrantPng(int w, int h)
{
    var raw = new byte[h * (1 + w * 3)];
    var p = 0;
    for (var y = 0; y < h; y++)
    {
        raw[p++] = 0; // filter: none
        for (var x = 0; x < w; x++)
        {
            // center band (the part the demo crop keeps) is white; quadrants differ.
            bool cx = x > w / 4 && x < 3 * w / 4, cy = y > h / 4 && y < 3 * h / 4;
            byte r, g, b;
            if (cx && cy) { r = 0xff; g = 0xff; b = 0xff; }
            else if (x < w / 2 && y < h / 2) { r = 0x1a; g = 0x73; b = 0xe8; }
            else if (x >= w / 2 && y < h / 2) { r = 0x9c; g = 0x27; b = 0xb0; }
            else if (x < w / 2) { r = 0x18; g = 0x80; b = 0x38; }
            else { r = 0xe8; g = 0x71; b = 0x0a; }
            raw[p++] = r; raw[p++] = g; raw[p++] = b;
        }
    }
    return EncodePng(w, h, raw);
}

static byte[] SolidPng(int w, int h, byte r, byte g, byte b)
{
    var raw = new byte[h * (1 + w * 3)];
    var p = 0;
    for (var y = 0; y < h; y++)
    {
        raw[p++] = 0; // filter: none
        for (var x = 0; x < w; x++) { raw[p++] = r; raw[p++] = g; raw[p++] = b; }
    }
    return EncodePng(w, h, raw);
}

static byte[] EncodePng(int w, int h, byte[] raw)
{

    byte[] idat;
    using (var z = new MemoryStream())
    {
        using (var zl = new ZLibStream(z, CompressionLevel.Optimal, leaveOpen: true)) zl.Write(raw, 0, raw.Length);
        idat = z.ToArray();
    }

    using var ms = new MemoryStream();
    ms.Write(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 });

    var ihdr = new byte[13];
    WriteBE(ihdr, 0, (uint)w);
    WriteBE(ihdr, 4, (uint)h);
    ihdr[8] = 8;  // bit depth
    ihdr[9] = 2;  // color type: truecolor RGB
    Chunk(ms, "IHDR", ihdr);
    Chunk(ms, "IDAT", idat);
    Chunk(ms, "IEND", Array.Empty<byte>());
    return ms.ToArray();

    static void WriteBE(byte[] a, int o, uint v) { a[o] = (byte)(v >> 24); a[o + 1] = (byte)(v >> 16); a[o + 2] = (byte)(v >> 8); a[o + 3] = (byte)v; }

    static void Chunk(Stream s, string type, byte[] data)
    {
        var len = new byte[4]; WriteBE(len, 0, (uint)data.Length); s.Write(len);
        var t = System.Text.Encoding.ASCII.GetBytes(type);
        s.Write(t); s.Write(data);
        uint c = 0xffffffff;
        foreach (var bb in t) c = Crc(c, bb);
        foreach (var bb in data) c = Crc(c, bb);
        var crc = new byte[4]; WriteBE(crc, 0, c ^ 0xffffffff); s.Write(crc);
    }

    static uint Crc(uint c, byte bb)
    {
        c ^= bb;
        for (var k = 0; k < 8; k++) c = (c & 1) != 0 ? 0xedb88320 ^ (c >> 1) : c >> 1;
        return c;
    }
}
