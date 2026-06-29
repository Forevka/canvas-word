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

    // ---- Content controls ----
    .Paragraph("Content controls", p => p.WithStyle("Heading1"))
    .Paragraph(p => p
        .Text("Every Word content-control kind round-trips: ")
        .RichTextControl("rich text", new SdtOptions { Alias = "Rich text" }).Text(", ")
        .PlainTextControl("plain text", new SdtOptions { Alias = "Plain text" }).Text(", a dropdown ")
        .DropDown("One", new[] { new SdtListItem("One", "1"), new SdtListItem("Two", "2") }, new SdtOptions { Alias = "Choice" }).Text(", a combo box ")
        .ComboBox("Alpha", new[] { new SdtListItem("Alpha", "a"), new SdtListItem("Beta", "b"), new SdtListItem("Gamma", "g") }, new SdtOptions { Alias = "Combo" }).Text(", a date picker ")
        .DateControl("6/16/2026", "M/d/yyyy", new SdtOptions { Alias = "Pick a date" }).Text(", and a checkbox ")
        .Checkbox(true).Text("."))

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
    // ---- Contextual spacing: same-style runs sit tight (w:contextualSpacing) ----
    .Paragraph("Contextual spacing — each verse line below carries 12px after-spacing, yet w:contextualSpacing collapses the gaps between adjacent same-style paragraphs (Word's list-style default); only the run's outer edges keep their space:",
        p => p.Spacing(new SpacingOptions { Before = 10, After = 4 }))
    .Paragraph("Roses are red,", p => p.Spacing(new SpacingOptions { After = 12 }).ContextualSpacing())
    .Paragraph("violets are blue,", p => p.Spacing(new SpacingOptions { After = 12 }).ContextualSpacing())
    .Paragraph("contextual spacing keeps these lines tight,", p => p.Spacing(new SpacingOptions { After = 12 }).ContextualSpacing())
    .Paragraph("the way Word's list paragraphs do.", p => p.Spacing(new SpacingOptions { After = 12 }).ContextualSpacing())

    // ---- Minor paragraph properties (w:widowControl / w:suppressLineNumbers / w:mirrorIndents / w:adjustRightInd / w:textAlignment) ----
    .Paragraph("Minor paragraph properties", p => p.WithStyle("Heading1"))
    .Paragraph("Lower-frequency w:pPr settings round-trip too: this paragraph disables widow/orphan control (w:widowControl), is excluded from line numbering (w:suppressLineNumbers), and carries mirrored indents (w:mirrorIndents) with right-indent adjustment (w:adjustRightInd) — each preserved through .docx.",
        p => p.Spacing(new SpacingOptions { Before = 6 })
              .WidowControl(false)
              .SuppressLineNumbers()
              .MirrorIndents()
              .AdjustRightInd())
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

    // ---- International text (CJK) ----
    .Paragraph("International text — CJK", p => p.WithStyle("Heading1"))
    .Paragraph("East-Asian scripts lay out the way Word does — measured on canvas, with Unicode line-breaking and kinsoku, not the browser's contenteditable.")
    .Paragraph("日本語 — CJK line-breaking & kinsoku", p => p.Bold().Color(Ink))
    .Paragraph("日本語の文章は単語の間にスペースを入れません。それでもエンジンは文字単位で行を折り返し、句読点が行頭に来ないように禁則処理（kinsoku）を行います。「角括弧」のような約物も正しく扱われ、長い段落でもページをまたいで自然に流れます。")
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
Console.WriteLine($"Output written to: {outDir}");
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
