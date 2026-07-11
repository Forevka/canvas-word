using Microsoft.ClearScript;

namespace WordCanvas.ClearScript.Builder;

// Typed option records + enums mirroring @forevka/wordcanvas/builder. Each builds a
// native JS plain-object/array (via the engine) carrying only the fields that are
// set, so it maps exactly onto the JS builder's Partial<...> option shapes.

/// <summary>Helpers to construct native JS values the builder methods consume.</summary>
internal static class Js
{
    public static ScriptObject Obj(WordCanvasEngine e) => (ScriptObject)e.NewObject();
    public static ScriptObject Arr(WordCanvasEngine e) => (ScriptObject)e.NewArray();

    public static void Set(ScriptObject o, string key, object? value)
    {
        if (value is not null) o.SetProperty(key, value);
    }

    public static void Push(ScriptObject arr, object? value) => arr.InvokeMethod("push", value!);

    /// <summary>Marshal an IEnumerable to a native JS array.</summary>
    public static ScriptObject ToArray(WordCanvasEngine e, IEnumerable<object?> items)
    {
        var arr = Arr(e);
        foreach (var it in items) Push(arr, it);
        return arr;
    }
}

public enum BandVariant { Default, First, Even }
public enum ListKind { Bullet, Number }
public enum TextAlign { Left, Center, Right, Justify }
public enum Orientation { Portrait, Landscape }
/// <summary>How a section begins (OOXML w:sectPr/w:type).</summary>
public enum SectionBreakType { NextPage, EvenPage, OddPage }
/// <summary>When line numbering restarts (OOXML w:lnNumType/@w:restart).</summary>
public enum LineNumberRestart { Continuous, NewPage, NewSection }
public enum ImageWrap { Block, Square }
public enum PageSizeName { Letter, Legal, A4, A3, Tabloid }
/// <summary>Block/inline equation horizontal placement (display equations only).</summary>
public enum EquationAlign { Left, Center, Right }
/// <summary>Paragraph base writing direction (OOXML w:bidi).</summary>
public enum Direction { Ltr, Rtl }
/// <summary>Fixed line-spacing rule (docx w:spacing/@w:lineRule).</summary>
public enum LineRule { Exact, AtLeast }
/// <summary>Tab-stop alignment (docx w:tab/@w:val).</summary>
public enum TabAlign { Left, Center, Right, Decimal }
/// <summary>Tab-stop leader fill (docx w:tab/@w:leader).</summary>
public enum TabLeader { None, Dot, Dash, Underscore }
/// <summary>Column sizing strategy (docx w:tblLayout + w:tblW).</summary>
public enum TableWidthMode { Fixed, AutofitContents, AutofitWindow }
/// <summary>Cell border line style.</summary>
public enum BorderStyle { Single, Double, Dashed, Dotted }

/// <summary>Emphasis-mark style drawn over each character (OOXML w:em).</summary>
public enum EmphasisMark { Dot, Comma, Circle, UnderDot }
/// <summary>Preferred-width unit: absolute px or percent of table width.</summary>
public enum WidthType { Abs, Pct }
/// <summary>Cell vertical content alignment (OOXML w:tcPr/w:vAlign).</summary>
public enum CellVAlign { Top, Center, Bottom }
/// <summary>How a row's fixed height is enforced (OOXML w:trHeight/@w:hRule).</summary>
public enum RowHeightRule { AtLeast, Exact }
/// <summary>Vertical alignment of text within a line box (OOXML w:pPr/w:textAlignment).</summary>
public enum LineVAlign { Top, Center, Bottom, Baseline }
/// <summary>Cell text flow direction (OOXML w:tcPr/w:textDirection).</summary>
public enum CellTextDirection { LrTb, TbRl, BtLr, LrTbV, TbRlV, TbLrV }
/// <summary>Floating-table overlap behavior (OOXML w:tblPr/w:tblOverlap).</summary>
public enum TableOverlap { Never, Overlap }
/// <summary>List number format (docx w:numFmt).</summary>
public enum ListNumberFormat { Bullet, Decimal, LowerLetter, UpperLetter, LowerRoman, UpperRoman }
/// <summary>Conditional-format slot of a table style (OOXML w:tblStylePr/@w:type).</summary>
public enum TableCond
{
    WholeTable, FirstRow, LastRow, FirstCol, LastCol,
    Band1Horz, Band2Horz, Band1Vert, Band2Vert,
    NwCell, NeCell, SwCell, SeCell,
}

internal static class EnumJs
{
    public static string Align(TextAlign a) => a switch
    {
        TextAlign.Left => "left",
        TextAlign.Center => "center",
        TextAlign.Right => "right",
        TextAlign.Justify => "justify",
        _ => "left",
    };

    public static string Wrap(ImageWrap w) => w == ImageWrap.Square ? "square" : "block";
    public static string Orient(Orientation o) => o == Orientation.Landscape ? "landscape" : "portrait";
    public static string BreakType(SectionBreakType t) => t switch
    {
        SectionBreakType.EvenPage => "evenPage",
        SectionBreakType.OddPage => "oddPage",
        _ => "nextPage",
    };
    public static string LineRestart(LineNumberRestart r) => r switch
    {
        LineNumberRestart.Continuous => "continuous",
        LineNumberRestart.NewSection => "newSection",
        _ => "newPage",
    };
    public static string Variant(BandVariant v) => v switch
    {
        BandVariant.First => "first",
        BandVariant.Even => "even",
        _ => "default",
    };
    public static string PageSize(PageSizeName n) => n.ToString();

    public static string EqAlign(EquationAlign a) => a switch
    {
        EquationAlign.Left => "left",
        EquationAlign.Right => "right",
        _ => "center",
    };
    public static string Dir(Direction d) => d == Direction.Rtl ? "rtl" : "ltr";
    public static string LineRule(LineRule r) => r == Builder.LineRule.Exact ? "exact" : "atLeast";
    public static string TabAlign(TabAlign a) => a switch
    {
        Builder.TabAlign.Center => "center",
        Builder.TabAlign.Right => "right",
        Builder.TabAlign.Decimal => "decimal",
        _ => "left",
    };
    public static string TabLeader(TabLeader l) => l switch
    {
        Builder.TabLeader.Dot => "dot",
        Builder.TabLeader.Dash => "dash",
        Builder.TabLeader.Underscore => "underscore",
        _ => "none",
    };
    public static string WidthMode(TableWidthMode m) => m switch
    {
        TableWidthMode.AutofitContents => "autofitContents",
        TableWidthMode.AutofitWindow => "autofitWindow",
        _ => "fixed",
    };
    public static string Border(BorderStyle s) => s switch
    {
        BorderStyle.Double => "double",
        BorderStyle.Dashed => "dashed",
        BorderStyle.Dotted => "dotted",
        _ => "single",
    };
    public static string Width(WidthType t) => t == WidthType.Pct ? "pct" : "abs";
    public static string Emphasis(EmphasisMark m) => m switch
    {
        EmphasisMark.Comma => "comma",
        EmphasisMark.Circle => "circle",
        EmphasisMark.UnderDot => "underDot",
        _ => "dot",
    };
    public static string VAlign(CellVAlign v) => v switch
    {
        CellVAlign.Center => "center",
        CellVAlign.Bottom => "bottom",
        _ => "top",
    };
    public static string TextVAlign(LineVAlign v) => v switch
    {
        LineVAlign.Top => "top",
        LineVAlign.Center => "center",
        LineVAlign.Bottom => "bottom",
        _ => "baseline",
    };
    public static string TextDir(CellTextDirection d) => d switch
    {
        CellTextDirection.TbRl => "tbRl",
        CellTextDirection.BtLr => "btLr",
        CellTextDirection.LrTbV => "lrTbV",
        CellTextDirection.TbRlV => "tbRlV",
        CellTextDirection.TbLrV => "tbLrV",
        _ => "lrTb",
    };
    public static string Overlap(TableOverlap o) => o == TableOverlap.Never ? "never" : "overlap";
    public static string ListFmt(ListNumberFormat f) => f switch
    {
        ListNumberFormat.Decimal => "decimal",
        ListNumberFormat.LowerLetter => "lowerLetter",
        ListNumberFormat.UpperLetter => "upperLetter",
        ListNumberFormat.LowerRoman => "lowerRoman",
        ListNumberFormat.UpperRoman => "upperRoman",
        _ => "bullet",
    };
    public static string Cond(TableCond c) => c switch
    {
        TableCond.WholeTable => "wholeTable",
        TableCond.FirstRow => "firstRow",
        TableCond.LastRow => "lastRow",
        TableCond.FirstCol => "firstCol",
        TableCond.LastCol => "lastCol",
        TableCond.Band1Horz => "band1Horz",
        TableCond.Band2Horz => "band2Horz",
        TableCond.Band1Vert => "band1Vert",
        TableCond.Band2Vert => "band2Vert",
        TableCond.NwCell => "nwCell",
        TableCond.NeCell => "neCell",
        TableCond.SwCell => "swCell",
        _ => "seCell",
    };
}

/// <summary>A custom page size in px @96dpi (alternative to a named size).</summary>
public sealed record PageSize(double WidthPx, double HeightPx);

/// <summary>Per-side margins in px @96dpi; null sides keep their current value.</summary>
public sealed record Margins
{
    public double? Top { get; init; }
    public double? Right { get; init; }
    public double? Bottom { get; init; }
    public double? Left { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (Top is { } t) Js.Set(o, "top", t);
        if (Right is { } r) Js.Set(o, "right", r);
        if (Bottom is { } b) Js.Set(o, "bottom", b);
        if (Left is { } l) Js.Set(o, "left", l);
        return o;
    }
}

/// <summary>Character formatting patch (only set fields are applied).</summary>
public sealed record CharStyle
{
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public bool? Underline { get; init; }
    /// <summary>Underline line style (OOXML w:u/@w:val): "double", "dotted", "dash",
    /// "dotDash", "dotDotDash", "wave", "thick"; null/"single" = a plain line.</summary>
    public string? UnderlineStyle { get; init; }
    /// <summary>Explicit underline color (CSS hex). Null = follow the text color.</summary>
    public string? UnderlineColor { get; init; }
    public bool? Strikethrough { get; init; }
    public string? Color { get; init; }
    public string? HighlightColor { get; init; }
    /// <summary>Explicitly clear the run highlight (OOXML &lt;w:highlight w:val="none"/&gt;) — overrides
    /// a highlight the run's character style carries so the clear survives round-trips (issue #155).
    /// Ignored if HighlightColor is set.</summary>
    public bool? HighlightCleared { get; init; }
    public string? FontFamily { get; init; }
    /// <summary>Complex-script (bidi) font (OOXML w:rFonts/@w:cs); preserved through round-trip.</summary>
    public string? FontFamilyComplexScript { get; init; }
    /// <summary>East-Asian (CJK) font (OOXML w:rFonts/@w:eastAsia); preserved through round-trip.</summary>
    public string? FontFamilyEastAsia { get; init; }
    /// <summary>Proofing/typographic language (OOXML w:lang); round-trip only.</summary>
    public RunLangOptions? Lang { get; init; }
    public double? FontSizePx { get; init; }
    public string? Link { get; init; }
    public double? LetterSpacingPx { get; init; }
    public bool? Hidden { get; init; }
    /// <summary>Sub/superscript: "sub" or "super" (measured at 0.65× size, baseline-shifted).</summary>
    public string? VerticalAlign { get; init; }
    /// <summary>All-caps display (OOXML w:caps): every letter renders uppercased.</summary>
    public bool? Caps { get; init; }
    /// <summary>Small-capitals display (OOXML w:smallCaps): uppercased, originally-lowercase letters drawn smaller.</summary>
    public bool? SmallCaps { get; init; }
    /// <summary>Explicit right-to-left run (OOXML w:rtl).</summary>
    public bool? Rtl { get; init; }
    /// <summary>Character-style reference (OOXML w:rStyle → a character NamedStyle).</summary>
    public string? CharStyleId { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (VerticalAlign is { } va) Js.Set(o, "verticalAlign", va);
        if (Bold is { } b) Js.Set(o, "bold", b);
        if (Italic is { } i) Js.Set(o, "italic", i);
        if (Underline is { } u) Js.Set(o, "underline", u);
        if (UnderlineStyle is { } us) Js.Set(o, "underlineStyle", us);
        if (UnderlineColor is { } uc) Js.Set(o, "underlineColor", uc);
        if (Strikethrough is { } s) Js.Set(o, "strikethrough", s);
        if (Color is { } c) Js.Set(o, "color", c);
        if (HighlightColor is { } h) Js.Set(o, "highlightColor", h);
        if (HighlightCleared is { } hc) Js.Set(o, "highlightCleared", hc);
        if (FontFamily is { } f) Js.Set(o, "fontFamily", f);
        if (FontFamilyComplexScript is { } fcs) Js.Set(o, "fontFamilyComplexScript", fcs);
        if (FontFamilyEastAsia is { } fea) Js.Set(o, "fontFamilyEastAsia", fea);
        if (Lang is { } lang) Js.Set(o, "lang", lang.ToJs(e));
        if (FontSizePx is { } fs) Js.Set(o, "fontSizePx", fs);
        if (Link is { } lk) Js.Set(o, "link", lk);
        if (LetterSpacingPx is { } ls) Js.Set(o, "letterSpacingPx", ls);
        if (Hidden is { } hd) Js.Set(o, "hidden", hd);
        if (Caps is { } cp) Js.Set(o, "caps", cp);
        if (SmallCaps is { } sc) Js.Set(o, "smallCaps", sc);
        if (Rtl is { } rtl) Js.Set(o, "rtl", rtl);
        if (CharStyleId is { } csi) Js.Set(o, "charStyleId", csi);
        return o;
    }
}

/// <summary>Minor run typography &amp; effects (OOXML w:rPr extras) for
/// ParagraphBuilder.Effects(): double strikethrough, baseline raise/lower,
/// character width scaling, kerning threshold, emphasis marks, the
/// outline/shadow/emboss/imprint text effects, document-grid snapping, a run
/// border, and fitText. Only the set fields are applied (additive).</summary>
public sealed record RunEffectsOptions
{
    /// <summary>Double strikethrough (w:dstrike).</summary>
    public bool? DoubleStrikethrough { get; init; }
    /// <summary>Baseline raise (+) / lower (−) in px (w:position).</summary>
    public double? PositionPx { get; init; }
    /// <summary>Character width scaling as a percentage (w:w; 100 = normal).</summary>
    public double? WidthScalePct { get; init; }
    /// <summary>Kerning threshold in px — the min font size kerning applies at (w:kern).</summary>
    public double? KerningMinPx { get; init; }
    /// <summary>Emphasis marks over each character (w:em).</summary>
    public EmphasisMark? EmphasisMark { get; init; }
    /// <summary>Outlined text effect (w:outline).</summary>
    public bool? Outline { get; init; }
    /// <summary>Drop-shadow text effect (w:shadow).</summary>
    public bool? Shadow { get; init; }
    /// <summary>Embossed text effect (w:emboss).</summary>
    public bool? Emboss { get; init; }
    /// <summary>Imprinted/engraved text effect (w:imprint).</summary>
    public bool? Imprint { get; init; }
    /// <summary>Snap glyph advances to the document grid (w:snapToGrid; Word default ON).
    /// Round-trip only — pass false to opt a run out.</summary>
    public bool? SnapToGrid { get; init; }
    /// <summary>A box drawn around the run (w:bdr).</summary>
    public CellBorder? Border { get; init; }
    /// <summary>Fit-text target width in px (w:fitText).</summary>
    public double? FitTextPx { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (DoubleStrikethrough is { } ds) Js.Set(o, "doubleStrikethrough", ds);
        if (PositionPx is { } pos) Js.Set(o, "positionPx", pos);
        if (WidthScalePct is { } ws)
        {
            // OOXML w:w (and the round-trip) only preserve a positive percentage in
            // 1..600; reject out-of-range values rather than author dropped state.
            if (ws <= 0 || ws > 600) throw new ArgumentOutOfRangeException(nameof(WidthScalePct), ws, "widthScalePct must be in the range 1..600");
            Js.Set(o, "widthScalePct", ws);
        }
        if (KerningMinPx is { } km) Js.Set(o, "kerningMinPx", km);
        if (EmphasisMark is { } em) Js.Set(o, "emphasisMark", EnumJs.Emphasis(em));
        if (Outline is { } ol) Js.Set(o, "outline", ol);
        if (Shadow is { } sh) Js.Set(o, "shadow", sh);
        if (Emboss is { } eb) Js.Set(o, "emboss", eb);
        if (Imprint is { } im) Js.Set(o, "imprint", im);
        if (SnapToGrid is { } sg) Js.Set(o, "snapToGrid", sg);
        if (Border is { } bd) Js.Set(o, "border", bd.ToJs(e));
        if (FitTextPx is { } ft)
        {
            // w:fitText is a positive width; a non-positive value would not round-trip.
            if (ft <= 0) throw new ArgumentOutOfRangeException(nameof(FitTextPx), ft, "fitTextPx must be greater than 0");
            Js.Set(o, "fitTextPx", ft);
        }
        return o;
    }
}

/// <summary>Proofing/typographic language (OOXML w:lang). Each field is a BCP-47 tag
/// (e.g. "en-US"); only the set fields are emitted. Round-trip only — no layout effect.</summary>
public sealed record RunLangOptions
{
    /// <summary>w:lang/@w:val — Latin/Western language.</summary>
    public string? Val { get; init; }
    /// <summary>w:lang/@w:eastAsia — East-Asian language.</summary>
    public string? EastAsia { get; init; }
    /// <summary>w:lang/@w:bidi — complex-script (bidi) language.</summary>
    public string? Bidi { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (Val is { } v) Js.Set(o, "val", v);
        if (EastAsia is { } ea) Js.Set(o, "eastAsia", ea);
        if (Bidi is { } bd) Js.Set(o, "bidi", bd);
        return o;
    }
}

public sealed record CreateOptions
{
    public PageSizeName? PageSize { get; init; }
    public PageSize? CustomPageSize { get; init; }
    public Margins? Margins { get; init; }
    public string? IdSeed { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        SpecJs.SetPageSize(e, o, PageSize, CustomPageSize);
        if (Margins is { } m) Js.Set(o, "margins", m.ToJs(e));
        if (IdSeed is { } seed) Js.Set(o, "idSeed", seed);
        return o;
    }
}

public sealed record TemplateOptions
{
    public bool? KeepBody { get; init; }
    public string? IdSeed { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (KeepBody is { } k) Js.Set(o, "keepBody", k);
        if (IdSeed is { } seed) Js.Set(o, "idSeed", seed);
        return o;
    }
}

public sealed record ColumnsSpec(int Count, double? GapPx = null);

/// <summary>Line numbering in the page margin (OOXML w:lnNumType). Only set fields
/// are emitted; absent fields take Word's defaults (countBy 1, start 1, newPage).</summary>
public sealed record LineNumbering
{
    /// <summary>Print every Nth line number (default 1 = every line).</summary>
    public int? CountBy { get; init; }
    /// <summary>The first line's number (default 1).</summary>
    public int? Start { get; init; }
    /// <summary>When the counter restarts.</summary>
    public LineNumberRestart? Restart { get; init; }
    /// <summary>Gap (px) between the numbers and the text edge.</summary>
    public double? DistancePx { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (CountBy is { } cb) Js.Set(o, "countBy", cb);
        if (Start is { } st) Js.Set(o, "start", st);
        if (Restart is { } r) Js.Set(o, "restart", EnumJs.LineRestart(r));
        if (DistancePx is { } d) Js.Set(o, "distancePx", d);
        return o;
    }
}

public sealed record PageSetup
{
    public PageSizeName? PageSize { get; init; }
    public PageSize? CustomPageSize { get; init; }
    public Orientation? Orientation { get; init; }
    public Margins? Margins { get; init; }
    public ColumnsSpec? Columns { get; init; }
    public double? HeaderDistancePx { get; init; }
    public double? FooterDistancePx { get; init; }
    public int? PageNumberStart { get; init; }
    /// <summary>Line numbering in the margin (w:lnNumType) for the document section.</summary>
    public LineNumbering? LineNumbering { get; init; }
    /// <summary>How the final (body) section starts: next page (default) or even/odd parity.</summary>
    public SectionBreakType? BreakType { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        SpecJs.SetPageSize(e, o, PageSize, CustomPageSize);
        if (Orientation is { } orient) Js.Set(o, "orientation", EnumJs.Orient(orient));
        if (Margins is { } m) Js.Set(o, "margins", m.ToJs(e));
        if (Columns is { } col) Js.Set(o, "columns", SpecJs.Columns(e, col));
        if (HeaderDistancePx is { } hd) Js.Set(o, "headerDistancePx", hd);
        if (FooterDistancePx is { } fd) Js.Set(o, "footerDistancePx", fd);
        if (PageNumberStart is { } pns) Js.Set(o, "pageNumberStart", pns);
        if (LineNumbering is { } ln) Js.Set(o, "lineNumbering", ln.ToJs(e));
        if (BreakType is { } bt) Js.Set(o, "breakType", EnumJs.BreakType(bt));
        return o;
    }
}

public sealed record SpacingOptions
{
    public double? Before { get; init; }
    public double? After { get; init; }
    /// <summary>Line height multiplier (1.0 = single). Used when <see cref="LineRule"/> is omitted.</summary>
    public double? LineHeight { get; init; }
    /// <summary>Fixed line-spacing rule (docx w:lineRule="exact"|"atLeast"). Set <see cref="LineHeightPx"/> alongside it.</summary>
    public LineRule? LineRule { get; init; }
    /// <summary>Fixed line height in px — used with <see cref="LineRule"/>.</summary>
    public double? LineHeightPx { get; init; }
    /// <summary>Word's "automatic" space before/after (OOXML w:beforeAutospacing / w:afterAutospacing,
    /// issue #160). When true, bakes the approximate auto value and re-emits the attribute; overrides
    /// <see cref="Before"/>/<see cref="After"/> respectively.</summary>
    public bool? BeforeAuto { get; init; }
    public bool? AfterAuto { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (Before is { } b) Js.Set(o, "before", b);
        if (After is { } a) Js.Set(o, "after", a);
        if (LineHeight is { } lh) Js.Set(o, "lineHeight", lh);
        if (LineRule is { } lr) Js.Set(o, "lineRule", EnumJs.LineRule(lr));
        if (LineHeightPx is { } lhp) Js.Set(o, "lineHeightPx", lhp);
        if (BeforeAuto is { } ba) Js.Set(o, "beforeAuto", ba);
        if (AfterAuto is { } aa) Js.Set(o, "afterAuto", aa);
        return o;
    }
}

public sealed record IndentOptions
{
    public double? Left { get; init; }
    public double? Right { get; init; }
    public double? FirstLine { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (Left is { } l) Js.Set(o, "left", l);
        if (Right is { } r) Js.Set(o, "right", r);
        if (FirstLine is { } fl) Js.Set(o, "firstLine", fl);
        return o;
    }
}

public sealed record ImageOptions
{
    public required double WidthPx { get; init; }
    public required double HeightPx { get; init; }
    public TextAlign? Align { get; init; }
    public ImageWrap? Wrap { get; init; }
    /// <summary>Crop insets (OOXML a:srcRect), each a 0..1 fraction trimmed off that
    /// edge — so WidthPx/HeightPx describe the cropped box. Null = no crop.</summary>
    public ImageCrop? Crop { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        Js.Set(o, "widthPx", WidthPx);
        Js.Set(o, "heightPx", HeightPx);
        if (Align is { } a)
        {
            if (a == TextAlign.Justify)
                throw new ArgumentOutOfRangeException(nameof(Align), "Image align supports left/center/right only (not Justify).");
            Js.Set(o, "align", EnumJs.Align(a));
        }
        if (Wrap is { } w) Js.Set(o, "wrap", EnumJs.Wrap(w));
        if (Crop is { } cr) Js.Set(o, "crop", cr.ToJs(e));
        return o;
    }
}

/// <summary>Image crop insets (OOXML a:srcRect): 0..1 fractions trimmed off each edge.</summary>
public sealed record ImageCrop(double Left, double Top, double Right, double Bottom)
{
    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        Js.Set(o, "left", Left);
        Js.Set(o, "top", Top);
        Js.Set(o, "right", Right);
        Js.Set(o, "bottom", Bottom);
        return o;
    }
}

public sealed record ListItem(string Text, int? Level = null, CharStyle? Style = null)
{
    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        Js.Set(o, "text", Text);
        if (Level is { } lvl) Js.Set(o, "level", lvl);
        if (Style is { } st) Js.Set(o, "style", st.ToJs(e));
        return o;
    }
}

public sealed record ListOptions
{
    public ListKind? Kind { get; init; }
    public string? ListId { get; init; }
    public int? Level { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (Kind is { } k) Js.Set(o, "kind", k == ListKind.Number ? "number" : "bullet");
        if (ListId is { } id) Js.Set(o, "listId", id);
        if (Level is { } lvl) Js.Set(o, "level", lvl);
        return o;
    }
}

public sealed record TableOptions
{
    public IReadOnlyList<double>? ColFractions { get; init; }
    public bool? HeaderRow { get; init; }
    /// <summary>Apply a registered builder-only table-style PRESET by name (baked into
    /// cells; nothing style-id-like enters the model). See DocumentBuilder.TableStylePreset.</summary>
    public string? Style { get; init; }
    /// <summary>Reference a REAL table style registered via DocumentBuilder.TableStyle()
    /// and bake its effective per-cell formatting onto the cells (kept for round-trip).</summary>
    public string? StyleId { get; init; }
    /// <summary>Which conditional bands of the referenced StyleId are active (w:tblLook).</summary>
    public TableCondOverrides? CondOverrides { get; init; }
    /// <summary>Column sizing strategy (w:tblLayout).</summary>
    public TableWidthMode? WidthMode { get; init; }
    /// <summary>Table-level default borders (w:tblPr/w:tblBorders), incl. interior
    /// edges; re-emitted at tblPr level and cascaded onto cells (issue #48).</summary>
    public TableBorders? Borders { get; init; }
    /// <summary>Table-level default shading fill (w:tblPr/w:shd); cells may override.</summary>
    public string? Shading { get; init; }
    /// <summary>Table-level default cell margins (w:tblPr/w:tblCellMar).</summary>
    public CellMargin? CellMargin { get; init; }
    /// <summary>Table indent from the leading content edge (w:tblPr/w:tblInd), in px.</summary>
    public double? Indent { get; init; }
    /// <summary>Render columns right-to-left (w:tblPr/w:bidiVisual).</summary>
    public bool? BidiVisual { get; init; }
    /// <summary>Floating-table overlap behavior (w:tblPr/w:tblOverlap).</summary>
    public TableOverlap? Overlap { get; init; }
    /// <summary>Table caption / title (w:tblPr/w:tblCaption) — accessibility metadata.</summary>
    public string? Caption { get; init; }
    /// <summary>Table description / alt text (w:tblPr/w:tblDescription) — accessibility metadata.</summary>
    public string? Description { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (ColFractions is { } cf) Js.Set(o, "colFractions", Js.ToArray(e, cf.Select(x => (object?)x)));
        if (HeaderRow is { } hr) Js.Set(o, "headerRow", hr);
        if (Style is { } s) Js.Set(o, "style", s);
        if (StyleId is { } sid) Js.Set(o, "styleId", sid);
        if (CondOverrides is { } co) Js.Set(o, "condOverrides", co.ToJs(e));
        if (WidthMode is { } wm) Js.Set(o, "widthMode", EnumJs.WidthMode(wm));
        if (Borders is { } bd) Js.Set(o, "borders", bd.ToJs(e));
        if (Shading is { } sh) Js.Set(o, "shading", sh);
        if (CellMargin is { } cm) Js.Set(o, "cellMargin", cm.ToJs(e));
        if (Indent is { } ind) Js.Set(o, "indent", ind);
        if (BidiVisual is { } bv) Js.Set(o, "bidiVisual", bv);
        if (Overlap is { } ov) Js.Set(o, "overlap", EnumJs.Overlap(ov));
        if (Caption is { } cap) Js.Set(o, "caption", cap);
        if (Description is { } desc) Js.Set(o, "description", desc);
        return o;
    }
}

public sealed record CellSpec
{
    public string? Text { get; init; }
    public int? ColSpan { get; init; }
    public int? RowSpan { get; init; }
    public string? Shading { get; init; }
    /// <summary>Explicitly clear shading (OOXML &lt;w:shd w:val="clear" w:fill="auto"/&gt;,
    /// Word's "No Color") — overrides a fill inherited from the table's default shading or
    /// table style so the clear survives round-trips (issue #150). Ignored if Shading is set.</summary>
    public bool? ShadingCleared { get; init; }
    public CharStyle? Style { get; init; }
    public TextAlign? Align { get; init; }
    /// <summary>Per-edge borders; absent = the renderer's default light grid.</summary>
    public CellBorders? Borders { get; init; }
    /// <summary>Inner padding override, px per side.</summary>
    public CellMargin? Margin { get; init; }
    /// <summary>Preferred cell width (w:tcW).</summary>
    public PreferredWidth? PreferredWidth { get; init; }
    /// <summary>Vertical content alignment (w:vAlign). Absent = top.</summary>
    public CellVAlign? VAlign { get; init; }
    /// <summary>Text flow direction (w:textDirection). Absent = lrTb (horizontal).</summary>
    public CellTextDirection? TextDirection { get; init; }
    /// <summary>Suppress content wrapping (w:noWrap).</summary>
    public bool? NoWrap { get; init; }
    /// <summary>Fit text to the cell width (w:tcFitText).</summary>
    public bool? FitText { get; init; }
    /// <summary>Ignore the end-of-cell mark for row height (w:hideMark).</summary>
    public bool? HideMark { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (Text is { } t) Js.Set(o, "text", t);
        if (ColSpan is { } cs) Js.Set(o, "colSpan", cs);
        if (RowSpan is { } rs) Js.Set(o, "rowSpan", rs);
        if (Shading is { } sh) Js.Set(o, "shading", sh);
        if (ShadingCleared is { } shc) Js.Set(o, "shadingCleared", shc);
        if (Style is { } st) Js.Set(o, "style", st.ToJs(e));
        if (Align is { } a) Js.Set(o, "align", EnumJs.Align(a));
        if (Borders is { } bd) Js.Set(o, "borders", bd.ToJs(e));
        if (Margin is { } mg) Js.Set(o, "margin", mg.ToJs(e));
        if (PreferredWidth is { } pw) Js.Set(o, "preferredWidth", pw.ToJs(e));
        if (VAlign is { } va) Js.Set(o, "vAlign", EnumJs.VAlign(va));
        if (TextDirection is { } td) Js.Set(o, "textDirection", EnumJs.TextDir(td));
        if (NoWrap is { } nw) Js.Set(o, "noWrap", nw);
        if (FitText is { } ft) Js.Set(o, "fitText", ft);
        if (HideMark is { } hm) Js.Set(o, "hideMark", hm);
        return o;
    }
}

/// <summary>Cell formatting options (a <see cref="CellSpec"/> without the text — the
/// content is supplied separately as the first argument to <c>Cell()</c>), matching the
/// JS <c>CellOptions = Omit&lt;CellSpec, "text"&gt;</c> contract.</summary>
public sealed record CellOptions
{
    public int? ColSpan { get; init; }
    public int? RowSpan { get; init; }
    public string? Shading { get; init; }
    /// <summary>Explicitly clear shading (OOXML &lt;w:shd w:val="clear" w:fill="auto"/&gt;,
    /// Word's "No Color") — overrides a fill inherited from the table's default shading or
    /// table style so the clear survives round-trips (issue #150). Ignored if Shading is set.</summary>
    public bool? ShadingCleared { get; init; }
    public CharStyle? Style { get; init; }
    public TextAlign? Align { get; init; }
    public CellBorders? Borders { get; init; }
    public CellMargin? Margin { get; init; }
    public PreferredWidth? PreferredWidth { get; init; }
    /// <summary>Vertical content alignment (w:vAlign). Absent = top.</summary>
    public CellVAlign? VAlign { get; init; }
    /// <summary>Text flow direction (w:textDirection). Absent = lrTb (horizontal).</summary>
    public CellTextDirection? TextDirection { get; init; }
    /// <summary>Suppress content wrapping (w:noWrap).</summary>
    public bool? NoWrap { get; init; }
    /// <summary>Fit text to the cell width (w:tcFitText).</summary>
    public bool? FitText { get; init; }
    /// <summary>Ignore the end-of-cell mark for row height (w:hideMark).</summary>
    public bool? HideMark { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (ColSpan is { } cs) Js.Set(o, "colSpan", cs);
        if (RowSpan is { } rs) Js.Set(o, "rowSpan", rs);
        if (Shading is { } sh) Js.Set(o, "shading", sh);
        if (ShadingCleared is { } shc) Js.Set(o, "shadingCleared", shc);
        if (Style is { } st) Js.Set(o, "style", st.ToJs(e));
        if (Align is { } a) Js.Set(o, "align", EnumJs.Align(a));
        if (Borders is { } bd) Js.Set(o, "borders", bd.ToJs(e));
        if (Margin is { } mg) Js.Set(o, "margin", mg.ToJs(e));
        if (PreferredWidth is { } pw) Js.Set(o, "preferredWidth", pw.ToJs(e));
        if (VAlign is { } va) Js.Set(o, "vAlign", EnumJs.VAlign(va));
        if (TextDirection is { } td) Js.Set(o, "textDirection", EnumJs.TextDir(td));
        if (NoWrap is { } nw) Js.Set(o, "noWrap", nw);
        if (FitText is { } ft) Js.Set(o, "fitText", ft);
        if (HideMark is { } hm) Js.Set(o, "hideMark", hm);
        return o;
    }
}

/// <summary>Common content-control (SDT) options.</summary>
public sealed record SdtOptions
{
    /// <summary>Friendly label shown in Word's control chrome.</summary>
    public string? Alias { get; init; }
    /// <summary>Programmatic tag (machine identifier).</summary>
    public string? Tag { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (Alias is { } a) Js.Set(o, "alias", a);
        if (Tag is { } t) Js.Set(o, "tag", t);
        return o;
    }
}

/// <summary>A drop-down / combo-box list item (display text + stored value).</summary>
public sealed record SdtListItem(string Display, string Value)
{
    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        Js.Set(o, "display", Display);
        Js.Set(o, "value", Value);
        return o;
    }
}

/// <summary>Row-level properties (w:trPr) for a single <c>Row()</c> — fixed/min
/// height, cant-split, and repeat-header. Mirrors the JS <c>RowOptions</c>.</summary>
public sealed record RowOptions
{
    /// <summary>Fixed/minimum row height in px (w:trHeight). Pair with <see cref="HeightRule"/>.</summary>
    public double? Height { get; init; }
    /// <summary>How <see cref="Height"/> is enforced: AtLeast (min, grows with content — the
    /// default) or Exact (pinned; taller content is clipped).</summary>
    public RowHeightRule? HeightRule { get; init; }
    /// <summary>Keep the whole row on one page — never split across a page/column break (w:cantSplit).</summary>
    public bool? CantSplit { get; init; }
    /// <summary>Repeat this row as a header at the top of each page the table continues onto (w:tblHeader).</summary>
    public bool? Header { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (Height is { } h) Js.Set(o, "height", h);
        if (HeightRule is { } r) Js.Set(o, "heightRule", r == RowHeightRule.Exact ? "exact" : "atLeast");
        if (CantSplit is { } cs) Js.Set(o, "cantSplit", cs);
        if (Header is { } hd) Js.Set(o, "header", hd);
        return o;
    }
}

/// <summary>String or rich cell spec — what a table cell accepts.</summary>
public readonly struct CellContent
{
    private readonly string? _text;
    private readonly CellSpec? _spec;
    private CellContent(string? text, CellSpec? spec) { _text = text; _spec = spec; }

    public static implicit operator CellContent(string text) => new(text, null);
    public static implicit operator CellContent(CellSpec spec) => new(null, spec);

    internal object ToJs(WordCanvasEngine e) => _spec is not null ? _spec.ToJs(e) : (object)(_text ?? "");
}

/// <summary>A named-style definition for DocumentBuilder.Style(). Only set fields
/// are emitted; Para/Char are formatting patches.</summary>
public sealed record NamedStyle
{
    public required string Id { get; init; }
    public required string Name { get; init; }
    /// <summary>"paragraph" | "character" (default "paragraph").</summary>
    public string? Type { get; init; }
    public string? BasedOn { get; init; }
    public string? Next { get; init; }
    public CharStyle? Char { get; init; }
    public ParaStylePatch? Para { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        Js.Set(o, "id", Id);
        Js.Set(o, "name", Name);
        Js.Set(o, "type", Type ?? "paragraph");
        if (BasedOn is { } b) Js.Set(o, "basedOn", b);
        if (Next is { } n) Js.Set(o, "next", n);
        // char/para are REQUIRED (Partial) on the model — always emit, defaulting to
        // empty objects, so the stylesheet cascade never sees undefined.
        Js.Set(o, "char", (Char ?? new CharStyle()).ToJs(e));
        Js.Set(o, "para", (Para ?? new ParaStylePatch()).ToJs(e));
        return o;
    }
}

/// <summary>Paragraph formatting patch (used by NamedStyle.Para and table conditions).</summary>
public sealed record ParaStylePatch
{
    public TextAlign? Align { get; init; }
    public double? LineHeight { get; init; }
    public double? SpaceBeforePx { get; init; }
    public double? SpaceAfterPx { get; init; }
    public double? IndentLeftPx { get; init; }
    public double? IndentRightPx { get; init; }
    public double? IndentFirstLinePx { get; init; }
    public bool? KeepWithNext { get; init; }
    /// <summary>Never split this paragraph across pages/columns (docx w:keepLines).</summary>
    public bool? KeepLinesTogether { get; init; }
    /// <summary>Suppress before/after spacing between adjacent same-style paragraphs (docx w:contextualSpacing).</summary>
    public bool? ContextualSpacing { get; init; }
    /// <summary>Base writing direction (OOXML w:bidi).</summary>
    public Direction? Direction { get; init; }
    /// <summary>Outline level 0..8 (TOC levels 1..9; docx w:outlineLvl) — heading styles carry this.</summary>
    public int? OutlineLevel { get; init; }
    public bool? PageBreakBefore { get; init; }
    public bool? ColumnBreakBefore { get; init; }
    /// <summary>Explicit tab stops (docx w:tabs).</summary>
    public IReadOnlyList<TabStop>? TabStops { get; init; }
    /// <summary>Paragraph shading fill (OOXML w:shd) — a CSS color painted behind the paragraph.</summary>
    public string? Shading { get; init; }
    /// <summary>Explicitly clear paragraph shading (OOXML &lt;w:shd w:val="clear" w:fill="auto"/&gt;,
    /// Word's "No Color") — overrides a fill an inherited style carries so the clear survives
    /// export/re-import instead of the style's fill silently returning (issue #147).</summary>
    public bool? ShadingCleared { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (Align is { } a) Js.Set(o, "align", EnumJs.Align(a));
        if (LineHeight is { } lh) Js.Set(o, "lineHeight", lh);
        if (SpaceBeforePx is { } sb) Js.Set(o, "spaceBeforePx", sb);
        if (SpaceAfterPx is { } sa) Js.Set(o, "spaceAfterPx", sa);
        if (IndentLeftPx is { } il) Js.Set(o, "indentLeftPx", il);
        if (IndentRightPx is { } ir) Js.Set(o, "indentRightPx", ir);
        if (IndentFirstLinePx is { } ifl) Js.Set(o, "indentFirstLinePx", ifl);
        if (KeepWithNext is { } kwn) Js.Set(o, "keepWithNext", kwn);
        if (KeepLinesTogether is { } klt) Js.Set(o, "keepLinesTogether", klt);
        if (ContextualSpacing is { } cs) Js.Set(o, "contextualSpacing", cs);
        if (Direction is { } dir) Js.Set(o, "direction", EnumJs.Dir(dir));
        if (OutlineLevel is { } ol) Js.Set(o, "outlineLevel", Math.Clamp(ol, 0, 8));
        if (PageBreakBefore is { } pbb) Js.Set(o, "pageBreakBefore", pbb);
        if (ColumnBreakBefore is { } cbb) Js.Set(o, "columnBreakBefore", cbb);
        if (TabStops is { } ts) Js.Set(o, "tabStops", Js.ToArray(e, ts.Select(t => (object?)t.ToJs(e))));
        if (Shading is { } shd) Js.Set(o, "shading", shd);
        if (ShadingCleared is { } shc) Js.Set(o, "shadingCleared", shc);
        return o;
    }
}

/// <summary>Table-of-contents options (builder + headless TOC recalc).</summary>
public sealed record TocOptions
{
    public int? MaxLevel { get; init; }
    public bool? Hyperlink { get; init; }
    /// <summary>Show trailing page numbers in entries (default true). Used by the
    /// headless TOC generator.</summary>
    public bool? IncludePageNumbers { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (MaxLevel is { } ml) Js.Set(o, "maxLevel", ml);
        if (Hyperlink is { } h) Js.Set(o, "hyperlink", h);
        if (IncludePageNumbers is { } ipn) Js.Set(o, "includePageNumbers", ipn);
        return o;
    }
}

/// <summary>Block/inline display-equation options.</summary>
public sealed record EquationOptions
{
    public EquationAlign? Align { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (Align is { } a) Js.Set(o, "align", EnumJs.EqAlign(a));
        return o;
    }
}

/// <summary>An explicit tab stop (docx w:tab).</summary>
public sealed record TabStop
{
    public required double PosPx { get; init; }
    public TabAlign? Align { get; init; }
    public TabLeader? Leader { get; init; }
    /// <summary>Explicit clear (OOXML &lt;w:tab w:val="clear" w:pos="…"/&gt;) — removes an inherited
    /// tab stop at this position (issue #154). Not a real stop; when set, Align/Leader are ignored.</summary>
    public bool? Cleared { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        Js.Set(o, "posPx", PosPx);
        if (Align is { } a) Js.Set(o, "align", EnumJs.TabAlign(a));
        if (Leader is { } l) Js.Set(o, "leader", EnumJs.TabLeader(l));
        if (Cleared is { } cl) Js.Set(o, "cleared", cl);
        return o;
    }
}

/// <summary>One cell edge's border (color + width, optional line style).</summary>
public sealed record CellBorder
{
    public required string Color { get; init; }
    public required double WidthPx { get; init; }
    public BorderStyle? Style { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        Js.Set(o, "color", Color);
        Js.Set(o, "widthPx", WidthPx);
        if (Style is { } s) Js.Set(o, "style", EnumJs.Border(s));
        return o;
    }
}

/// <summary>Per-edge cell borders; absent edges keep the renderer's default grid.</summary>
public sealed record CellBorders
{
    public CellBorder? Top { get; init; }
    public CellBorder? Right { get; init; }
    public CellBorder? Bottom { get; init; }
    public CellBorder? Left { get; init; }

    /// <summary>The same border on all four edges.</summary>
    public static CellBorders All(CellBorder border) =>
        new() { Top = border, Right = border, Bottom = border, Left = border };

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (Top is { } t) Js.Set(o, "top", t.ToJs(e));
        if (Right is { } r) Js.Set(o, "right", r.ToJs(e));
        if (Bottom is { } b) Js.Set(o, "bottom", b.ToJs(e));
        if (Left is { } l) Js.Set(o, "left", l.ToJs(e));
        return o;
    }
}

/// <summary>Table-level default borders (w:tblPr/w:tblBorders): the four outer edges
/// plus the two INTERIOR edges Word draws between adjacent cells (insideH/insideV).</summary>
public sealed record TableBorders
{
    public CellBorder? Top { get; init; }
    public CellBorder? Right { get; init; }
    public CellBorder? Bottom { get; init; }
    public CellBorder? Left { get; init; }
    /// <summary>Interior horizontal edge between vertically-adjacent cells (w:insideH).</summary>
    public CellBorder? InsideH { get; init; }
    /// <summary>Interior vertical edge between horizontally-adjacent cells (w:insideV).</summary>
    public CellBorder? InsideV { get; init; }

    /// <summary>The same border on every edge, interior and exterior (a uniform grid).</summary>
    public static TableBorders All(CellBorder border) =>
        new() { Top = border, Right = border, Bottom = border, Left = border, InsideH = border, InsideV = border };

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (Top is { } t) Js.Set(o, "top", t.ToJs(e));
        if (Right is { } r) Js.Set(o, "right", r.ToJs(e));
        if (Bottom is { } b) Js.Set(o, "bottom", b.ToJs(e));
        if (Left is { } l) Js.Set(o, "left", l.ToJs(e));
        if (InsideH is { } ih) Js.Set(o, "insideH", ih.ToJs(e));
        if (InsideV is { } iv) Js.Set(o, "insideV", iv.ToJs(e));
        return o;
    }
}

/// <summary>Per-edge paragraph borders (OOXML w:pBdr); absent edges draw no line.
/// Reuses the table <see cref="CellBorder"/> value type. <c>Between</c> is the
/// inter-paragraph rule — it round-trips but is not drawn for a standalone paragraph.</summary>
public sealed record ParaBorders
{
    public CellBorder? Top { get; init; }
    public CellBorder? Right { get; init; }
    public CellBorder? Bottom { get; init; }
    public CellBorder? Left { get; init; }
    public CellBorder? Between { get; init; }

    /// <summary>The same border on all four outer edges.</summary>
    public static ParaBorders All(CellBorder border) =>
        new() { Top = border, Right = border, Bottom = border, Left = border };

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (Top is { } t) Js.Set(o, "top", t.ToJs(e));
        if (Right is { } r) Js.Set(o, "right", r.ToJs(e));
        if (Bottom is { } b) Js.Set(o, "bottom", b.ToJs(e));
        if (Left is { } l) Js.Set(o, "left", l.ToJs(e));
        if (Between is { } bw) Js.Set(o, "between", bw.ToJs(e));
        return o;
    }
}

/// <summary>Inner cell padding, px per side (docx w:tcMar).</summary>
public sealed record CellMargin(double Top, double Right, double Bottom, double Left)
{
    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        Js.Set(o, "top", Top);
        Js.Set(o, "right", Right);
        Js.Set(o, "bottom", Bottom);
        Js.Set(o, "left", Left);
        return o;
    }
}

/// <summary>Preferred cell width (w:tcW): absolute px or percent of table width.</summary>
public sealed record PreferredWidth(double Px, WidthType Type)
{
    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        Js.Set(o, "px", Px);
        Js.Set(o, "type", EnumJs.Width(Type));
        return o;
    }
}

/// <summary>Which conditional bands of a referenced table style are active (w:tblLook).</summary>
public sealed record TableCondOverrides
{
    public bool? FirstRow { get; init; }
    public bool? LastRow { get; init; }
    public bool? FirstCol { get; init; }
    public bool? LastCol { get; init; }
    public bool? BandRows { get; init; }
    public bool? BandCols { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (FirstRow is { } a) Js.Set(o, "firstRow", a);
        if (LastRow is { } b) Js.Set(o, "lastRow", b);
        if (FirstCol is { } c) Js.Set(o, "firstCol", c);
        if (LastCol is { } d) Js.Set(o, "lastCol", d);
        if (BandRows is { } f) Js.Set(o, "bandRows", f);
        if (BandCols is { } g) Js.Set(o, "bandCols", g);
        return o;
    }
}

/// <summary>Formatting for one conditional band of a table style.</summary>
public sealed record TableCondProps
{
    public CharStyle? Char { get; init; }
    public ParaStylePatch? Para { get; init; }
    public string? Shading { get; init; }
    public CellBorders? Borders { get; init; }
    public CellMargin? Margin { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (Char is { } c) Js.Set(o, "char", c.ToJs(e));
        if (Para is { } p) Js.Set(o, "para", p.ToJs(e));
        if (Shading is { } s) Js.Set(o, "shading", s);
        if (Borders is { } b) Js.Set(o, "borders", b.ToJs(e));
        if (Margin is { } m) Js.Set(o, "margin", m.ToJs(e));
        return o;
    }
}

/// <summary>A REAL table style (OOXML w:style[type=table]) for DocumentBuilder.TableStyle().
/// Conditional bands keyed by <see cref="TableCond"/>; `WholeTable` is the base.</summary>
public sealed record TableStyleDef
{
    public required string Id { get; init; }
    public required string Name { get; init; }
    public string? BasedOn { get; init; }
    public required IReadOnlyDictionary<TableCond, TableCondProps> Conds { get; init; }
    public int? RowBandSize { get; init; }
    public int? ColBandSize { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        Js.Set(o, "id", Id);
        Js.Set(o, "name", Name);
        if (BasedOn is { } b) Js.Set(o, "basedOn", b);
        var conds = Js.Obj(e);
        foreach (var kv in Conds) Js.Set(conds, EnumJs.Cond(kv.Key), kv.Value.ToJs(e));
        Js.Set(o, "conds", conds);
        if (RowBandSize is { } rb) Js.Set(o, "rowBandSize", rb);
        if (ColBandSize is { } cb) Js.Set(o, "colBandSize", cb);
        return o;
    }
}

/// <summary>One level of a custom list definition (docx w:lvl).</summary>
public sealed record ListLevel
{
    public required ListNumberFormat Format { get; init; }
    /// <summary>Marker pattern; %N is level N-1's counter (e.g. "%1."). Ignored for bullets.</summary>
    public required string Text { get; init; }
    public string? BulletChar { get; init; }
    public required double IndentLeftPx { get; init; }
    public required double HangingPx { get; init; }
    public required int Start { get; init; }
    public CharStyle? MarkerStyle { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        Js.Set(o, "format", EnumJs.ListFmt(Format));
        Js.Set(o, "text", Text);
        if (BulletChar is { } bc) Js.Set(o, "bulletChar", bc);
        Js.Set(o, "indentLeftPx", IndentLeftPx);
        Js.Set(o, "hangingPx", HangingPx);
        Js.Set(o, "start", Start);
        if (MarkerStyle is { } ms) Js.Set(o, "markerStyle", ms.ToJs(e));
        return o;
    }
}

/// <summary>How DocumentBuilder.ListDefinition() builds a custom list. Use the static
/// factories — they mirror the JS union (bullet / number / multilevel / explicit levels).</summary>
public sealed record ListDefinitionSpec
{
    private readonly Func<WordCanvasEngine, ScriptObject> _build;
    private ListDefinitionSpec(Func<WordCanvasEngine, ScriptObject> build) => _build = build;
    internal ScriptObject ToJs(WordCanvasEngine e) => _build(e);

    public static ListDefinitionSpec Bullet(string bulletChar) => new(e =>
    {
        var o = Js.Obj(e);
        Js.Set(o, "kind", "bullet");
        Js.Set(o, "char", bulletChar);
        return o;
    });

    public static ListDefinitionSpec Number(ListNumberFormat format, string? suffix = null) => new(e =>
    {
        var o = Js.Obj(e);
        Js.Set(o, "kind", "number");
        Js.Set(o, "format", EnumJs.ListFmt(format));
        if (suffix is { } s) Js.Set(o, "suffix", s);
        return o;
    });

    public static ListDefinitionSpec Multilevel() => new(e =>
    {
        var o = Js.Obj(e);
        Js.Set(o, "kind", "multilevel");
        return o;
    });

    public static ListDefinitionSpec Levels(IEnumerable<ListLevel> levels) => new(e =>
    {
        var o = Js.Obj(e);
        Js.Set(o, "levels", Js.ToArray(e, levels.Select(l => (object?)l.ToJs(e))));
        return o;
    });
}

/// <summary>A builder-only table-style PRESET (DocumentBuilder.TableStylePreset). Baked
/// into concrete cell formatting at build time; nothing style-id-like enters the model.</summary>
public sealed record TableStylePreset
{
    public bool? HeaderRow { get; init; }
    public CharStyle? HeaderChar { get; init; }
    public string? HeaderShading { get; init; }
    public CellBorders? Borders { get; init; }
    public string? Shading { get; init; }
    public string? StripeShading { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (HeaderRow is { } hr) Js.Set(o, "headerRow", hr);
        if (HeaderChar is { } hc) Js.Set(o, "headerChar", hc.ToJs(e));
        if (HeaderShading is { } hs) Js.Set(o, "headerShading", hs);
        if (Borders is { } b) Js.Set(o, "borders", b.ToJs(e));
        if (Shading is { } s) Js.Set(o, "shading", s);
        if (StripeShading is { } ss) Js.Set(o, "stripeShading", ss);
        return o;
    }
}

internal static class SpecJs
{
    public static void SetPageSize(WordCanvasEngine e, ScriptObject target, PageSizeName? name, PageSize? custom)
    {
        if (custom is { } c)
        {
            var o = Js.Obj(e);
            Js.Set(o, "pageWidthPx", c.WidthPx);
            Js.Set(o, "pageHeightPx", c.HeightPx);
            Js.Set(target, "pageSize", o);
        }
        else if (name is { } n)
        {
            Js.Set(target, "pageSize", EnumJs.PageSize(n));
        }
    }

    public static ScriptObject Columns(WordCanvasEngine e, ColumnsSpec col)
    {
        var o = Js.Obj(e);
        Js.Set(o, "count", col.Count);
        if (col.GapPx is { } g) Js.Set(o, "gapPx", g);
        return o;
    }
}
