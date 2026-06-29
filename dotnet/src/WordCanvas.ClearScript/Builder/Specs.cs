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
public enum ImageWrap { Block, Square }
public enum PageSizeName { Letter, Legal, A4, A3, Tabloid }
/// <summary>Block/inline equation horizontal placement (display equations only).</summary>
public enum EquationAlign { Left, Center, Right }
/// <summary>Paragraph base writing direction (OOXML w:bidi).</summary>
public enum Direction { Ltr, Rtl }
/// <summary>Tab-stop alignment (docx w:tab/@w:val).</summary>
public enum TabAlign { Left, Center, Right, Decimal }
/// <summary>Tab-stop leader fill (docx w:tab/@w:leader).</summary>
public enum TabLeader { None, Dot, Dash, Underscore }
/// <summary>Column sizing strategy (docx w:tblLayout + w:tblW).</summary>
public enum TableWidthMode { Fixed, AutofitContents, AutofitWindow }
/// <summary>Cell border line style.</summary>
public enum BorderStyle { Single, Double, Dashed, Dotted }
/// <summary>Preferred-width unit: absolute px or percent of table width.</summary>
public enum WidthType { Abs, Pct }
/// <summary>Cell vertical content alignment (OOXML w:tcPr/w:vAlign).</summary>
public enum CellVAlign { Top, Center, Bottom }
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
    public static string VAlign(CellVAlign v) => v switch
    {
        CellVAlign.Center => "center",
        CellVAlign.Bottom => "bottom",
        _ => "top",
    };
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
    public bool? Strikethrough { get; init; }
    public string? Color { get; init; }
    public string? HighlightColor { get; init; }
    public string? FontFamily { get; init; }
    public double? FontSizePx { get; init; }
    public string? Link { get; init; }
    public double? LetterSpacingPx { get; init; }
    public bool? Hidden { get; init; }
    /// <summary>Sub/superscript: "sub" or "super" (measured at 0.65× size, baseline-shifted).</summary>
    public string? VerticalAlign { get; init; }
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
        if (Strikethrough is { } s) Js.Set(o, "strikethrough", s);
        if (Color is { } c) Js.Set(o, "color", c);
        if (HighlightColor is { } h) Js.Set(o, "highlightColor", h);
        if (FontFamily is { } f) Js.Set(o, "fontFamily", f);
        if (FontSizePx is { } fs) Js.Set(o, "fontSizePx", fs);
        if (Link is { } lk) Js.Set(o, "link", lk);
        if (LetterSpacingPx is { } ls) Js.Set(o, "letterSpacingPx", ls);
        if (Hidden is { } hd) Js.Set(o, "hidden", hd);
        if (Rtl is { } rtl) Js.Set(o, "rtl", rtl);
        if (CharStyleId is { } csi) Js.Set(o, "charStyleId", csi);
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
        return o;
    }
}

public sealed record SpacingOptions
{
    public double? Before { get; init; }
    public double? After { get; init; }
    public double? LineHeight { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (Before is { } b) Js.Set(o, "before", b);
        if (After is { } a) Js.Set(o, "after", a);
        if (LineHeight is { } lh) Js.Set(o, "lineHeight", lh);
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

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (ColFractions is { } cf) Js.Set(o, "colFractions", Js.ToArray(e, cf.Select(x => (object?)x)));
        if (HeaderRow is { } hr) Js.Set(o, "headerRow", hr);
        if (Style is { } s) Js.Set(o, "style", s);
        if (StyleId is { } sid) Js.Set(o, "styleId", sid);
        if (CondOverrides is { } co) Js.Set(o, "condOverrides", co.ToJs(e));
        if (WidthMode is { } wm) Js.Set(o, "widthMode", EnumJs.WidthMode(wm));
        return o;
    }
}

public sealed record CellSpec
{
    public string? Text { get; init; }
    public int? ColSpan { get; init; }
    public int? RowSpan { get; init; }
    public string? Shading { get; init; }
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

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (Text is { } t) Js.Set(o, "text", t);
        if (ColSpan is { } cs) Js.Set(o, "colSpan", cs);
        if (RowSpan is { } rs) Js.Set(o, "rowSpan", rs);
        if (Shading is { } sh) Js.Set(o, "shading", sh);
        if (Style is { } st) Js.Set(o, "style", st.ToJs(e));
        if (Align is { } a) Js.Set(o, "align", EnumJs.Align(a));
        if (Borders is { } bd) Js.Set(o, "borders", bd.ToJs(e));
        if (Margin is { } mg) Js.Set(o, "margin", mg.ToJs(e));
        if (PreferredWidth is { } pw) Js.Set(o, "preferredWidth", pw.ToJs(e));
        if (VAlign is { } va) Js.Set(o, "vAlign", EnumJs.VAlign(va));
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
    public CharStyle? Style { get; init; }
    public TextAlign? Align { get; init; }
    public CellBorders? Borders { get; init; }
    public CellMargin? Margin { get; init; }
    public PreferredWidth? PreferredWidth { get; init; }
    /// <summary>Vertical content alignment (w:vAlign). Absent = top.</summary>
    public CellVAlign? VAlign { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (ColSpan is { } cs) Js.Set(o, "colSpan", cs);
        if (RowSpan is { } rs) Js.Set(o, "rowSpan", rs);
        if (Shading is { } sh) Js.Set(o, "shading", sh);
        if (Style is { } st) Js.Set(o, "style", st.ToJs(e));
        if (Align is { } a) Js.Set(o, "align", EnumJs.Align(a));
        if (Borders is { } bd) Js.Set(o, "borders", bd.ToJs(e));
        if (Margin is { } mg) Js.Set(o, "margin", mg.ToJs(e));
        if (PreferredWidth is { } pw) Js.Set(o, "preferredWidth", pw.ToJs(e));
        if (VAlign is { } va) Js.Set(o, "vAlign", EnumJs.VAlign(va));
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
    /// <summary>Base writing direction (OOXML w:bidi).</summary>
    public Direction? Direction { get; init; }
    /// <summary>Outline level 0..8 (TOC levels 1..9; docx w:outlineLvl) — heading styles carry this.</summary>
    public int? OutlineLevel { get; init; }
    public bool? PageBreakBefore { get; init; }
    public bool? ColumnBreakBefore { get; init; }
    /// <summary>Explicit tab stops (docx w:tabs).</summary>
    public IReadOnlyList<TabStop>? TabStops { get; init; }

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
        if (Direction is { } dir) Js.Set(o, "direction", EnumJs.Dir(dir));
        if (OutlineLevel is { } ol) Js.Set(o, "outlineLevel", Math.Clamp(ol, 0, 8));
        if (PageBreakBefore is { } pbb) Js.Set(o, "pageBreakBefore", pbb);
        if (ColumnBreakBefore is { } cbb) Js.Set(o, "columnBreakBefore", cbb);
        if (TabStops is { } ts) Js.Set(o, "tabStops", Js.ToArray(e, ts.Select(t => (object?)t.ToJs(e))));
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

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        Js.Set(o, "posPx", PosPx);
        if (Align is { } a) Js.Set(o, "align", EnumJs.TabAlign(a));
        if (Leader is { } l) Js.Set(o, "leader", EnumJs.TabLeader(l));
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
