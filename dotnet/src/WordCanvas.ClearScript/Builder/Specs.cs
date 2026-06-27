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

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
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

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (ColFractions is { } cf) Js.Set(o, "colFractions", Js.ToArray(e, cf.Select(x => (object?)x)));
        if (HeaderRow is { } hr) Js.Set(o, "headerRow", hr);
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

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (Text is { } t) Js.Set(o, "text", t);
        if (ColSpan is { } cs) Js.Set(o, "colSpan", cs);
        if (RowSpan is { } rs) Js.Set(o, "rowSpan", rs);
        if (Shading is { } sh) Js.Set(o, "shading", sh);
        if (Style is { } st) Js.Set(o, "style", st.ToJs(e));
        if (Align is { } a) Js.Set(o, "align", EnumJs.Align(a));
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

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        if (ColSpan is { } cs) Js.Set(o, "colSpan", cs);
        if (RowSpan is { } rs) Js.Set(o, "rowSpan", rs);
        if (Shading is { } sh) Js.Set(o, "shading", sh);
        if (Style is { } st) Js.Set(o, "style", st.ToJs(e));
        if (Align is { } a) Js.Set(o, "align", EnumJs.Align(a));
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

/// <summary>Paragraph formatting patch (used by NamedStyle.Para).</summary>
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
