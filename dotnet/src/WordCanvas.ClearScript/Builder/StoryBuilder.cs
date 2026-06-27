using Microsoft.ClearScript;

namespace WordCanvas.ClearScript.Builder;

/// <summary>
/// Block-scope surface shared by the document body, header/footer bands, and table
/// cells — a strongly-typed C# facade over the JS StoryBuilder it wraps. Methods
/// append eagerly and return the concrete self type for fluent chaining. Paragraph
/// run-styling is done through a configure lambda (the C#-idiomatic shape of the JS
/// ParagraphBuilder scope).
/// </summary>
/// <typeparam name="TSelf">The concrete builder (CRTP) so chains keep their type.</typeparam>
public abstract class StoryBuilderBase<TSelf> where TSelf : StoryBuilderBase<TSelf>
{
    internal readonly WordCanvasEngine Engine;
    internal readonly ScriptObject JsScope;

    private protected StoryBuilderBase(WordCanvasEngine engine, ScriptObject jsScope)
    {
        Engine = engine;
        JsScope = jsScope;
    }

    private protected abstract TSelf Self { get; }

    /// <summary>Append a paragraph; configure its runs/formatting via the lambda.</summary>
    public TSelf Paragraph(string? text = null, Action<ParagraphBuilder>? configure = null)
    {
        var pb = text is null
            ? (ScriptObject)JsScope.InvokeMethod("paragraph")
            : (ScriptObject)JsScope.InvokeMethod("paragraph", text);
        configure?.Invoke(new ParagraphBuilder(Engine, pb));
        return Self;
    }

    /// <summary>Data-driven table from a 2D grid of cell text/specs.</summary>
    public TSelf Table(IEnumerable<IEnumerable<CellContent>> rows, TableOptions? opts = null)
    {
        var jsRows = Js.Arr(Engine);
        foreach (var row in rows)
        {
            var jsRow = Js.Arr(Engine);
            foreach (var cell in row) Js.Push(jsRow, cell.ToJs(Engine));
            Js.Push(jsRows, jsRow);
        }
        if (opts is not null) JsScope.InvokeMethod("table", jsRows, opts.ToJs(Engine));
        else JsScope.InvokeMethod("table", jsRows);
        return Self;
    }

    /// <summary>Structural table: a callback scope for spans, nested blocks, full control.</summary>
    public TSelf Table(Action<TableBuilder> build, TableOptions? opts = null)
    {
        Action<object> cb = t => build(new TableBuilder(Engine, (ScriptObject)t));
        if (opts is not null) JsScope.InvokeMethod("table", cb, opts.ToJs(Engine));
        else JsScope.InvokeMethod("table", cb);
        return Self;
    }

    /// <summary>Insert an image from a URL (https:/data:).</summary>
    public TSelf Image(string src, ImageOptions opts)
    {
        JsScope.InvokeMethod("image", src, opts.ToJs(Engine));
        return Self;
    }

    /// <summary>Insert an image from raw bytes (inlined as a data: URL).</summary>
    public TSelf Image(byte[] data, string mime, ImageOptions opts)
    {
        var srcObj = Js.Obj(Engine);
        Js.Set(srcObj, "data", Engine.AllocBytes(data));
        Js.Set(srcObj, "mime", mime);
        JsScope.InvokeMethod("image", srcObj, opts.ToJs(Engine));
        return Self;
    }

    /// <summary>One paragraph per item, marked as list members.</summary>
    public TSelf List(IEnumerable<ListItem> items, ListOptions? opts = null)
    {
        var jsItems = Js.ToArray(Engine, items.Select(i => (object?)i.ToJs(Engine)));
        if (opts is not null) JsScope.InvokeMethod("list", jsItems, opts.ToJs(Engine));
        else JsScope.InvokeMethod("list", jsItems);
        return Self;
    }

    public TSelf BulletList(params string[] items)
    {
        JsScope.InvokeMethod("bulletList", Js.ToArray(Engine, items.Select(s => (object?)s)));
        return Self;
    }

    public TSelf NumberedList(params string[] items)
    {
        JsScope.InvokeMethod("numberedList", Js.ToArray(Engine, items.Select(s => (object?)s)));
        return Self;
    }

    /// <summary>The next block starts a new page.</summary>
    public TSelf PageBreak() { JsScope.InvokeMethod("pageBreak"); return Self; }

    /// <summary>The next block starts a new newspaper column.</summary>
    public TSelf ColumnBreak() { JsScope.InvokeMethod("columnBreak"); return Self; }

    /// <summary>Bookmark everything the callback appends to this story.</summary>
    public TSelf BookmarkRange(string name, Action<TSelf> build)
    {
        Action<object> cb = _ => build(Self);
        JsScope.InvokeMethod("bookmarkRange", name, cb);
        return Self;
    }
}

/// <summary>Concrete block scope for header/footer bands and table cells.</summary>
public sealed class StoryBuilder : StoryBuilderBase<StoryBuilder>
{
    internal StoryBuilder(WordCanvasEngine engine, ScriptObject jsScope) : base(engine, jsScope) { }
    private protected override StoryBuilder Self => this;
}

/// <summary>
/// Paragraph scope. Character formatting patches every run in the paragraph and
/// becomes the default for later Text() runs; mixed formatting uses Text(t, style).
/// </summary>
public sealed class ParagraphBuilder
{
    private readonly WordCanvasEngine _engine;
    private readonly ScriptObject _js;

    internal ParagraphBuilder(WordCanvasEngine engine, ScriptObject js)
    {
        _engine = engine;
        _js = js;
    }

    /// <summary>Apply a named style (direct formatting applied after wins).</summary>
    public ParagraphBuilder WithStyle(string id) { _js.InvokeMethod("withStyle", id); return this; }

    /// <summary>Append a run; inherits the scope's accrued formatting, <paramref name="style"/> wins.</summary>
    public ParagraphBuilder Text(string text, CharStyle? style = null)
    {
        if (style is not null) _js.InvokeMethod("text", text, style.ToJs(_engine));
        else _js.InvokeMethod("text", text);
        return this;
    }

    public ParagraphBuilder Bold(bool on = true) { _js.InvokeMethod("bold", on); return this; }
    public ParagraphBuilder Italic(bool on = true) { _js.InvokeMethod("italic", on); return this; }
    public ParagraphBuilder Underline(bool on = true) { _js.InvokeMethod("underline", on); return this; }
    public ParagraphBuilder Strikethrough(bool on = true) { _js.InvokeMethod("strikethrough", on); return this; }
    public ParagraphBuilder Color(string cssColor) { _js.InvokeMethod("color", cssColor); return this; }
    public ParagraphBuilder Highlight(string cssColor) { _js.InvokeMethod("highlight", cssColor); return this; }
    public ParagraphBuilder FontSize(double px) { _js.InvokeMethod("fontSize", px); return this; }
    public ParagraphBuilder Font(string family) { _js.InvokeMethod("font", family); return this; }
    public ParagraphBuilder Link(string url) { _js.InvokeMethod("link", url); return this; }
    public ParagraphBuilder Superscript(bool on = true) { _js.InvokeMethod("superscript", on); return this; }
    public ParagraphBuilder Subscript(bool on = true) { _js.InvokeMethod("subscript", on); return this; }
    public ParagraphBuilder LetterSpacing(double px) { _js.InvokeMethod("letterSpacing", px); return this; }
    public ParagraphBuilder Hidden(bool on = true) { _js.InvokeMethod("hidden", on); return this; }

    public ParagraphBuilder Align(TextAlign align) { _js.InvokeMethod("align", EnumJs.Align(align)); return this; }
    public ParagraphBuilder Spacing(SpacingOptions opts) { _js.InvokeMethod("spacing", opts.ToJs(_engine)); return this; }
    public ParagraphBuilder Indent(IndentOptions opts) { _js.InvokeMethod("indent", opts.ToJs(_engine)); return this; }
    public ParagraphBuilder KeepWithNext(bool on = true) { _js.InvokeMethod("keepWithNext", on); return this; }

    // ---- inline fields ----
    public ParagraphBuilder PageField() { _js.InvokeMethod("pageField"); return this; }
    public ParagraphBuilder NumPagesField() { _js.InvokeMethod("numPagesField"); return this; }
    public ParagraphBuilder DateField(string format = "M/d/yyyy") { _js.InvokeMethod("dateField", format); return this; }
    public ParagraphBuilder TimeField(string format = "h:mm AM/PM") { _js.InvokeMethod("timeField", format); return this; }

    /// <summary>A custom (host-resolved) field: verbatim instruction + cached result.</summary>
    public ParagraphBuilder CustomField(string instruction, string resultText)
    {
        _js.InvokeMethod("customField", instruction, resultText);
        return this;
    }

    // ---- bookmarks + footnotes ----
    /// <summary>Append text and bookmark exactly that run's range.</summary>
    public ParagraphBuilder Bookmark(string name, string text, CharStyle? style = null)
    {
        if (style is not null) _js.InvokeMethod("bookmark", name, text, style.ToJs(_engine));
        else _js.InvokeMethod("bookmark", name, text);
        return this;
    }

    /// <summary>Append an auto-numbered footnote whose body is a single paragraph.</summary>
    public ParagraphBuilder Footnote(string text) { _js.InvokeMethod("footnote", text); return this; }

    /// <summary>Append an auto-numbered footnote whose body is built via the callback.</summary>
    public ParagraphBuilder Footnote(Action<StoryBuilder> build)
    {
        Action<object> cb = s => build(new StoryBuilder(_engine, (ScriptObject)s));
        _js.InvokeMethod("footnote", cb);
        return this;
    }
}

/// <summary>Structural table scope (rows, spans, column fractions).</summary>
public sealed class TableBuilder
{
    private readonly WordCanvasEngine _engine;
    private readonly ScriptObject _js;

    internal TableBuilder(WordCanvasEngine engine, ScriptObject js)
    {
        _engine = engine;
        _js = js;
    }

    public TableBuilder Row(IEnumerable<CellContent> cells)
    {
        var arr = Js.Arr(_engine);
        foreach (var c in cells) Js.Push(arr, c.ToJs(_engine));
        _js.InvokeMethod("row", arr);
        return this;
    }

    public TableBuilder Row(Action<RowBuilder> build)
    {
        Action<object> cb = r => build(new RowBuilder(_engine, (ScriptObject)r));
        _js.InvokeMethod("row", cb);
        return this;
    }

    public TableBuilder Rows(IEnumerable<IEnumerable<CellContent>> data)
    {
        var jsRows = Js.Arr(_engine);
        foreach (var row in data)
        {
            var jr = Js.Arr(_engine);
            foreach (var c in row) Js.Push(jr, c.ToJs(_engine));
            Js.Push(jsRows, jr);
        }
        _js.InvokeMethod("rows", jsRows);
        return this;
    }

    public TableBuilder ColFractions(params double[] fractions)
    {
        _js.InvokeMethod("colFractions", Js.ToArray(_engine, fractions.Select(x => (object?)x)));
        return this;
    }
}

/// <summary>Row scope inside a structural table.</summary>
public sealed class RowBuilder
{
    private readonly WordCanvasEngine _engine;
    private readonly ScriptObject _js;

    internal RowBuilder(WordCanvasEngine engine, ScriptObject js)
    {
        _engine = engine;
        _js = js;
    }

    /// <summary>Append a cell from text/spec, with optional cell formatting.</summary>
    public RowBuilder Cell(CellContent content, CellOptions? opts = null)
    {
        if (opts is not null) _js.InvokeMethod("cell", content.ToJs(_engine), opts.ToJs(_engine));
        else _js.InvokeMethod("cell", content.ToJs(_engine));
        return this;
    }

    /// <summary>Append a cell whose blocks are built via the callback scope.</summary>
    public RowBuilder Cell(Action<StoryBuilder> build, CellOptions? opts = null)
    {
        Action<object> cb = s => build(new StoryBuilder(_engine, (ScriptObject)s));
        if (opts is not null) _js.InvokeMethod("cell", cb, opts.ToJs(_engine));
        else _js.InvokeMethod("cell", cb);
        return this;
    }
}
