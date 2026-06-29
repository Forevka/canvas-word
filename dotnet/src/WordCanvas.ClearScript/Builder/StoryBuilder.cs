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

    /// <summary>Append an empty paragraph (a spacer / caret home).</summary>
    public TSelf Paragraph() => ParagraphCore(null, null);

    /// <summary>Append a paragraph with the given text.</summary>
    public TSelf Paragraph(string text) => ParagraphCore(text, null);

    /// <summary>Append a paragraph and configure its runs/formatting via the lambda.</summary>
    public TSelf Paragraph(Action<ParagraphBuilder> configure) => ParagraphCore(null, configure);

    /// <summary>Append a paragraph with the given text, then configure it via the lambda.</summary>
    public TSelf Paragraph(string text, Action<ParagraphBuilder> configure) => ParagraphCore(text, configure);

    private TSelf ParagraphCore(string? text, Action<ParagraphBuilder>? configure)
    {
        var pb = text is null
            ? (ScriptObject)JsScope.InvokeMethod("paragraph")
            : (ScriptObject)JsScope.InvokeMethod("paragraph", text);
        configure?.Invoke(new ParagraphBuilder(Engine, pb));
        return Self;
    }

    /// <summary>Append a display (block) equation from a LaTeX source (e.g. <c>\frac{a}{2}</c>).</summary>
    public TSelf Equation(string latex, EquationOptions? opts = null)
    {
        if (opts is not null) JsScope.InvokeMethod("equation", latex, opts.ToJs(Engine));
        else JsScope.InvokeMethod("equation", latex);
        return Self;
    }

    /// <summary>Append a display (block) equation from a presentation-MathML string.</summary>
    public TSelf EquationMathml(string mathml, EquationOptions? opts = null)
    {
        if (opts is not null) JsScope.InvokeMethod("equationMathml", mathml, opts.ToJs(Engine));
        else JsScope.InvokeMethod("equationMathml", mathml);
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

    /// <summary>Wrap every block the callback appends in a BLOCK-LEVEL content control
    /// (a rich-text <c>w:sdt</c> around whole paragraphs/tables — the multi-block analog of
    /// the inline controls). Nesting composes outer→inner.</summary>
    public TSelf ContentControlRange(Action<TSelf> build, SdtOptions? opts = null)
    {
        var props = Js.Obj(Engine);
        Js.Set(props, "type", "richText");
        if (opts?.Alias is { } a) Js.Set(props, "alias", a);
        if (opts?.Tag is { } t) Js.Set(props, "tag", t);
        Action<object> cb = _ => build(Self);
        JsScope.InvokeMethod("contentControlRange", props, cb);
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
    /// <summary>Underline with an explicit line style (e.g. "double", "dotted", "dash",
    /// "dotDash", "dotDotDash", "wave", "thick") and/or color (CSS hex). A null style
    /// paints a plain single line; a null color follows the run's text color.</summary>
    public ParagraphBuilder Underline(bool on, string? style, string? color = null)
    {
        var opts = Js.Obj(_engine);
        if (style is { } st) Js.Set(opts, "style", st);
        if (color is { } c) Js.Set(opts, "color", c);
        _js.InvokeMethod("underline", on, opts);
        return this;
    }
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
    /// <summary>All-caps display (OOXML w:caps): every letter renders uppercased; model text is unchanged.</summary>
    public ParagraphBuilder Caps(bool on = true) { _js.InvokeMethod("caps", on); return this; }
    /// <summary>Small-capitals display (OOXML w:smallCaps): uppercased letters with the originally-lowercase ones drawn smaller. Takes precedence over Caps.</summary>
    public ParagraphBuilder SmallCaps(bool on = true) { _js.InvokeMethod("smallCaps", on); return this; }
    /// <summary>Force this run's text to a right-to-left embedding (OOXML w:rtl).</summary>
    public ParagraphBuilder Rtl(bool on = true) { _js.InvokeMethod("rtl", on); return this; }

    /// <summary>Append a symbol-font glyph (OOXML w:sym). <paramref name="font"/> is the
    /// symbol font (e.g. "Wingdings"); <paramref name="charHex"/> is the hex code point Word
    /// stores (e.g. "F0E0"). The run renders the decoded glyph and re-emits w:sym on export.</summary>
    public ParagraphBuilder Symbol(string font, string charHex) { _js.InvokeMethod("symbol", font, charHex); return this; }

    /// <summary>Apply a registered character style (a character NamedStyle): bakes its
    /// formatting onto the runs AND sets the w:rStyle reference.</summary>
    public ParagraphBuilder CharStyle(string id) { _js.InvokeMethod("charStyle", id); return this; }

    /// <summary>Append an inline equation from a LaTeX source (a single replaced glyph).</summary>
    public ParagraphBuilder InlineEquation(string latex, CharStyle? style = null)
    {
        if (style is not null) _js.InvokeMethod("inlineEquation", latex, style.ToJs(_engine));
        else _js.InvokeMethod("inlineEquation", latex);
        return this;
    }

    /// <summary>Append an inline equation from a presentation-MathML string.</summary>
    public ParagraphBuilder InlineEquationMathml(string mathml, CharStyle? style = null)
    {
        if (style is not null) _js.InvokeMethod("inlineEquationMathml", mathml, style.ToJs(_engine));
        else _js.InvokeMethod("inlineEquationMathml", mathml);
        return this;
    }

    public ParagraphBuilder Align(TextAlign align) { _js.InvokeMethod("align", EnumJs.Align(align)); return this; }
    public ParagraphBuilder Spacing(SpacingOptions opts) { _js.InvokeMethod("spacing", opts.ToJs(_engine)); return this; }
    public ParagraphBuilder Indent(IndentOptions opts) { _js.InvokeMethod("indent", opts.ToJs(_engine)); return this; }
    public ParagraphBuilder KeepWithNext(bool on = true) { _js.InvokeMethod("keepWithNext", on); return this; }
    /// <summary>Never split this paragraph across pages/columns (docx w:keepLines).</summary>
    public ParagraphBuilder KeepTogether(bool on = true) { _js.InvokeMethod("keepTogether", on); return this; }
    /// <summary>Suppress before/after spacing between this paragraph and an adjacent same-style
    /// paragraph (docx w:contextualSpacing) — Word's list-style default.</summary>
    public ParagraphBuilder ContextualSpacing(bool on = true) { _js.InvokeMethod("contextualSpacing", on); return this; }
    /// <summary>Base writing direction (OOXML w:bidi); "rtl" lays the paragraph out right-to-left.</summary>
    public ParagraphBuilder Direction(Direction dir) { _js.InvokeMethod("direction", EnumJs.Dir(dir)); return this; }
    /// <summary>Outline level 0..8 (TOC levels 1..9) — a TOC entry without a heading style.</summary>
    public ParagraphBuilder OutlineLevel(int level) { _js.InvokeMethod("outlineLevel", level); return this; }

    /// <summary>Set explicit tab stops (docx w:tabs); a <c>\t</c> in text advances to the next.</summary>
    public ParagraphBuilder TabStops(IEnumerable<TabStop> stops)
    {
        _js.InvokeMethod("tabStops", Js.ToArray(_engine, stops.Select(s => (object?)s.ToJs(_engine))));
        return this;
    }

    /// <summary>Paragraph borders (OOXML w:pBdr) — a box drawn around the paragraph.
    /// Each edge reuses the table <see cref="CellBorder"/> value type.</summary>
    public ParagraphBuilder Borders(ParaBorders borders)
    {
        _js.InvokeMethod("borders", borders.ToJs(_engine));
        return this;
    }

    /// <summary>Paragraph shading — a CSS fill painted behind the paragraph (OOXML w:shd).</summary>
    public ParagraphBuilder Shading(string cssColor) { _js.InvokeMethod("shading", cssColor); return this; }

    /// <summary>Widow/orphan control (OOXML w:widowControl). Word's default is ON;
    /// pass false to let a lone first/last line break across a page boundary.</summary>
    public ParagraphBuilder WidowControl(bool on = true) { _js.InvokeMethod("widowControl", on); return this; }
    /// <summary>Exclude this paragraph from line numbering (OOXML w:suppressLineNumbers).</summary>
    public ParagraphBuilder SuppressLineNumbers(bool on = true) { _js.InvokeMethod("suppressLineNumbers", on); return this; }
    /// <summary>Vertical alignment of text within each line box (OOXML w:textAlignment).</summary>
    public ParagraphBuilder TextAlignment(LineVAlign v) { _js.InvokeMethod("textAlignment", EnumJs.TextVAlign(v)); return this; }
    /// <summary>Symmetric (mirrored) indents for facing-page layouts (OOXML w:mirrorIndents).</summary>
    public ParagraphBuilder MirrorIndents(bool on = true) { _js.InvokeMethod("mirrorIndents", on); return this; }
    /// <summary>Auto-adjust the right indent to the document grid (OOXML w:adjustRightInd).</summary>
    public ParagraphBuilder AdjustRightInd(bool on = true) { _js.InvokeMethod("adjustRightInd", on); return this; }

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

    /// <summary>A conditional (IF) field — branch text by comparing two operands.
    /// <paramref name="op"/> is one of <c>= &lt;&gt; &lt; &gt; &lt;= &gt;=</c>.</summary>
    public ParagraphBuilder IfField(string operandA, string op, string operandB, string ifTrue, string ifFalse)
    {
        _js.InvokeMethod("ifField", operandA, op, operandB, ifTrue, ifFalse);
        return this;
    }

    /// <summary>A cross-reference to a bookmark as a REF/PAGEREF field.</summary>
    /// <param name="pageRef">true → PAGEREF (the target's page number); false → REF (its text).</param>
    public ParagraphBuilder CrossReference(string bookmarkName, bool pageRef = false)
    {
        var opts = Js.Obj(_engine);
        Js.Set(opts, "kind", pageRef ? "pageRef" : "ref");
        _js.InvokeMethod("crossReference", bookmarkName, opts);
        return this;
    }

    // ---- inline content controls (SDT) ----
    public ParagraphBuilder RichTextControl(string text, SdtOptions? opts = null) => Sdt("richTextControl", text, opts);
    public ParagraphBuilder PlainTextControl(string text, SdtOptions? opts = null) => Sdt("plainTextControl", text, opts);

    /// <summary>A checkbox content control (☒ / ☐).</summary>
    public ParagraphBuilder Checkbox(bool isChecked, SdtOptions? opts = null)
    {
        if (opts is not null) _js.InvokeMethod("checkbox", isChecked, opts.ToJs(_engine));
        else _js.InvokeMethod("checkbox", isChecked);
        return this;
    }

    /// <summary>A drop-down content control (fixed list).</summary>
    public ParagraphBuilder DropDown(string selected, IEnumerable<SdtListItem> items, SdtOptions? opts = null)
        => SdtList("dropDown", selected, items, opts);

    /// <summary>A combo-box content control (list + free text).</summary>
    public ParagraphBuilder ComboBox(string selected, IEnumerable<SdtListItem> items, SdtOptions? opts = null)
        => SdtList("comboBox", selected, items, opts);

    /// <summary>A date-picker content control.</summary>
    public ParagraphBuilder DateControl(string text, string dateFormat = "M/d/yyyy", SdtOptions? opts = null)
    {
        if (opts is not null) _js.InvokeMethod("dateControl", text, dateFormat, opts.ToJs(_engine));
        else _js.InvokeMethod("dateControl", text, dateFormat);
        return this;
    }

    private ParagraphBuilder Sdt(string method, string text, SdtOptions? opts)
    {
        if (opts is not null) _js.InvokeMethod(method, text, opts.ToJs(_engine));
        else _js.InvokeMethod(method, text);
        return this;
    }

    private ParagraphBuilder SdtList(string method, string selected, IEnumerable<SdtListItem> items, SdtOptions? opts)
    {
        var arr = Js.ToArray(_engine, items.Select(i => (object?)i.ToJs(_engine)));
        if (opts is not null) _js.InvokeMethod(method, selected, arr, opts.ToJs(_engine));
        else _js.InvokeMethod(method, selected, arr);
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

    /// <summary>Append an auto-numbered endnote whose body is a single paragraph (collected at the document end).</summary>
    public ParagraphBuilder Endnote(string text) { _js.InvokeMethod("endnote", text); return this; }

    /// <summary>Append an auto-numbered endnote whose body is built via the callback.</summary>
    public ParagraphBuilder Endnote(Action<StoryBuilder> build)
    {
        Action<object> cb = s => build(new StoryBuilder(_engine, (ScriptObject)s));
        _js.InvokeMethod("endnote", cb);
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

    public TableBuilder Row(IEnumerable<CellContent> cells, RowOptions? opts = null)
    {
        var arr = Js.Arr(_engine);
        foreach (var c in cells) Js.Push(arr, c.ToJs(_engine));
        if (opts is not null) _js.InvokeMethod("row", arr, opts.ToJs(_engine));
        else _js.InvokeMethod("row", arr);
        return this;
    }

    public TableBuilder Row(Action<RowBuilder> build, RowOptions? opts = null)
    {
        Action<object> cb = r => build(new RowBuilder(_engine, (ScriptObject)r));
        if (opts is not null) _js.InvokeMethod("row", cb, opts.ToJs(_engine));
        else _js.InvokeMethod("row", cb);
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
