using Microsoft.ClearScript;

namespace WordCanvas.ClearScript.Builder;

/// <summary>A `\nextPage` section break's per-section overrides (geometry + its own
/// header/footer bands).</summary>
public sealed record SectionBreakOptions
{
    public PageSizeName? PageSize { get; init; }
    public PageSize? CustomPageSize { get; init; }
    public Orientation? Orientation { get; init; }
    public Margins? Margins { get; init; }
    /// <summary>Newspaper columns for the new section; set <see cref="ClearColumns"/>
    /// to drop back to a single column.</summary>
    public ColumnsSpec? Columns { get; init; }
    public bool ClearColumns { get; init; }
    public int? PageNumberStart { get; init; }
    /// <summary>How the new section begins (next page / even / odd page parity).</summary>
    public SectionBreakType? BreakType { get; init; }
    /// <summary>Line numbering in the margin (w:lnNumType) for the new section.</summary>
    public LineNumbering? LineNumbering { get; init; }
    public Action<StoryBuilder>? Header { get; init; }
    public Action<StoryBuilder>? Footer { get; init; }
    public Action<StoryBuilder>? HeaderFirst { get; init; }
    public Action<StoryBuilder>? FooterFirst { get; init; }
    public Action<StoryBuilder>? HeaderEven { get; init; }
    public Action<StoryBuilder>? FooterEven { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine e)
    {
        var o = Js.Obj(e);
        SpecJs.SetPageSize(e, o, PageSize, CustomPageSize);
        if (Orientation is { } orient) Js.Set(o, "orientation", EnumJs.Orient(orient));
        if (Margins is { } m) Js.Set(o, "margins", m.ToJs(e));
        if (ClearColumns) o.SetProperty("columns", null!);
        else if (Columns is { } col) Js.Set(o, "columns", SpecJs.Columns(e, col));
        if (PageNumberStart is { } pns) Js.Set(o, "pageNumberStart", pns);
        if (BreakType is { } bt) Js.Set(o, "breakType", EnumJs.BreakType(bt));
        if (LineNumbering is { } ln) Js.Set(o, "lineNumbering", ln.ToJs(e));
        Band(e, o, "header", Header);
        Band(e, o, "footer", Footer);
        Band(e, o, "headerFirst", HeaderFirst);
        Band(e, o, "footerFirst", FooterFirst);
        Band(e, o, "headerEven", HeaderEven);
        Band(e, o, "footerEven", FooterEven);
        return o;
    }

    private static void Band(WordCanvasEngine e, ScriptObject target, string key, Action<StoryBuilder>? build)
    {
        if (build is null) return;
        Action<object> cb = s => build(new StoryBuilder(e, (ScriptObject)s));
        Js.Set(target, key, cb);
    }
}

/// <summary>
/// Root of the fluent builder — a strongly-typed C# facade over the JS
/// DocumentBuilder. Produces the same plain-data document the editor renders and the
/// exporters write; <see cref="Build"/> returns an in-V8 <see cref="WordDocument"/>
/// handle ready to export.
/// </summary>
public sealed class DocumentBuilder : StoryBuilderBase<DocumentBuilder>
{
    private DocumentBuilder(WordCanvasEngine engine, ScriptObject jsScope) : base(engine, jsScope) { }
    private protected override DocumentBuilder Self => this;

    internal static DocumentBuilder Create(WordCanvasEngine engine, CreateOptions? options)
    {
        var ctor = (ScriptObject)engine.Api.GetProperty("DocumentBuilder");
        var js = options is null
            ? (ScriptObject)ctor.InvokeMethod("create")
            : (ScriptObject)ctor.InvokeMethod("create", options.ToJs(engine));
        return new DocumentBuilder(engine, js);
    }

    internal static DocumentBuilder FromTemplate(WordCanvasEngine engine, byte[] templateDocx, TemplateOptions? options)
    {
        ArgumentNullException.ThrowIfNull(templateDocx);
        var ctor = (ScriptObject)engine.Api.GetProperty("DocumentBuilder");
        var bytes = engine.AllocBytes(templateDocx);
        var promise = options is null
            ? ctor.InvokeMethod("fromTemplate", bytes)
            : ctor.InvokeMethod("fromTemplate", bytes, options.ToJs(engine));
        var js = engine.ResolveValue(promise, "DocumentBuilder.fromTemplate") as ScriptObject
                 ?? throw new WordCanvasException("fromTemplate returned no builder");
        return new DocumentBuilder(engine, js);
    }

    /// <summary>Compose a header band (replaces the variant's existing story).</summary>
    public DocumentBuilder Header(Action<StoryBuilder> build, BandVariant variant = BandVariant.Default)
    {
        Action<object> cb = s => build(new StoryBuilder(Engine, (ScriptObject)s));
        JsScope.InvokeMethod("header", cb, BandOpts(variant));
        return this;
    }

    /// <summary>Compose a footer band (replaces the variant's existing story).</summary>
    public DocumentBuilder Footer(Action<StoryBuilder> build, BandVariant variant = BandVariant.Default)
    {
        Action<object> cb = s => build(new StoryBuilder(Engine, (ScriptObject)s));
        JsScope.InvokeMethod("footer", cb, BandOpts(variant));
        return this;
    }

    /// <summary>Merge page geometry onto the section (template values are the base).</summary>
    public DocumentBuilder PageSetup(PageSetup setup)
    {
        JsScope.InvokeMethod("pageSetup", setup.ToJs(Engine));
        return this;
    }

    /// <summary>Register (or override) a named style — call BEFORE applying it.</summary>
    public DocumentBuilder Style(NamedStyle def)
    {
        JsScope.InvokeMethod("style", def.ToJs(Engine));
        return this;
    }

    /// <summary>Make <paramref name="id"/> the document default style (call early).</summary>
    public DocumentBuilder DefaultStyle(string id)
    {
        JsScope.InvokeMethod("defaultStyle", id);
        return this;
    }

    /// <summary>Set the document's default tab interval in px (OOXML settings.xml
    /// w:defaultTabStop). A <c>\t</c> past the last explicit tab stop advances to the next
    /// multiple of this. Non-positive values are ignored.</summary>
    public DocumentBuilder DefaultTabStop(double px)
    {
        JsScope.InvokeMethod("defaultTabStop", px);
        return this;
    }

    /// <summary>Register (or override) a REAL table style (OOXML w:style[type=table]) with
    /// conditional bands. Reference it from <c>.Table(rows, new TableOptions { StyleId = def.Id })</c>.</summary>
    public DocumentBuilder TableStyle(TableStyleDef def)
    {
        JsScope.InvokeMethod("tableStyle", def.ToJs(Engine));
        return this;
    }

    /// <summary>Register (or override) a builder-only table-style PRESET — reusable cell
    /// formatting applied via <c>.Table(rows, new TableOptions { Style = name })</c>.</summary>
    public DocumentBuilder TableStylePreset(string name, TableStylePreset preset)
    {
        JsScope.InvokeMethod("tableStylePreset", name, preset.ToJs(Engine));
        return this;
    }

    /// <summary>Register a custom list definition; reference it via
    /// <c>.List(items, new ListOptions { ListId = id })</c>.</summary>
    public DocumentBuilder ListDefinition(string id, ListDefinitionSpec spec)
    {
        JsScope.InvokeMethod("listDefinition", id, spec.ToJs(Engine));
        return this;
    }

    /// <summary>Insert a table of contents built from the document's headings at Build().</summary>
    public DocumentBuilder TableOfContents(TocOptions? opts = null)
    {
        if (opts is not null) JsScope.InvokeMethod("tableOfContents", opts.ToJs(Engine));
        else JsScope.InvokeMethod("tableOfContents");
        return this;
    }

    /// <summary>End the current section and start a new one on the next page.</summary>
    public DocumentBuilder SectionBreak(SectionBreakOptions? opts = null)
    {
        if (opts is not null) JsScope.InvokeMethod("sectionBreak", opts.ToJs(Engine));
        else JsScope.InvokeMethod("sectionBreak");
        return this;
    }

    /// <summary>Lossy decisions made while building (unknown style ids, template warnings…).</summary>
    public IReadOnlyList<WordWarning> Warnings
    {
        get
        {
            if (JsScope.GetProperty("warnings") is not ScriptObject arr) return Array.Empty<WordWarning>();
            var len = Convert.ToInt32(arr.GetProperty("length"));
            var list = new List<WordWarning>(len);
            for (var i = 0; i < len; i++)
            {
                if (arr.GetProperty(i) is not ScriptObject w) continue;
                var code = w.GetProperty("code")?.ToString() ?? "";
                var message = w.GetProperty("message");
                list.Add(new WordWarning(code, message is Undefined ? null : message?.ToString()));
            }
            return list;
        }
    }

    /// <summary>Snapshot the document into an in-V8 handle ready to export. The builder
    /// stays usable afterwards.</summary>
    public WordDocument Build()
    {
        var doc = JsScope.InvokeMethod("build");
        return WordDocument.FromBuilt(Engine, doc);
    }

    private ScriptObject BandOpts(BandVariant variant)
    {
        var o = Js.Obj(Engine);
        Js.Set(o, "variant", EnumJs.Variant(variant));
        return o;
    }
}
