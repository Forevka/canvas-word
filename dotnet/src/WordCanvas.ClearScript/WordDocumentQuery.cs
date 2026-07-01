using Microsoft.ClearScript;

namespace WordCanvas.ClearScript;

// Read-only query surface over an imported/built document — the rough analog of
// inspecting a .NET WordprocessingDocument. Each method calls a JS bridge function
// (see frontend/src/clearscript/queryBridge.ts) with the opaque in-V8 Doc handle
// and marshals the flat DTO array back. The DTO shapes MUST stay in lockstep with
// queryBridge.ts.

/// <summary>A paragraph flattened for the host: its text and where it lives.
/// <see cref="Row"/>/<see cref="Col"/> are -1 when the paragraph is not in a table
/// cell; the nullable fields are null when not applicable.</summary>
public sealed record ParagraphInfo(
    string Id,
    string Text,
    string Container,
    string? TableId,
    int Row,
    int Col,
    string? NoteKind,
    string? NoteId,
    string? StyleName,
    int? OutlineLevel);

/// <summary>One document section: page geometry plus the top-level block range it
/// covers. A single-section document reports one entry.</summary>
public sealed record SectionInfo(
    int Index,
    int StartBlock,
    int EndBlock,
    string BreakType,
    double PageWidthPx,
    double PageHeightPx,
    double MarginTop,
    double MarginRight,
    double MarginBottom,
    double MarginLeft,
    int ColumnCount);

/// <summary>One laid-out page: geometry plus the top-level body block ids placed on
/// it (a block split across a boundary appears on each page it covers).</summary>
public sealed record PageInfo(
    int Index,
    int Number,
    double WidthPx,
    double HeightPx,
    double MarginTop,
    double MarginRight,
    double MarginBottom,
    double MarginLeft,
    IReadOnlyList<string> BlockIds);

public sealed partial class WordDocument
{
    /// <summary>Every paragraph in the document (body, table cells, header/footer
    /// bands, and footnote/endnote bodies), flattened with its location.</summary>
    public IReadOnlyList<ParagraphInfo> GetParagraphs() =>
        ReadArray(_engine.Api.InvokeMethod("queryParagraphs", Doc), ReadParagraph);

    /// <summary>Paragraphs whose text contains <paramref name="needle"/> (substring).</summary>
    public IReadOnlyList<ParagraphInfo> FindText(string needle)
    {
        ArgumentNullException.ThrowIfNull(needle);
        return ReadArray(_engine.Api.InvokeMethod("findText", Doc, needle), ReadParagraph);
    }

    /// <summary>Enumerate the document's sections (page size, margins, columns) with
    /// the top-level block range each one covers.</summary>
    public IReadOnlyList<SectionInfo> GetSections() =>
        ReadArray(_engine.Api.InvokeMethod("querySections", Doc), ReadSection);

    /// <summary>Lay the document out and report which content lands on each page —
    /// the answer to "what's on page N". Page numbers can shift after edits, so
    /// re-query after mutating. Runs a full layout pass (like export).</summary>
    public IReadOnlyList<PageInfo> GetPages() =>
        ReadArray(_engine.ResolveValue(_engine.Api.InvokeMethod("layoutPages", Doc), "layoutPages"), ReadPage);

    // ---- marshalling --------------------------------------------------------

    private static ParagraphInfo ReadParagraph(ScriptObject p) => new(
        p.GetProperty("id")?.ToString() ?? "",
        p.GetProperty("text")?.ToString() ?? "",
        p.GetProperty("container")?.ToString() ?? "body",
        Str(p.GetProperty("tableId")),
        Convert.ToInt32(p.GetProperty("row")),
        Convert.ToInt32(p.GetProperty("col")),
        Str(p.GetProperty("noteKind")),
        Str(p.GetProperty("noteId")),
        Str(p.GetProperty("styleName")),
        Nullable(p.GetProperty("outlineLevel")));

    private static SectionInfo ReadSection(ScriptObject s) => new(
        Convert.ToInt32(s.GetProperty("index")),
        Convert.ToInt32(s.GetProperty("startBlock")),
        Convert.ToInt32(s.GetProperty("endBlock")),
        s.GetProperty("breakType")?.ToString() ?? "nextPage",
        Convert.ToDouble(s.GetProperty("pageWidthPx")),
        Convert.ToDouble(s.GetProperty("pageHeightPx")),
        Convert.ToDouble(s.GetProperty("marginTop")),
        Convert.ToDouble(s.GetProperty("marginRight")),
        Convert.ToDouble(s.GetProperty("marginBottom")),
        Convert.ToDouble(s.GetProperty("marginLeft")),
        Convert.ToInt32(s.GetProperty("columnCount")));

    private static PageInfo ReadPage(ScriptObject p) => new(
        Convert.ToInt32(p.GetProperty("index")),
        Convert.ToInt32(p.GetProperty("number")),
        Convert.ToDouble(p.GetProperty("widthPx")),
        Convert.ToDouble(p.GetProperty("heightPx")),
        Convert.ToDouble(p.GetProperty("marginTop")),
        Convert.ToDouble(p.GetProperty("marginRight")),
        Convert.ToDouble(p.GetProperty("marginBottom")),
        Convert.ToDouble(p.GetProperty("marginLeft")),
        ReadStrings(p.GetProperty("blockIds")));

    private static string? Str(object? v) => v is null or Undefined ? null : v.ToString();

    private static int? Nullable(object? v) => v is null or Undefined ? null : Convert.ToInt32(v);

    private static IReadOnlyList<T> ReadArray<T>(object? array, Func<ScriptObject, T> map)
    {
        if (array is not ScriptObject arr) return Array.Empty<T>();
        var len = Convert.ToInt32(arr.GetProperty("length"));
        var list = new List<T>(len);
        for (var i = 0; i < len; i++)
            if (arr.GetProperty(i) is ScriptObject el)
                list.Add(map(el));
        return list;
    }

    private static IReadOnlyList<string> ReadStrings(object? array)
    {
        if (array is not ScriptObject arr) return Array.Empty<string>();
        var len = Convert.ToInt32(arr.GetProperty("length"));
        var list = new List<string>(len);
        for (var i = 0; i < len; i++)
            list.Add(arr.GetProperty(i)?.ToString() ?? "");
        return list;
    }
}
