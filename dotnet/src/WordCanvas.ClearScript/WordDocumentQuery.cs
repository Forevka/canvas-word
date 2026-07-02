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

/// <summary>A content control (OOXML w:sdt) flattened with its place in the nesting
/// tree and the text it encloses — the primary templating surface. Controls nest
/// (membership is an ordered ancestry path), so <see cref="ParentId"/>/
/// <see cref="ChildIds"/>/<see cref="Path"/>/<see cref="Depth"/> describe the forest.
/// <see cref="Checked"/> is null for non-checkbox controls.</summary>
public sealed record SdtInfo(
    string Id,
    string SdtType,
    string? Tag,
    string? Alias,
    bool? Checked,
    bool Placeholder,
    string? ParentId,
    IReadOnlyList<string> ChildIds,
    IReadOnlyList<string> Path,
    int Depth,
    string Text);

/// <summary>A tracked field (custom + built-in): its verbatim instruction and kind.</summary>
public sealed record FieldInfo(string Id, string Name, string Kind, string Instruction);

/// <summary>A bookmark: name + its character range (block id + UTF-16 offset).</summary>
public sealed record BookmarkInfo(string Name, string StartBlockId, int StartOffset, string EndBlockId, int EndOffset);

/// <summary>A footnote/endnote story: its ref id and concatenated text.</summary>
public sealed record NoteInfo(string Id, string Text);

/// <summary>A paragraph bound to a list definition, with its resolved marker.</summary>
public sealed record ListItemInfo(string ParagraphId, int Level, string Marker, string Text, string Container);

/// <summary>A named style from the stylesheet.</summary>
public sealed record StyleInfo(string Id, string Name, string StyleType, string? BasedOn);

/// <summary>Where a block sits: its container plus (when in a table cell) the table
/// id and row/col (-1 when not in a cell), and note membership when in a note body.</summary>
public sealed record BlockPathInfo(string Container, string? TableId, int Row, int Col, string? NoteKind, string? NoteId);

/// <summary>A document position: block id + UTF-16 offset.</summary>
public sealed record PositionInfo(string BlockId, int Offset);

/// <summary>A content control's typed value. <see cref="Checked"/> is set for
/// checkboxes; <see cref="Selected"/> for dropDown/comboBox.</summary>
public sealed record SdtValueInfo(string SdtType, string Text, bool? Checked, string? Selected);

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

    /// <summary>Every content control (w:sdt) — the primary templating surface —
    /// flattened with its nesting links (parent/children/path/depth) and the text it
    /// encloses. Controls nest, so this is the whole forest in one call.</summary>
    public IReadOnlyList<SdtInfo> GetSdts() =>
        ReadArray(_engine.Api.InvokeMethod("querySdts", Doc), ReadSdt);

    // Each convenience filter re-queries by default; pass a list from a prior
    // GetSdts() to reuse it (WordDocument wraps an immutable snapshot) and avoid a
    // fresh querySdts bridge traversal.

    /// <summary>A single content control by id, or null.</summary>
    public SdtInfo? GetSdt(string id, IReadOnlyList<SdtInfo>? sdts = null)
    {
        ArgumentNullException.ThrowIfNull(id);
        return (sdts ?? GetSdts()).FirstOrDefault(s => s.Id == id);
    }

    /// <summary>The top-level controls — those not nested inside any other.</summary>
    public IReadOnlyList<SdtInfo> GetSdtRoots(IReadOnlyList<SdtInfo>? sdts = null) =>
        (sdts ?? GetSdts()).Where(s => s.ParentId is null).ToList();

    /// <summary>The controls nested directly (one level) inside <paramref name="id"/>.</summary>
    public IReadOnlyList<SdtInfo> GetSdtChildren(string id, IReadOnlyList<SdtInfo>? sdts = null)
    {
        ArgumentNullException.ThrowIfNull(id);
        return (sdts ?? GetSdts()).Where(s => s.ParentId == id).ToList();
    }

    /// <summary>Controls whose machine-readable tag (w:tag) equals <paramref name="tag"/>
    /// — the usual way to locate a template's content controls. Pure over the flattened list.</summary>
    public IReadOnlyList<SdtInfo> GetSdtsByTag(string tag, IReadOnlyList<SdtInfo>? sdts = null)
    {
        ArgumentNullException.ThrowIfNull(tag);
        return (sdts ?? GetSdts()).Where(s => s.Tag == tag).ToList();
    }

    /// <summary>Controls whose title (w:alias) equals <paramref name="alias"/>. Pure over
    /// the flattened list.</summary>
    public IReadOnlyList<SdtInfo> GetSdtsByAlias(string alias, IReadOnlyList<SdtInfo>? sdts = null)
    {
        ArgumentNullException.ThrowIfNull(alias);
        return (sdts ?? GetSdts()).Where(s => s.Alias == alias).ToList();
    }

    /// <summary>The controls wrapping <paramref name="id"/>, outermost→innermost
    /// (excluding <paramref name="id"/> itself). Pure over the flattened list.</summary>
    public IReadOnlyList<SdtInfo> GetSdtAncestors(string id, IReadOnlyList<SdtInfo>? sdts = null)
    {
        ArgumentNullException.ThrowIfNull(id);
        var all = sdts ?? GetSdts();
        var self = all.FirstOrDefault(s => s.Id == id);
        if (self is null) return Array.Empty<SdtInfo>();
        var byId = all.ToDictionary(s => s.Id);
        return self.Path
            .Take(self.Path.Count - 1) // drop self (the last path entry)
            .Where(byId.ContainsKey)
            .Select(pid => byId[pid])
            .ToList();
    }

    /// <summary>Every control nested anywhere below <paramref name="id"/> (depth-first,
    /// pre-order). Pure over the flattened list.</summary>
    public IReadOnlyList<SdtInfo> GetSdtDescendants(string id, IReadOnlyList<SdtInfo>? sdts = null)
    {
        ArgumentNullException.ThrowIfNull(id);
        var all = sdts ?? GetSdts();
        var byId = all.ToDictionary(s => s.Id);
        var result = new List<SdtInfo>();
        void Visit(string nid)
        {
            if (!byId.TryGetValue(nid, out var node)) return;
            foreach (var childId in node.ChildIds)
            {
                if (!byId.TryGetValue(childId, out var child)) continue;
                result.Add(child);
                Visit(childId);
            }
        }
        if (byId.ContainsKey(id)) Visit(id);
        return result;
    }

    /// <summary>A content control's typed value (text, plus checked/selected), or null.</summary>
    public SdtValueInfo? GetSdtValue(string id)
    {
        ArgumentNullException.ThrowIfNull(id);
        return ReadOne(_engine.Api.InvokeMethod("querySdtValue", Doc, id), ReadSdtValue);
    }

    // ---- fields / bookmarks / notes / lists / styles / location / text ------

    /// <summary>Every tracked field (custom + built-in), in document order.</summary>
    public IReadOnlyList<FieldInfo> GetFields() =>
        ReadArray(_engine.Api.InvokeMethod("queryFields", Doc), ReadField);

    /// <summary>A field by id, or null.</summary>
    public FieldInfo? GetField(string id, IReadOnlyList<FieldInfo>? fields = null)
    {
        ArgumentNullException.ThrowIfNull(id);
        return (fields ?? GetFields()).FirstOrDefault(f => f.Id == id);
    }

    /// <summary>Fields whose keyword matches <paramref name="name"/> (case-insensitive).</summary>
    public IReadOnlyList<FieldInfo> GetFieldsByName(string name, IReadOnlyList<FieldInfo>? fields = null)
    {
        ArgumentNullException.ThrowIfNull(name);
        return (fields ?? GetFields()).Where(f => string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase)).ToList();
    }

    /// <summary>Every bookmark (name → character range).</summary>
    public IReadOnlyList<BookmarkInfo> GetBookmarks() =>
        ReadArray(_engine.Api.InvokeMethod("queryBookmarks", Doc), ReadBookmark);

    /// <summary>A bookmark's range by name, or null.</summary>
    public BookmarkInfo? GetBookmark(string name, IReadOnlyList<BookmarkInfo>? bookmarks = null)
    {
        ArgumentNullException.ThrowIfNull(name);
        return (bookmarks ?? GetBookmarks()).FirstOrDefault(b => b.Name == name);
    }

    /// <summary>Footnote stories (ref id + text).</summary>
    public IReadOnlyList<NoteInfo> GetFootnotes() =>
        ReadArray(_engine.Api.InvokeMethod("queryFootnotes", Doc), ReadNote);

    /// <summary>Endnote stories (ref id + text).</summary>
    public IReadOnlyList<NoteInfo> GetEndnotes() =>
        ReadArray(_engine.Api.InvokeMethod("queryEndnotes", Doc), ReadNote);

    /// <summary>Paragraphs bound to a list definition, in body reading order, each
    /// with its resolved marker ("1.", "a.", "•", …).</summary>
    public IReadOnlyList<ListItemInfo> GetListItems(string listId)
    {
        ArgumentNullException.ThrowIfNull(listId);
        return ReadArray(_engine.Api.InvokeMethod("queryListItems", Doc, listId), ReadListItem);
    }

    /// <summary>Enumerate the stylesheet (Word's style gallery).</summary>
    public IReadOnlyList<StyleInfo> GetStyles() =>
        ReadArray(_engine.Api.InvokeMethod("queryStyles", Doc), ReadStyle);

    /// <summary>A named style by id, or null.</summary>
    public StyleInfo? GetStyleById(string id, IReadOnlyList<StyleInfo>? styles = null)
    {
        ArgumentNullException.ThrowIfNull(id);
        return (styles ?? GetStyles()).FirstOrDefault(s => s.Id == id);
    }

    /// <summary>Where a block sits (container / table cell / note), or null.</summary>
    public BlockPathInfo? GetBlockPath(string id)
    {
        ArgumentNullException.ThrowIfNull(id);
        return ReadOne(_engine.Api.InvokeMethod("queryBlockPath", Doc, id), ReadBlockPath);
    }

    /// <summary>The first position of <paramref name="needle"/> (a substring) in
    /// reading order, or null.</summary>
    public PositionInfo? PositionOfText(string needle)
    {
        ArgumentNullException.ThrowIfNull(needle);
        return ReadOne(_engine.Api.InvokeMethod("queryPositionOfText", Doc, needle), ReadPosition);
    }

    /// <summary>The text a selection covers (see the TS <c>rangeText</c>): a single-block
    /// range slices that block; a multi-block range spans top-level body blocks.</summary>
    public string RangeText(string startBlockId, int startOffset, string endBlockId, int endOffset)
    {
        ArgumentNullException.ThrowIfNull(startBlockId);
        ArgumentNullException.ThrowIfNull(endBlockId);
        return _engine.Api.InvokeMethod("queryRangeText", Doc, startBlockId, startOffset, endBlockId, endOffset)?.ToString() ?? "";
    }

    /// <summary>Where a top-level block sits in a laid-out page map: its page index and
    /// 0-based order among that page's blocks, or null. Pure over the page list.</summary>
    public (int PageIndex, int Order)? IndexOnPage(string blockId, IReadOnlyList<PageInfo>? pages = null)
    {
        ArgumentNullException.ThrowIfNull(blockId);
        foreach (var page in pages ?? GetPages())
        {
            for (var i = 0; i < page.BlockIds.Count; i++)
                if (page.BlockIds[i] == blockId) return (page.Index, i);
        }
        return null;
    }

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

    private static SdtInfo ReadSdt(ScriptObject s) => new(
        s.GetProperty("id")?.ToString() ?? "",
        s.GetProperty("sdtType")?.ToString() ?? "",
        Str(s.GetProperty("tag")),
        Str(s.GetProperty("alias")),
        Bool(s.GetProperty("checked")),
        Convert.ToBoolean(s.GetProperty("placeholder")),
        Str(s.GetProperty("parentId")),
        ReadStrings(s.GetProperty("childIds")),
        ReadStrings(s.GetProperty("path")),
        Convert.ToInt32(s.GetProperty("depth")),
        s.GetProperty("text")?.ToString() ?? "");

    private static FieldInfo ReadField(ScriptObject f) => new(
        f.GetProperty("id")?.ToString() ?? "",
        f.GetProperty("name")?.ToString() ?? "",
        f.GetProperty("kind")?.ToString() ?? "",
        f.GetProperty("instruction")?.ToString() ?? "");

    private static BookmarkInfo ReadBookmark(ScriptObject b) => new(
        b.GetProperty("name")?.ToString() ?? "",
        b.GetProperty("startBlockId")?.ToString() ?? "",
        Convert.ToInt32(b.GetProperty("startOffset")),
        b.GetProperty("endBlockId")?.ToString() ?? "",
        Convert.ToInt32(b.GetProperty("endOffset")));

    private static NoteInfo ReadNote(ScriptObject n) => new(
        n.GetProperty("id")?.ToString() ?? "",
        n.GetProperty("text")?.ToString() ?? "");

    private static ListItemInfo ReadListItem(ScriptObject l) => new(
        l.GetProperty("paragraphId")?.ToString() ?? "",
        Convert.ToInt32(l.GetProperty("level")),
        l.GetProperty("marker")?.ToString() ?? "",
        l.GetProperty("text")?.ToString() ?? "",
        l.GetProperty("container")?.ToString() ?? "body");

    private static StyleInfo ReadStyle(ScriptObject s) => new(
        s.GetProperty("id")?.ToString() ?? "",
        s.GetProperty("name")?.ToString() ?? "",
        s.GetProperty("styleType")?.ToString() ?? "paragraph",
        Str(s.GetProperty("basedOn")));

    private static BlockPathInfo ReadBlockPath(ScriptObject p) => new(
        p.GetProperty("container")?.ToString() ?? "body",
        Str(p.GetProperty("tableId")),
        Convert.ToInt32(p.GetProperty("row")),
        Convert.ToInt32(p.GetProperty("col")),
        Str(p.GetProperty("noteKind")),
        Str(p.GetProperty("noteId")));

    private static PositionInfo ReadPosition(ScriptObject p) => new(
        p.GetProperty("blockId")?.ToString() ?? "",
        Convert.ToInt32(p.GetProperty("offset")));

    private static SdtValueInfo ReadSdtValue(ScriptObject v) => new(
        v.GetProperty("sdtType")?.ToString() ?? "",
        v.GetProperty("text")?.ToString() ?? "",
        Bool(v.GetProperty("checked")),
        Str(v.GetProperty("selected")));

    private static string? Str(object? v) => v is null or Undefined ? null : v.ToString();

    private static int? Nullable(object? v) => v is null or Undefined ? null : Convert.ToInt32(v);

    private static bool? Bool(object? v) => v is null or Undefined ? null : Convert.ToBoolean(v);

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

    /// <summary>Marshal a single bridge result that may be an object or JS null/undefined.</summary>
    private static T? ReadOne<T>(object? value, Func<ScriptObject, T> map) where T : class =>
        value is ScriptObject obj ? map(obj) : null;

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
