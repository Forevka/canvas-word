using Microsoft.ClearScript;

namespace WordCanvas.ClearScript;

/// <summary>
/// An opaque handle to a document living inside the V8 engine (imported from a
/// .docx or produced by <see cref="Builder.DocumentBuilder"/>). The model data
/// never leaves V8; only export blobs cross the boundary.
/// </summary>
public sealed partial class WordDocument
{
    private readonly WordCanvasEngine _engine;

    /// <summary>The in-V8 <c>Document</c> object.</summary>
    internal object Doc { get; }

    /// <summary>The in-V8 <c>cw-media:N -&gt; Uint8Array</c> map (import only; null
    /// for builder output), passed back to export so embedded images round-trip.</summary>
    internal object? Images { get; }

    public int BlockCount { get; }
    public int MediaCount { get; }
    public IReadOnlyList<WordWarning> Warnings { get; }

    private WordDocument(WordCanvasEngine engine, object doc, object? images, int blockCount, int mediaCount, IReadOnlyList<WordWarning> warnings)
    {
        _engine = engine;
        Doc = doc;
        Images = images;
        BlockCount = blockCount;
        MediaCount = mediaCount;
        Warnings = warnings;
    }

    internal static WordDocument FromImport(WordCanvasEngine engine, ScriptObject handle)
    {
        var doc = handle.GetProperty("doc");
        var images = handle.GetProperty("images");
        var blocks = Convert.ToInt32(handle.GetProperty("blockCount"));
        var media = Convert.ToInt32(handle.GetProperty("mediaCount"));
        var warnings = ReadWarnings(handle.GetProperty("warnings") as ScriptObject);
        return new WordDocument(engine, doc, images, blocks, media, warnings);
    }

    internal static WordDocument FromBuilt(WordCanvasEngine engine, object doc) =>
        new(engine, doc, null, engine.CountBlocks(doc), 0, Array.Empty<WordWarning>());

    /// <summary>A new handle over a transformed model doc, preserving this document's
    /// image bytes + warnings (used by UpdateFields).</summary>
    internal WordDocument WithDoc(object newDoc) =>
        new(_engine, newDoc, Images, _engine.CountBlocks(newDoc), MediaCount, Warnings);

    /// <summary>Update fields (TOC entries + cached page numbers) on this in-memory
    /// document and return the updated handle. See <see cref="WordCanvasEngine.UpdateFields"/>.</summary>
    public WordDocument UpdateFields(Builder.TocOptions? options = null) => _engine.UpdateFields(this, options);

    /// <summary>Open a stateful editor over this document for programmatic in-place
    /// edits (with undo/redo). Call <see cref="WordDocumentEditor.ToDocument"/> to get
    /// a handle over the edited model to query or export.</summary>
    public WordDocumentEditor Edit() =>
        new(_engine, this, (ScriptObject)_engine.Api.InvokeMethod("openEditor", Doc));

    /// <summary>Append another document's content after this one — the headless
    /// equivalent of Word's "insert file at end" (a drop-in for Syncfusion
    /// <c>ImportContent</c>). Reconciles every id space (styles, lists, table styles,
    /// content controls, fields, notes, bookmarks) so the two cannot collide, unions the
    /// embedded-image maps, and inserts a section seam per <paramref name="options"/>.
    /// Returns a NEW handle; both inputs are unchanged. <paramref name="other"/> MUST
    /// belong to the SAME engine (they live in one V8 isolate).</summary>
    public WordDocument Append(WordDocument other, MergeOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(other);
        var res = (ScriptObject)_engine.Api.InvokeMethod(
            "mergeDocuments",
            Doc, Images ?? (object)Undefined.Value,
            other.Doc, other.Images ?? (object)Undefined.Value,
            options?.ToJs(_engine) ?? (object)Undefined.Value);
        return FromMerge(_engine, res);
    }

    internal static WordDocument FromMerge(WordCanvasEngine engine, ScriptObject handle)
    {
        var doc = handle.GetProperty("doc");
        var images = handle.GetProperty("images");
        var mediaCount = Convert.ToInt32(handle.GetProperty("mediaCount"));
        var warnings = ReadWarnings(handle.GetProperty("warnings") as ScriptObject);
        return new WordDocument(engine, doc, images, engine.CountBlocks(doc), mediaCount, warnings);
    }

    private static IReadOnlyList<WordWarning> ReadWarnings(ScriptObject? arr)
    {
        if (arr is null) return Array.Empty<WordWarning>();
        var len = Convert.ToInt32(arr.GetProperty("length"));
        var list = new List<WordWarning>(len);
        for (var i = 0; i < len; i++)
        {
            if (arr.GetProperty(i) is not ScriptObject w) continue;
            var code = w.GetProperty("code")?.ToString() ?? "";
            var detail = w.GetProperty("detail");
            list.Add(new WordWarning(code, detail is Undefined ? null : detail?.ToString()));
        }
        return list;
    }

    // ---- export -------------------------------------------------------------

    /// <summary>Render to a page-accurate PDF (reuses the layout engine).</summary>
    public byte[] ExportPdf() => _engine.ExportBytes(pdf: true, Doc, Images);

    /// <summary>Write to DOCX (hand-rolled OOXML, the inverse of the importer).</summary>
    public byte[] ExportDocx() => _engine.ExportBytes(pdf: false, Doc, Images);

    public Task<byte[]> ExportPdfAsync()
    {
        try { return Task.FromResult(ExportPdf()); }
        catch (Exception ex) { return Task.FromException<byte[]>(ex); }
    }

    public Task<byte[]> ExportDocxAsync()
    {
        try { return Task.FromResult(ExportDocx()); }
        catch (Exception ex) { return Task.FromException<byte[]>(ex); }
    }

    /// <summary>Render to PDF straight into <paramref name="destination"/> via a pooled
    /// buffer (no large output byte[] allocated) — pass a <c>RecyclableMemoryStream</c>
    /// for GC-friendly export under load. Returns the bytes written.</summary>
    public long ExportPdf(Stream destination) => _engine.ExportToStream(pdf: true, Doc, Images, destination);

    /// <summary>Write DOCX straight into <paramref name="destination"/> via a pooled
    /// buffer (no large output byte[] allocated). Returns the bytes written.</summary>
    public long ExportDocx(Stream destination) => _engine.ExportToStream(pdf: false, Doc, Images, destination);
}
