using Microsoft.ClearScript;

namespace WordCanvas.ClearScript;

/// <summary>
/// Resolves external ("Link to File") image URLs to their bytes for PDF export.
/// Given every distinct linked-image URL the document references (a:blip r:link →
/// TargetMode="External"), return the bytes for the ones you fetched — omit any URL
/// you couldn't resolve (that image renders as a placeholder box, exactly as with no
/// resolver at all).
/// <para>The HOST owns the fetch: wire your own <c>HttpClient</c> with the proxy,
/// headers, TLS, auth and timeout policy you need, and cap concurrency + downscale to
/// bound memory (120 full-res photos decoded at once is easily hundreds of MB).</para>
/// <para>Only PDF export uses this — DOCX export re-emits linked images as r:link and
/// needs no bytes.</para>
/// </summary>
public delegate IReadOnlyDictionary<string, byte[]> LinkedImageResolver(IReadOnlyList<string> urls);

/// <summary>Async variant of <see cref="LinkedImageResolver"/> for HttpClient-first
/// backends: awaited before the (synchronous) render, so you can fan out downloads
/// with <c>Task.WhenAll</c> under your own <c>SemaphoreSlim</c> cap.</summary>
public delegate Task<IReadOnlyDictionary<string, byte[]>> LinkedImageResolverAsync(IReadOnlyList<string> urls);

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
    /// equivalent of Word's "insert file at end". Reconciles every id space (styles, lists, table styles,
    /// content controls, fields, notes, bookmarks) so the two cannot collide, unions the
    /// embedded-image maps, and inserts a section seam per <paramref name="options"/>.
    /// Returns a NEW handle; both inputs are unchanged. <paramref name="other"/> MUST
    /// belong to the SAME engine (they live in one V8 isolate).</summary>
    public WordDocument Append(WordDocument other, MergeOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(other);
        if (!ReferenceEquals(other._engine, _engine))
            throw new WordCanvasException("Append requires both documents to belong to the same WordCanvasEngine.");
        var res = (ScriptObject)_engine.Api.InvokeMethod(
            "mergeDocuments",
            Doc, Images ?? (object)Undefined.Value,
            other.Doc, other.Images ?? (object)Undefined.Value,
            options?.ToJs(_engine) ?? (object)Undefined.Value);
        return FromMerge(_engine, res);
    }

    /// <summary>Replace a block-level content control's entire content with another
    /// document — the templating primitive: find control <paramref name="sdtId"/> (e.g.
    /// via <see cref="GetSdtsByTag"/>), drop what's inside it, and splice in
    /// <paramref name="source"/>'s content, reconciling every id space and unioning the
    /// embedded-image maps. The control's ancestry is preserved (the source's own nested
    /// controls survive inside it). Returns a NEW handle; both inputs are unchanged.
    /// The control must be block-level at the top level of the body (an inline control,
    /// or one in a table cell / band / note, throws JS-side). <paramref name="source"/>
    /// MUST belong to the SAME engine.</summary>
    public WordDocument ReplaceSdtContent(string sdtId, WordDocument source, MergeOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(sdtId);
        ArgumentNullException.ThrowIfNull(source);
        if (!ReferenceEquals(source._engine, _engine))
            throw new WordCanvasException("ReplaceSdtContent requires both documents to belong to the same WordCanvasEngine.");
        var res = (ScriptObject)_engine.Api.InvokeMethod(
            "replaceSdtContent",
            Doc, Images ?? (object)Undefined.Value,
            sdtId,
            source.Doc, source.Images ?? (object)Undefined.Value,
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

    /// <summary>Render to a page-accurate PDF (reuses the layout engine). External
    /// ("Link to File") images are NOT fetched here — they render as placeholder boxes;
    /// pass a <see cref="LinkedImageResolver"/> overload to embed them.</summary>
    public byte[] ExportPdf() => _engine.ExportBytes(pdf: true, Doc, Images);

    /// <summary>Every distinct external ("Link to File") image URL this document
    /// references (ImageBlock.externalSrc). Inspect or pre-fetch them yourself, or let
    /// an <see cref="ExportPdf(LinkedImageResolver)"/> overload drive your resolver.</summary>
    public IReadOnlyList<string> LinkedImageUrls() => _engine.LinkedImageUrls(Doc);

    /// <summary>Render to PDF, resolving external ("linked") images through a host-owned
    /// <paramref name="imagesResolver"/> — your own HttpClient / proxy / concurrency
    /// policy. Any linked URL the resolver omits still renders as a placeholder box.
    /// With no linked images (or an empty result) this behaves like <see cref="ExportPdf()"/>.</summary>
    public byte[] ExportPdf(LinkedImageResolver imagesResolver)
    {
        ArgumentNullException.ThrowIfNull(imagesResolver);
        return _engine.ExportBytes(pdf: true, Doc, ResolveExportImages(imagesResolver));
    }

    /// <summary>Stream overload of <see cref="ExportPdf(LinkedImageResolver)"/> (pooled
    /// buffer, no large output byte[]).</summary>
    public long ExportPdf(Stream destination, LinkedImageResolver imagesResolver)
    {
        ArgumentNullException.ThrowIfNull(destination);
        ArgumentNullException.ThrowIfNull(imagesResolver);
        return _engine.ExportToStream(pdf: true, Doc, ResolveExportImages(imagesResolver), destination);
    }

    /// <summary>Async overload of <see cref="ExportPdf(LinkedImageResolver)"/>: the
    /// resolver is awaited (fan out your downloads with <c>Task.WhenAll</c>) before the
    /// synchronous render.</summary>
    public async Task<byte[]> ExportPdfAsync(LinkedImageResolverAsync imagesResolver)
    {
        ArgumentNullException.ThrowIfNull(imagesResolver);
        var images = await ResolveExportImagesAsync(imagesResolver).ConfigureAwait(false);
        return _engine.ExportBytes(pdf: true, Doc, images);
    }

    /// <summary>Async + stream overload — awaits the resolver, then renders into
    /// <paramref name="destination"/> via a pooled buffer.</summary>
    public async Task<long> ExportPdfAsync(Stream destination, LinkedImageResolverAsync imagesResolver)
    {
        ArgumentNullException.ThrowIfNull(destination);
        ArgumentNullException.ThrowIfNull(imagesResolver);
        var images = await ResolveExportImagesAsync(imagesResolver).ConfigureAwait(false);
        return _engine.ExportToStream(pdf: true, Doc, images, destination);
    }

    /// <summary>Enumerate the doc's linked-image URLs, drive the resolver, and marshal
    /// the fetched bytes into the export image map. Falls back to the embedded-media map
    /// when there are no linked images or the resolver returns nothing.</summary>
    private object? ResolveExportImages(LinkedImageResolver resolver)
    {
        var urls = _engine.LinkedImageUrls(Doc);
        if (urls.Count == 0) return Images;
        var resolved = resolver(urls);
        return resolved is { Count: > 0 } ? _engine.BuildExportImages(Images, resolved) : Images;
    }

    private async Task<object?> ResolveExportImagesAsync(LinkedImageResolverAsync resolver)
    {
        var urls = _engine.LinkedImageUrls(Doc);
        if (urls.Count == 0) return Images;
        var resolved = await resolver(urls).ConfigureAwait(false);
        return resolved is { Count: > 0 } ? _engine.BuildExportImages(Images, resolved) : Images;
    }

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
