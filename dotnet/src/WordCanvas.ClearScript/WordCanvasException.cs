namespace WordCanvas.ClearScript;

/// <summary>An error originating from the WordCanvas JS engine or its host interop.</summary>
public sealed class WordCanvasException : Exception
{
    public WordCanvasException(string message, Exception? inner = null) : base(message, inner) { }
}

/// <summary>A lossy-mapping note surfaced by import/export (deduplicated upstream).</summary>
public sealed record WordWarning(string Code, string? Detail);

/// <summary>Result of headlessly generating a Word <c>TOC</c> field's content.</summary>
/// <param name="Docx">The patched .docx (original bytes with the TOC field filled in).</param>
/// <param name="Generated">Entries written into the field result.</param>
/// <param name="Headings">Headings detected in the document.</param>
/// <param name="BookmarksSynthesized">Heading bookmarks injected (those that had none).</param>
public sealed record TocGenerateResult(byte[] Docx, int Generated, int Headings, int BookmarksSynthesized);

/// <summary>Result of headlessly recalculating cached TOC/PAGEREF page numbers.</summary>
/// <param name="Docx">The patched .docx (original bytes with page numbers refreshed).</param>
/// <param name="Changed">PAGEREF cached numbers rewritten to a new value.</param>
/// <param name="Skipped">PAGEREF results left untouched (non-arabic cached value).</param>
public sealed record TocRecalcResult(byte[] Docx, int Changed, int Skipped);
