namespace WordCanvas.ClearScript;

/// <summary>An error originating from the WordCanvas JS engine or its host interop.</summary>
public sealed class WordCanvasException : Exception
{
    public WordCanvasException(string message, Exception? inner = null) : base(message, inner) { }
}

/// <summary>A lossy-mapping note surfaced by import/export (deduplicated upstream).</summary>
public sealed record WordWarning(string Code, string? Detail);
