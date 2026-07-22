namespace WordCanvas.ClearScript;

/// <summary>
/// Locates the headless JS bundle and the bundled font directory. The defaults
/// resolve next to the assembly (the csproj copies <c>assets/</c> and
/// <c>fonts/</c> to the output directory).
/// </summary>
public sealed class WordCanvasEngineOptions
{
    /// <summary>Path to the esbuild IIFE bundle (<c>wordcanvas.clearscript.js</c>).</summary>
    public required string BundlePath { get; init; }

    /// <summary>Directory holding the bundled <c>*.ttf</c> metric-clone + math fonts.</summary>
    public required string FontsDirectory { get; init; }

    /// <summary>Initial V8 heap size limit in MB (0 = V8 default). The big real
    /// reports allocate a lot of layout state; bump this if you hit OOM.</summary>
    public uint MaxHeapSizeMb { get; init; } = 0;

    /// <summary>Custom fonts to register on the engine at construction, so a document
    /// that references their <see cref="CustomFont.Family"/> is measured and embedded
    /// as itself (not substituted). Preferred over
    /// <see cref="WordCanvasEngine.RegisterCustomFont"/> for pooled/prewarmed engines:
    /// every engine built from these options gets the same fonts. Empty by default.</summary>
    public IList<CustomFont> CustomFonts { get; init; } = new List<CustomFont>();

    /// <summary>
    /// Resolves the bundle + fonts next to the assembly, or from the
    /// <c>WORDCANVAS_BUNDLE</c> / <c>WORDCANVAS_FONTS</c> env vars when set (handy
    /// for test/benchmark runners that execute from a generated output directory).
    /// </summary>
    public static WordCanvasEngineOptions Default()
    {
        var baseDir = AppContext.BaseDirectory;
        var bundle = Environment.GetEnvironmentVariable("WORDCANVAS_BUNDLE");
        var fonts = Environment.GetEnvironmentVariable("WORDCANVAS_FONTS");
        return new WordCanvasEngineOptions
        {
            BundlePath = !string.IsNullOrWhiteSpace(bundle) ? bundle : Path.Combine(baseDir, "assets", "wordcanvas.clearscript.js"),
            FontsDirectory = !string.IsNullOrWhiteSpace(fonts) ? fonts : Path.Combine(baseDir, "fonts"),
        };
    }
}
