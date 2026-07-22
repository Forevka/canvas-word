namespace WordCanvas.ClearScript;

/// <summary>
/// Vertical metrics for a <see cref="CustomFont"/>, as fractions of em (the same
/// role as the built-in clones' metrics). <b>Required</b> so that measurement,
/// layout, and export all compute identical line heights — and therefore identical
/// pagination — for the custom face. Supply the values your font is designed
/// against (typically its <c>hhea</c>/<c>OS-2</c> ascent &amp; descent divided by
/// units-per-em); they must match whatever an editor instance uses for the same
/// family, or the headless output will paginate differently.
/// </summary>
/// <param name="Ascent">Ascent as a fraction of em (must be &gt; 0, e.g. 0.9).</param>
/// <param name="Descent">Descent as a fraction of em (must be &gt;= 0, e.g. 0.25).</param>
public readonly record struct FontSizing(double Ascent, double Descent);

/// <summary>
/// A caller-supplied font the engine registers so a document that references
/// <see cref="Family"/> is measured and PDF-subset-embedded as <i>itself</i> (never
/// substituted to a bundled metric clone). This is the bare-V8 analogue of the
/// browser/Node <c>fonts</c> config: there the faces are URLs the host fetches; here
/// the host hands the raw <b>TTF/OTF</b> bytes straight in (bare V8 has no fetch).
/// <para>Only <see cref="Regular"/> is required; a missing <see cref="Bold"/> /
/// <see cref="Italic"/> / <see cref="BoldItalic"/> resolves to the regular face
/// (deterministic parity with the editor, over faux-synthesis). WOFF/WOFF2 is not
/// supported — both fontkit (measuring) and pdfkit (embedding) need uncompressed
/// SFNT bytes.</para>
/// <para>Register via <see cref="WordCanvasEngineOptions.CustomFonts"/> (preferred —
/// pool-safe, applied to every engine built from those options) or
/// <see cref="WordCanvasEngine.RegisterCustomFont"/>.</para>
/// </summary>
public sealed class CustomFont
{
    /// <summary>Family name as referenced by the document's runs (builder
    /// <c>.Font("…")</c> or an imported .docx). Matching is case/quote-insensitive.</summary>
    public required string Family { get; init; }

    /// <summary>Required vertical metrics — see <see cref="FontSizing"/>.</summary>
    public required FontSizing Sizing { get; init; }

    /// <summary>The regular face's TTF/OTF bytes (required).</summary>
    public required byte[] Regular { get; init; }

    /// <summary>The bold face's bytes (optional; falls back to <see cref="Regular"/>).</summary>
    public byte[]? Bold { get; init; }

    /// <summary>The italic face's bytes (optional; falls back to <see cref="Regular"/>).</summary>
    public byte[]? Italic { get; init; }

    /// <summary>The bold-italic face's bytes (optional; falls back to <see cref="Regular"/>).</summary>
    public byte[]? BoldItalic { get; init; }

    /// <summary>Optional label (defaults to <see cref="Family"/>); carried through for
    /// parity with the browser config, unused by the headless host.</summary>
    public string? Label { get; init; }
}
