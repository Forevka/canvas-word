namespace WordCanvas.ClearScript;

/// <summary>How colliding named / table styles reconcile with the destination when
/// merging. Mirrors the JS <c>StyleMergeMode</c>.</summary>
public enum StyleMergeMode
{
    /// <summary>Repoint a same-named source style at the destination's; add source-only
    /// styles (Word's "use destination styles"). The default.</summary>
    UseDestination,
    /// <summary>Keep the source's own styles, renaming colliding ids.</summary>
    KeepSource,
}

/// <summary>The section seam inserted between the destination and the appended source.
/// Mirrors the JS <c>MergeSectionBreak</c>.</summary>
public enum MergeSectionBreak
{
    /// <summary>Start the source on a new page; each side keeps its own section
    /// geometry + header/footer bands. The default.</summary>
    NextPage,
    /// <summary>Like <see cref="NextPage"/>, forcing the source's first page onto an even page number.</summary>
    EvenPage,
    /// <summary>Like <see cref="NextPage"/>, forcing the source's first page onto an odd page number.</summary>
    OddPage,
    /// <summary>Append inline under the destination's section (the source's final section
    /// geometry/bands are dropped; its internal breaks are preserved).</summary>
    Continuous,
    /// <summary>Append with no seam at all (the caller manages section breaks).</summary>
    None,
}

/// <summary>Options for <see cref="WordDocument.Append"/> / <see cref="WordCanvasEngine.Merge(WordDocument[])"/>.
/// A direct mirror of the JS <c>MergeOptions</c>.</summary>
public sealed record MergeOptions
{
    /// <summary>Style reconciliation mode (default <see cref="StyleMergeMode.UseDestination"/>).</summary>
    public StyleMergeMode? Styles { get; init; }
    /// <summary>The section seam between the parts (default <see cref="MergeSectionBreak.NextPage"/>).</summary>
    public MergeSectionBreak? SectionBreak { get; init; }
    /// <summary>Rename a source bookmark whose name collides with the destination rather
    /// than dropping it (default true).</summary>
    public bool? RenameBookmarksOnCollision { get; init; }

    internal object ToJs(WordCanvasEngine engine)
    {
        dynamic o = engine.NewObject();
        if (Styles is { } s)
            o.styles = s == StyleMergeMode.UseDestination ? "useDestination" : "keepSource";
        if (SectionBreak is { } sb)
            o.sectionBreak = sb switch
            {
                MergeSectionBreak.NextPage => "nextPage",
                MergeSectionBreak.EvenPage => "evenPage",
                MergeSectionBreak.OddPage => "oddPage",
                MergeSectionBreak.Continuous => "continuous",
                _ => "none",
            };
        if (RenameBookmarksOnCollision is { } r)
            o.renameBookmarksOnCollision = r;
        return o;
    }
}
