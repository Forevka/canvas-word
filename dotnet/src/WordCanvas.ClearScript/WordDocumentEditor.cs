using Microsoft.ClearScript;

namespace WordCanvas.ClearScript;

/// <summary>
/// A stateful, in-place editor over a document — the write half of
/// WordprocessingDocument-style access. Wraps a JS <c>DocumentEditor</c> instance
/// living in V8 (see shared/src/model/documentEditor.ts), which rides the operation
/// engine: every edit swaps the underlying immutable model and records an inverse,
/// so undo/redo come for free. Fetch the current document with
/// <see cref="ToDocument"/> to query or export it.
/// </summary>
public sealed class WordDocumentEditor
{
    private readonly WordCanvasEngine _engine;
    private readonly WordDocument _source;
    private readonly ScriptObject _editor;

    internal WordDocumentEditor(WordCanvasEngine engine, WordDocument source, ScriptObject editor)
    {
        _engine = engine;
        _source = source;
        _editor = editor;
    }

    /// <summary>Replace all runs of a paragraph with a single run of <paramref name="text"/>.</summary>
    public WordDocumentEditor SetParagraphText(string blockId, string text)
    {
        _editor.InvokeMethod("setParagraphText", NonNull(blockId), NonNull(text));
        return this;
    }

    /// <summary>Insert <paramref name="text"/> at a UTF-16 offset within a paragraph.</summary>
    public WordDocumentEditor InsertText(string blockId, int offset, string text)
    {
        _editor.InvokeMethod("insertText", NonNull(blockId), offset, NonNull(text));
        return this;
    }

    /// <summary>Delete the UTF-16 range [start, end) within a paragraph.</summary>
    public WordDocumentEditor DeleteText(string blockId, int start, int end)
    {
        _editor.InvokeMethod("deleteText", NonNull(blockId), start, end);
        return this;
    }

    /// <summary>Replace the UTF-16 range [start, end) within a paragraph with
    /// <paramref name="text"/> (one undoable step).</summary>
    public WordDocumentEditor ReplaceText(string blockId, int start, int end, string text)
    {
        _editor.InvokeMethod("replaceText", NonNull(blockId), start, end, NonNull(text));
        return this;
    }

    /// <summary>Remove a top-level block (paragraph, table, image…) by id.</summary>
    public WordDocumentEditor RemoveBlock(string blockId)
    {
        _editor.InvokeMethod("removeBlock", NonNull(blockId));
        return this;
    }

    /// <summary>Insert a new paragraph after a top-level reference block, inheriting
    /// its style (minus structural markers).</summary>
    public WordDocumentEditor InsertParagraphAfter(string refBlockId, string text) =>
        InsertParagraph(refBlockId, text, "after");

    /// <summary>Insert a new paragraph before a top-level reference block.</summary>
    public WordDocumentEditor InsertParagraphBefore(string refBlockId, string text) =>
        InsertParagraph(refBlockId, text, "before");

    private WordDocumentEditor InsertParagraph(string refBlockId, string text, string position)
    {
        dynamic options = _engine.NewObject();
        options.position = position;
        _editor.InvokeMethod("insertParagraph", NonNull(refBlockId), NonNull(text), (object)options);
        return this;
    }

    // ---- content controls (SDTs) — the primary templating surface ----

    /// <summary>Merge a patch onto a content control's properties (tag, alias, checked,
    /// placeholder, date format, locks). The control must exist; its <c>type</c> is the
    /// discriminant and is always preserved (not patchable).</summary>
    public WordDocumentEditor SetSdtProps(string id, SdtPropsPatch patch)
    {
        ArgumentNullException.ThrowIfNull(patch);
        _editor.InvokeMethod("setSdtProps", NonNull(id), patch.ToJs(_engine));
        return this;
    }

    /// <summary>Set a checkbox content control's state (☒/☐). Throws JS-side if
    /// <paramref name="id"/> is not a checkbox control.</summary>
    public WordDocumentEditor SetCheckbox(string id, bool isChecked)
    {
        _editor.InvokeMethod("setCheckbox", NonNull(id), isChecked);
        return this;
    }

    /// <summary>Fill a single-paragraph content control's text (inline or block-level) —
    /// the "fill this field" primitive. Preserves the control's ancestry (nesting) and
    /// clears any placeholder. Throws JS-side for controls spanning multiple blocks.</summary>
    public WordDocumentEditor SetSdtText(string id, string text)
    {
        _editor.InvokeMethod("setSdtText", NonNull(id), NonNull(text));
        return this;
    }

    /// <summary>Undo the most recent edit. Returns false when there's nothing to undo.</summary>
    public bool Undo() => Convert.ToBoolean(_editor.InvokeMethod("undo"));

    /// <summary>Redo the most recently undone edit. Returns false when there's nothing to redo.</summary>
    public bool Redo() => Convert.ToBoolean(_editor.InvokeMethod("redo"));

    public bool CanUndo => Convert.ToBoolean(_editor.GetProperty("canUndo"));
    public bool CanRedo => Convert.ToBoolean(_editor.GetProperty("canRedo"));

    /// <summary>A document handle over the editor's CURRENT model — query or export
    /// it. Preserves the source document's embedded image bytes.</summary>
    public WordDocument ToDocument() => _source.WithDoc(_editor.GetProperty("doc"));

    private static string NonNull(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        return value;
    }
}

/// <summary>A patch of content-control properties for <see cref="WordDocumentEditor.SetSdtProps"/>.
/// Only the set fields are applied; the control's <c>type</c> is never patchable (it is
/// preserved JS-side). Mirrors the patchable subset of the model's SdtProps.</summary>
public sealed record SdtPropsPatch
{
    /// <summary>Machine-readable tag (w:tag).</summary>
    public string? Tag { get; init; }
    /// <summary>Title shown on the control's tab (w:alias).</summary>
    public string? Alias { get; init; }
    /// <summary>Checkbox state (checkbox controls).</summary>
    public bool? Checked { get; init; }
    /// <summary>Whether the content is the gray placeholder.</summary>
    public bool? Placeholder { get; init; }
    /// <summary>Date display format (date controls).</summary>
    public string? DateFormat { get; init; }
    /// <summary>w:lock="sdtContentLocked" — contents cannot be edited.</summary>
    public bool? LockContent { get; init; }
    /// <summary>w:lock="sdtLocked" — the control cannot be deleted.</summary>
    public bool? LockControl { get; init; }

    internal ScriptObject ToJs(WordCanvasEngine engine)
    {
        dynamic o = engine.NewObject();
        if (Tag is { } t) o.tag = t;
        if (Alias is { } a) o.alias = a;
        if (Checked is { } c) o.@checked = c;
        if (Placeholder is { } p) o.placeholder = p;
        if (DateFormat is { } d) o.dateFormat = d;
        if (LockContent is { } lc) o.lockContent = lc;
        if (LockControl is { } lk) o.lockControl = lk;
        return (ScriptObject)o;
    }
}
