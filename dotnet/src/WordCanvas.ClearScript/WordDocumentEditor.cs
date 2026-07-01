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
