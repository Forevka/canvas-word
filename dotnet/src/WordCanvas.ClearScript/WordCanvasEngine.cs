using Microsoft.ClearScript;
using Microsoft.ClearScript.JavaScript;
using Microsoft.ClearScript.V8;

namespace WordCanvas.ClearScript;

/// <summary>
/// Hosts the WordCanvas headless pipeline (Builder, DOCX import, PDF/DOCX export)
/// in a V8 engine via ClearScript. One engine owns one V8 isolate; it is NOT
/// thread-safe — create one per thread, or serialize access.
///
/// The document model never crosses the boundary as data: imported / built docs
/// stay inside V8 as opaque <see cref="WordDocument"/> handles, and only binary
/// blobs (docx in, pdf/docx out) are marshalled (via <c>Uint8Array</c>).
/// </summary>
public sealed class WordCanvasEngine : IDisposable
{
    private readonly V8ScriptEngine _engine;
    private readonly ScriptObject _api;
    private readonly ScriptObject _alloc;
    private readonly ScriptObject _newObj;
    private readonly ScriptObject _newArr;
    private readonly ScriptObject _runAsync;
    private readonly ScriptObject _drainMacro;
    private bool _disposed;

    public WordCanvasEngine(WordCanvasEngineOptions? options = null)
    {
        options ??= WordCanvasEngineOptions.Default();
        if (!File.Exists(options.BundlePath))
            throw new WordCanvasException($"JS bundle not found: {options.BundlePath} (build it with frontend/scripts/build-clearscript.mjs)");

        var constraints = new V8RuntimeConstraints();
        if (options.MaxHeapSizeMb > 0)
            constraints.MaxOldSpaceSize = (int)options.MaxHeapSizeMb * 1024 * 1024;

        _engine = new V8ScriptEngine(constraints, V8ScriptEngineFlags.EnableTaskPromiseConversion);
        try
        {
            _engine.Execute(new DocumentInfo("wordcanvas.clearscript.js"), File.ReadAllText(options.BundlePath));
            // Tiny host helpers for native JS value construction (typed-array alloc,
            // plain object/array factories used by the builder wrapper).
            _engine.Execute(
                "globalThis.__wc_alloc=function(n){return new Uint8Array(n);};" +
                "globalThis.__wc_obj=function(){return{};};" +
                "globalThis.__wc_arr=function(){return[];};" +
                // Wrap a promise in a sync-pollable box (the export pipeline is async,
                // but all its async is microtask-only — no real timers/IO — so a host
                // pump loop settles it). See ExportBytes/PumpUntilDone.
                "globalThis.__wc_run=function(p){var b={done:false,value:null,error:null};" +
                "Promise.resolve(p).then(function(v){b.value=v;b.done=true;}," +
                "function(e){b.error=''+((e&&e.stack)||e);b.done=true;});return b;};");

            _api = GetGlobal("WordCanvas") as ScriptObject
                   ?? throw new WordCanvasException("bundle did not define globalThis.WordCanvas");
            _alloc = (ScriptObject)GetGlobal("__wc_alloc");
            _newObj = (ScriptObject)GetGlobal("__wc_obj");
            _newArr = (ScriptObject)GetGlobal("__wc_arr");
            _runAsync = (ScriptObject)GetGlobal("__wc_run");
            _drainMacro = (ScriptObject)GetGlobal("__wc_drainMacro");

            InstallFonts(options.FontsDirectory);
        }
        catch
        {
            _engine.Dispose();
            throw;
        }
    }

    private object GetGlobal(string name) => _engine.Global.GetProperty(name);

    private void InstallFonts(string dir)
    {
        if (!Directory.Exists(dir))
            throw new WordCanvasException($"fonts directory not found: {dir}");
        var files = Directory.GetFiles(dir, "*.ttf");
        if (files.Length == 0)
            throw new WordCanvasException($"no .ttf fonts in {dir} — export/measurement needs the bundled clones");
        foreach (var file in files)
        {
            var bytes = File.ReadAllBytes(file);
            _api.InvokeMethod("installFont", Path.GetFileName(file), AllocBytes(bytes));
        }
    }

    // ---- binary marshalling -------------------------------------------------

    /// <summary>Copy host bytes into a fresh V8 <c>Uint8Array</c>.</summary>
    internal ITypedArray<byte> AllocBytes(byte[] data)
    {
        var u8 = (ITypedArray<byte>)_alloc.Invoke(false, data.Length);
        if (data.Length > 0)
            ((IArrayBufferView)u8).WriteBytes(data, 0, (ulong)data.Length, 0);
        return u8;
    }

    private static byte[] ReadBytes(object? value) => value switch
    {
        IArrayBufferView view => view.GetBytes(),
        _ => throw new WordCanvasException($"expected a Uint8Array result, got {value?.GetType().Name ?? "null"}"),
    };

    // ---- native JS value construction (used by the builder wrapper) ---------

    internal dynamic NewObject() => _newObj.Invoke(false);
    internal dynamic NewArray() => _newArr.Invoke(false);

    // ---- pipeline -----------------------------------------------------------

    /// <summary>Import a .docx into an opaque in-V8 document handle.</summary>
    public WordDocument ImportDocx(byte[] docxBytes)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(docxBytes);
        var handle = (ScriptObject)_api.InvokeMethod("importDocx", AllocBytes(docxBytes));
        return WordDocument.FromImport(this, handle);
    }

    public WordDocument ImportDocx(ReadOnlySpan<byte> docxBytes) => ImportDocx(docxBytes.ToArray());

    /// <summary>Import a .docx from a stream (e.g. a <see cref="MemoryStream"/>).
    /// Reads to the end from the current position; does not dispose the stream.</summary>
    public WordDocument ImportDocx(Stream docxStream)
    {
        ArgumentNullException.ThrowIfNull(docxStream);
        if (docxStream is MemoryStream ms && ms.TryGetBuffer(out var seg))
            return ImportDocx(seg.AsSpan());
        using var buffer = new MemoryStream();
        docxStream.CopyTo(buffer);
        return ImportDocx(buffer.GetBuffer().AsSpan(0, (int)buffer.Length));
    }

    /// <summary>
    /// Run an async JS export to completion synchronously. The export pipeline is
    /// async only through microtasks (no real timers/IO — our banner collapses
    /// setTimeout/setImmediate onto promise microtasks), so we wrap the promise in
    /// a pollable box and pump V8's microtask queue (each <c>Execute</c> triggers a
    /// checkpoint) until it settles. Blocking <c>await</c> would deadlock the single
    /// V8 thread; this does not.
    /// </summary>
    internal byte[] ExportBytes(bool pdf, object doc, object? images)
    {
        ThrowIfDisposed();
        var method = pdf ? "exportPdf" : "exportDocx";
        var value = ResolveValue(_api.InvokeMethod(method, doc, images ?? (object)Undefined.Value), method);
        return ReadBytes(value);
    }

    /// <summary>Drive a JS promise to completion (the pipeline's async is microtask/
    /// timer only — see <see cref="PumpUntilDone"/>) and return its resolved value.</summary>
    internal object? ResolveValue(object promise, string what)
    {
        var box = (ScriptObject)_runAsync.Invoke(false, promise);
        PumpUntilDone(box, what);
        var error = box.GetProperty("error");
        if (error is not (null or Undefined))
            throw new WordCanvasException($"JS {what} failed: {error}");
        return box.GetProperty("value");
    }

    private void PumpUntilDone(ScriptObject box, string what)
    {
        // Reproduce Node's event-loop ordering: drain the V8 microtask queue (each
        // Execute triggers a checkpoint), then run one batch of macrotasks (timers /
        // setImmediate that the stream machinery posts), and repeat. A round that
        // settles nothing AND has no macrotasks left means a genuinely pending
        // promise (real-async dependency we don't support) — fail fast instead of
        // spinning. maxRounds is a final backstop.
        const int maxRounds = 50_000_000;
        var idle = 0;
        long lastMicro = -1;
        for (var round = 0; round < maxRounds; round++)
        {
            _engine.Execute("void 0;"); // microtask checkpoint (drains nextTick + promise jobs)
            if (box.GetProperty("done") is true) return;

            var ran = Convert.ToInt32(_drainMacro.Invoke(false)); // run macrotasks (timers)
            if (box.GetProperty("done") is true) return;

            // Progress = a macrotask ran OR a microtask ran since last round (e.g. a
            // Readable stream draining in nextTick batches reports macroRan==0 but
            // keeps bumping __wc_microRan). Only consecutive rounds with NO progress
            // at all mean a genuinely pending promise — then fail fast.
            var micro = Convert.ToInt64(_engine.Evaluate("globalThis.__wc_microRan"));
            if (ran > 0 || micro != lastMicro)
            {
                idle = 0;
                lastMicro = micro;
            }
            else if (++idle >= 8)
            {
                throw new WordCanvasException(
                    $"{what} did not settle: promise pending with no progress (unsupported real-async dependency)");
            }
        }
        throw new WordCanvasException($"{what} did not settle within the pump budget");
    }

    internal int CountBlocks(object doc) => Convert.ToInt32(((dynamic)_api).countBlocks(doc));

    internal ScriptObject Api => _api;

    // ---- builder ------------------------------------------------------------

    /// <summary>Start a blank document (default stylesheet, US Letter, 1in margins).</summary>
    public Builder.DocumentBuilder NewBuilder(Builder.CreateOptions? options = null)
        => Builder.DocumentBuilder.Create(this, options);

    /// <summary>Start from a .docx template: its stylesheet, list definitions, page
    /// setup, and header/footer bands carry over (body discarded unless KeepBody).</summary>
    public Builder.DocumentBuilder NewBuilderFromTemplate(byte[] templateDocx, Builder.TemplateOptions? options = null)
        => Builder.DocumentBuilder.FromTemplate(this, templateDocx, options);

    /// <summary>Pin (or clear) the PDF export timestamp for byte-reproducible output.
    /// When set, CreationDate/ModDate and the content-derived /ID become deterministic;
    /// pass <c>null</c> to restore live wall-clock dates (the default).</summary>
    public void SetExportDate(DateTimeOffset? date)
    {
        var value = date is { } d ? d.ToUnixTimeMilliseconds().ToString() : "undefined";
        _engine.Execute($"globalThis.__wc_fixedDate = {value};");
    }

    /// <summary>Current V8 isolate heap usage (used, total committed) in bytes.</summary>
    public (long UsedBytes, long TotalBytes) V8HeapBytes()
    {
        var info = _engine.GetRuntimeHeapInfo();
        return ((long)info.UsedHeapSize, (long)info.TotalHeapSize);
    }

    /// <summary>The bundle's console ring buffer (warn/error from pdfkit/measureHost).</summary>
    public IReadOnlyList<string> Logs()
    {
        if (_engine.Evaluate("globalThis.__wc_logs") is not ScriptObject arr) return Array.Empty<string>();
        var len = Convert.ToInt32(arr.GetProperty("length"));
        var list = new List<string>(len);
        for (var i = 0; i < len; i++) list.Add(arr.GetProperty(i)?.ToString() ?? "");
        return list;
    }

    // ---- lifetime -----------------------------------------------------------

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(WordCanvasEngine));
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _engine.Dispose();
    }
}
