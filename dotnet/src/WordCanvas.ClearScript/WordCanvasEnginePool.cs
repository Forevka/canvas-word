using System.Collections.Concurrent;

namespace WordCanvas.ClearScript;

/// <summary>
/// A concurrency-bounded pool of <see cref="WordCanvasEngine"/> instances, for
/// safe reuse in multi-threaded hosts (ASP.NET Core, background workers, etc.).
///
/// <para>A single engine owns one V8 isolate and is <b>not thread-safe</b>: it must
/// never be touched by two threads at once. This pool solves that by leasing an
/// engine to exactly one caller at a time (via <see cref="UseAsync{T}"/> /
/// <see cref="Use{T}"/>) and by capping how many engines run concurrently. The
/// cap bounds three things at once: parallel V8 work, total isolate heap (each
/// engine keeps its own heap), and — since the calls are synchronous, CPU-bound
/// pumps — request-thread pressure.</para>
///
/// <para>Register it as a <b>singleton</b> and share it across requests; the pool
/// itself is thread-safe. Engines are created lazily on demand (construction is
/// heavy — it loads the JS bundle and installs fonts) and reused across leases,
/// so at most <see cref="MaxConcurrency"/> engines are ever live. Dispose the pool
/// on host shutdown to dispose its engines.</para>
///
/// <example>
/// <code>
/// // DI: one shared, thread-safe pool for the whole app.
/// builder.Services.AddSingleton(_ =&gt; new WordCanvasEnginePool(
///     maxConcurrency: 4,
///     engineOptions: new WordCanvasEngineOptions { BundlePath = ..., FontsDirectory = ..., MaxHeapSizeMb = 512 }));
///
/// // Endpoint: lease an engine for the duration of one operation.
/// app.MapPost("/export.pdf", async (Stream body, WordCanvasEnginePool pool, CancellationToken ct) =&gt;
/// {
///     var pdf = await pool.UseAsync(engine =&gt;
///     {
///         var doc = engine.ImportDocx(body);
///         return doc.ExportPdf();
///     }, ct);
///     return Results.File(pdf, "application/pdf");
/// });
/// </code>
/// </example>
/// </summary>
public sealed class WordCanvasEnginePool : IDisposable
{
    private readonly SemaphoreSlim _gate;
    private readonly ConcurrentBag<WordCanvasEngine> _idle = new();
    private readonly WordCanvasEngineOptions _options;
    private bool _disposed;

    /// <summary>The maximum number of engines that may run concurrently — the
    /// ceiling on parallel V8 work and total isolate heap.</summary>
    public int MaxConcurrency { get; }

    /// <summary>
    /// Create a pool.
    /// </summary>
    /// <param name="maxConcurrency">Maximum engines running at once. Size this to the
    /// host: <c>maxConcurrency × MaxHeapSizeMb</c> is roughly the memory ceiling for
    /// document work. Defaults to <see cref="Environment.ProcessorCount"/>.</param>
    /// <param name="engineOptions">Options used to construct each engine (bundle + fonts
    /// location, V8 heap limit). Defaults to <see cref="WordCanvasEngineOptions.Default"/>.
    /// The same options instance is reused for every engine, so treat it as immutable.</param>
    public WordCanvasEnginePool(int? maxConcurrency = null, WordCanvasEngineOptions? engineOptions = null)
    {
        var cap = maxConcurrency ?? Environment.ProcessorCount;
        if (cap < 1)
            throw new ArgumentOutOfRangeException(nameof(maxConcurrency), cap, "maxConcurrency must be at least 1.");

        MaxConcurrency = cap;
        _gate = new SemaphoreSlim(cap, cap);
        _options = engineOptions ?? WordCanvasEngineOptions.Default();
    }

    /// <summary>The number of constructed engines currently sitting idle in the pool
    /// (created but not leased). Diagnostic only.</summary>
    public int IdleCount => _idle.Count;

    /// <summary>
    /// Lease an engine, run <paramref name="work"/> against it on a thread-pool thread
    /// (so the synchronous, CPU-bound pump does not block the caller), and return the
    /// result. The engine is used single-threaded for the whole call and returned to
    /// the pool afterwards. Blocks (asynchronously) until a lease is available when all
    /// engines are busy.
    /// </summary>
    /// <remarks><paramref name="ct"/> is honored only while waiting for a free lease;
    /// once <paramref name="work"/> starts, the synchronous V8 call cannot be
    /// interrupted.</remarks>
    public async Task<T> UseAsync<T>(Func<WordCanvasEngine, T> work, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(work);
        ThrowIfDisposed();

        await _gate.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            // The pool may have been disposed while we waited on the gate — re-check
            // before renting/creating an engine so work never runs on a disposed pool.
            ThrowIfDisposed();
            return await Task.Run(() =>
            {
                var engine = Rent();
                var ok = false;
                try
                {
                    var result = work(engine);
                    ok = true;
                    return result;
                }
                finally
                {
                    Return(engine, healthy: ok);
                }
            }, ct).ConfigureAwait(false);
        }
        finally
        {
            ReleaseGate();
        }
    }

    /// <summary>Lease an engine and run <paramref name="work"/> (no result). See
    /// <see cref="UseAsync{T}"/>.</summary>
    public Task UseAsync(Action<WordCanvasEngine> work, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(work);
        return UseAsync<object?>(engine => { work(engine); return null; }, ct);
    }

    /// <summary>
    /// Synchronous lease — run <paramref name="work"/> against a pooled engine on the
    /// calling thread. Prefer <see cref="UseAsync{T}"/> in request pipelines; use this
    /// from console apps or background workers already off the request path. Blocks
    /// until a lease is available.
    /// </summary>
    public T Use<T>(Func<WordCanvasEngine, T> work, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(work);
        ThrowIfDisposed();

        _gate.Wait(ct);
        try
        {
            // Re-check after acquiring the permit (Dispose may have run while we waited).
            ThrowIfDisposed();
            var engine = Rent();
            var ok = false;
            try
            {
                var result = work(engine);
                ok = true;
                return result;
            }
            finally
            {
                Return(engine, healthy: ok);
            }
        }
        finally
        {
            ReleaseGate();
        }
    }

    /// <summary>Synchronous lease with no result. See <see cref="Use{T}"/>.</summary>
    public void Use(Action<WordCanvasEngine> work, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(work);
        Use<object?>(engine => { work(engine); return null; }, ct);
    }

    /// <summary>
    /// Eagerly construct up to <paramref name="count"/> engines (capped at
    /// <see cref="MaxConcurrency"/>) so the first requests don't pay the bundle-load +
    /// font-install cost. Construction runs on thread-pool threads. Optional.
    /// <para>Prewarming holds one gate permit per engine while it builds, so it obeys the
    /// same bounded-lease invariant as <see cref="UseAsync{T}"/> — engines are only ever
    /// created while a permit is held, so concurrent leases or repeated/concurrent prewarm
    /// calls can never leave more than <see cref="MaxConcurrency"/> live engines. Already-idle
    /// engines are reused, so repeated calls don't over-create.</para>
    /// </summary>
    public async Task PrewarmAsync(int count, CancellationToken ct = default)
    {
        ThrowIfDisposed();
        var target = Math.Min(count, MaxConcurrency);
        if (target <= 0) return;

        var acquired = 0;
        var warmed = new List<WordCanvasEngine>(target);
        try
        {
            // Hold `target` permits for the whole warm-up so total live engines
            // (idle + leased + warming) can't exceed MaxConcurrency.
            for (var i = 0; i < target; i++)
            {
                await _gate.WaitAsync(ct).ConfigureAwait(false);
                acquired++;
            }
            ThrowIfDisposed();
            for (var i = 0; i < target; i++)
            {
                ct.ThrowIfCancellationRequested();
                // Reuse an already-idle engine so repeated prewarm calls don't over-create.
                warmed.Add(_idle.TryTake(out var existing)
                    ? existing
                    : await Task.Run(() => new WordCanvasEngine(_options), ct).ConfigureAwait(false));
            }
            // Only publish once every requested engine is ready.
            foreach (var e in warmed) _idle.Add(e);
        }
        catch
        {
            foreach (var e in warmed) e.Dispose();
            throw;
        }
        finally
        {
            if (acquired > 0) ReleaseGate(acquired);
        }
    }

    private WordCanvasEngine Rent() => _idle.TryTake(out var engine) ? engine : new WordCanvasEngine(_options);

    /// <summary>Release gate permits, tolerating a pool disposed mid-lease (Dispose disposes
    /// the semaphore to fail-fast pending waiters, which can race an in-flight release).</summary>
    private void ReleaseGate(int count = 1)
    {
        try { _gate.Release(count); }
        catch (ObjectDisposedException) { /* pool disposed while a lease was in flight */ }
    }

    private void Return(WordCanvasEngine engine, bool healthy)
    {
        // Never re-pool an engine whose work threw (a faulted isolate could poison the
        // next lease), and never re-pool once the pool is disposed. Dispose instead.
        if (healthy && !Volatile.Read(ref _disposed))
            _idle.Add(engine);
        else
            engine.Dispose();
    }

    private void ThrowIfDisposed()
    {
        if (Volatile.Read(ref _disposed)) throw new ObjectDisposedException(nameof(WordCanvasEnginePool));
    }

    /// <summary>Dispose all idle engines. Engines currently on lease are disposed when
    /// their lease returns. Safe to call once; not safe to lease after disposing.</summary>
    public void Dispose()
    {
        if (_disposed) return;
        Volatile.Write(ref _disposed, true);
        while (_idle.TryTake(out var engine)) engine.Dispose();
        _gate.Dispose();
    }
}
