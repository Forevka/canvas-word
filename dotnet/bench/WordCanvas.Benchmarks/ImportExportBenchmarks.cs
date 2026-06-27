using BenchmarkDotNet.Attributes;
using WordCanvas.ClearScript;

namespace WordCanvas.Benchmarks;

/// <summary>A benchmark subject: one real .docx, displayed by a short name.</summary>
public sealed class DocCase
{
    public required string Path { get; init; }
    public required string Display { get; init; }
    public override string ToString() => Display;
}

/// <summary>
/// Times the headless pipeline (DOCX import, PDF export, DOCX export) over the real
/// reports in <c>frontend/</c>, all running through the ClearScript-hosted JS engine.
/// One V8 engine is created per case in <see cref="Setup"/>; the document is
/// pre-imported so the export benchmarks measure export alone.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(launchCount: 1, warmupCount: 2, iterationCount: 6)]
[HideColumns("Error", "StdDev", "Median", "RatioSD")]
public class ImportExportBenchmarks
{
    public static IEnumerable<DocCase> Cases()
    {
        foreach (var path in DocCorpus.Files())
        {
            // Disambiguating short name: distinctive tail tokens + size (the real
            // report names share long prefixes and would otherwise collide).
            var stem = System.IO.Path.GetFileNameWithoutExtension(path)
                .Replace("RenderedReportsDocx_Report-", "", StringComparison.Ordinal)
                .Replace("SignedReports_SignedReport-", "Signed-", StringComparison.Ordinal);
            var tokens = stem.Split('-', ' ');
            var tail = string.Join("-", tokens.Where(t => t.Length > 0).TakeLast(2));
            var head = tokens.FirstOrDefault(t => t.Length > 0) ?? stem;
            var mb = new FileInfo(path).Length / (1024.0 * 1024.0);
            yield return new DocCase { Path = path, Display = $"{head}…{tail} ({mb:F1}MB)" };
        }
    }

    [ParamsSource(nameof(Cases))]
    public DocCase Case { get; set; } = null!;

    private WordCanvasEngine _engine = null!;
    private byte[] _bytes = null!;
    private WordDocument _imported = null!;

    [GlobalSetup]
    public void Setup()
    {
        _engine = new WordCanvasEngine();
        _bytes = File.ReadAllBytes(Case.Path);
        _imported = _engine.ImportDocx(_bytes); // reused by the export benchmarks
    }

    [GlobalCleanup]
    public void Cleanup() => _engine.Dispose();

    /// <summary>Parse the .docx bytes into the in-V8 document model.</summary>
    [Benchmark(Description = "Import .docx")]
    public int Import() => _engine.ImportDocx(_bytes).BlockCount;

    /// <summary>Render the (pre-imported) document to a page-accurate PDF.</summary>
    [Benchmark(Description = "Export PDF")]
    public int ExportPdf() => _imported.ExportPdf().Length;

    /// <summary>Write the (pre-imported) document back to DOCX (OOXML).</summary>
    [Benchmark(Description = "Export DOCX")]
    public int ExportDocx() => _imported.ExportDocx().Length;
}
