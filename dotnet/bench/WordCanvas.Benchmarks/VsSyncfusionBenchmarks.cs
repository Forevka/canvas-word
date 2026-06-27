using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using WordCanvas.ClearScript;
using SfWordDocument = Syncfusion.DocIO.DLS.WordDocument;

namespace WordCanvas.Benchmarks;

/// <summary>
/// Head-to-head: the WordCanvas (ClearScript-hosted) pipeline vs Syncfusion DocIO /
/// DocIORenderer, over the same real reports, for the same three operations — open a
/// .docx, export to PDF, export to DOCX. Benchmarks are grouped by operation with
/// WordCanvas as the baseline, so the report shows a direct Ratio per document.
///
/// Run:  dotnet run -c Release --project ... -- --filter *VsSyncfusion*
/// Set SYNCFUSION_LICENSE_KEY (Essential Studio v33) first, or Syncfusion runs in
/// trial mode (watermark) and the numbers are not representative.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(launchCount: 1, warmupCount: 1, iterationCount: 5)]
[GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
[CategoriesColumn]
[HideColumns("Error", "StdDev", "Median", "RatioSD")]
public class VsSyncfusionBenchmarks
{
    public static IEnumerable<DocCase> Cases() => ImportExportBenchmarks.Cases();

    [ParamsSource(nameof(Cases))]
    public DocCase Case { get; set; } = null!;

    private WordCanvasEngine _engine = null!;
    private byte[] _bytes = null!;
    private WordDocument _wcDoc = null!;     // WordCanvas: imported once, reused for export
    private SfWordDocument _sfDoc = null!;   // Syncfusion: opened once, reused for export

    [GlobalSetup]
    public void Setup()
    {
        if (!SyncfusionOps.EnsureLicense())
        {
            Console.Error.WriteLine(
                "WARNING: SYNCFUSION_LICENSE_KEY not set — Syncfusion is running in TRIAL mode " +
                "(watermark + license overhead); comparison numbers are not representative.");
        }
        _engine = new WordCanvasEngine();
        _bytes = File.ReadAllBytes(Case.Path);
        _wcDoc = _engine.ImportDocx(_bytes);
        _sfDoc = SyncfusionOps.Open(_bytes);
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        _sfDoc?.Dispose();
        _engine?.Dispose();
    }

    // ---- Import / open --------------------------------------------------------
    [BenchmarkCategory("Import"), Benchmark(Baseline = true, Description = "WordCanvas")]
    public int Wc_Import() => _engine.ImportDocx(_bytes).BlockCount;

    [BenchmarkCategory("Import"), Benchmark(Description = "Syncfusion")]
    public int Sf_Import()
    {
        using var d = SyncfusionOps.Open(_bytes);
        return d.Sections.Count;
    }

    // ---- Export PDF -----------------------------------------------------------
    [BenchmarkCategory("ExportPdf"), Benchmark(Baseline = true, Description = "WordCanvas")]
    public int Wc_ExportPdf() => _wcDoc.ExportPdf().Length;

    [BenchmarkCategory("ExportPdf"), Benchmark(Description = "Syncfusion")]
    public int Sf_ExportPdf() => SyncfusionOps.ConvertToPdf(_sfDoc);

    // ---- Export DOCX ----------------------------------------------------------
    [BenchmarkCategory("ExportDocx"), Benchmark(Baseline = true, Description = "WordCanvas")]
    public int Wc_ExportDocx() => _wcDoc.ExportDocx().Length;

    [BenchmarkCategory("ExportDocx"), Benchmark(Description = "Syncfusion")]
    public int Sf_ExportDocx() => SyncfusionOps.SaveDocx(_sfDoc);
}
