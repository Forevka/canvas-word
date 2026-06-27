using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Licensing;
using Syncfusion.Pdf;

namespace WordCanvas.Benchmarks;

/// <summary>
/// Thin wrappers over Syncfusion DocIO / DocIORenderer for the head-to-head
/// benchmark: open a .docx, save it back to DOCX, and convert it to PDF — the same
/// three operations the WordCanvas pipeline is measured on.
///
/// The license is registered once from the <c>SYNCFUSION_LICENSE_KEY</c> env var
/// (Syncfusion v33). Without it Syncfusion runs in trial mode (watermark + a license
/// message), so the comparison numbers would not be representative.
/// </summary>
internal static class SyncfusionOps
{
    private static readonly object InitLock = new();
    private static bool _initialized;
    private static bool _hasKey;

    /// <summary>Register the license from the env var (idempotent, thread-safe).
    /// Returns whether a key was actually present.</summary>
    public static bool EnsureLicense()
    {
        if (_initialized) return _hasKey;
        lock (InitLock)
        {
            if (_initialized) return _hasKey;
            var key = Environment.GetEnvironmentVariable("SYNCFUSION_LICENSE_KEY");
            if (!string.IsNullOrWhiteSpace(key))
            {
                SyncfusionLicenseProvider.RegisterLicense(key.Trim());
                _hasKey = true;
            }
            _initialized = true;
            return _hasKey;
        }
    }

    /// <summary>Open (import) a .docx into a Syncfusion WordDocument.</summary>
    public static WordDocument Open(byte[] docx)
    {
        using var ms = new MemoryStream(docx, writable: false);
        return new WordDocument(ms, FormatType.Docx);
    }

    /// <summary>Save a WordDocument back to DOCX; returns the byte length.</summary>
    public static int SaveDocx(WordDocument doc)
    {
        using var ms = new MemoryStream();
        doc.Save(ms, FormatType.Docx);
        return (int)ms.Length;
    }

    /// <summary>Convert a WordDocument to PDF; returns the byte length.</summary>
    public static int ConvertToPdf(WordDocument doc)
    {
        using var renderer = new DocIORenderer();
        using PdfDocument pdf = renderer.ConvertToPDF(doc);
        using var ms = new MemoryStream();
        pdf.Save(ms);
        return (int)ms.Length;
    }
}
