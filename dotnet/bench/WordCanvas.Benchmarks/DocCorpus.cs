namespace WordCanvas.Benchmarks;

/// <summary>Locates the real-world .docx test subjects that live in <c>frontend/</c>.</summary>
public static class DocCorpus
{
    /// <summary>Directory holding the benchmark .docx files. Override with the
    /// <c>WORDCANVAS_DOCS</c> env var; otherwise walk up to the repo's
    /// <c>frontend/</c>.</summary>
    public static string DocsDirectory()
    {
        var env = Environment.GetEnvironmentVariable("WORDCANVAS_DOCS");
        if (!string.IsNullOrWhiteSpace(env))
        {
            if (!Directory.Exists(env) || !Directory.EnumerateFiles(env, "*.docx").Any())
                throw new DirectoryNotFoundException(
                    $"WORDCANVAS_DOCS ('{env}') must point to an existing directory containing .docx files.");
            return env;
        }

        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            var frontend = Path.Combine(dir.FullName, "frontend");
            if (Directory.Exists(frontend) && Directory.EnumerateFiles(frontend, "*.docx").Any())
                return frontend;
            dir = dir.Parent;
        }
        throw new DirectoryNotFoundException(
            "Could not find a frontend/ directory with .docx files. Set WORDCANVAS_DOCS to the folder containing the test docx.");
    }

    /// <summary>All .docx files in the corpus directory (sorted, smallest first).</summary>
    public static IReadOnlyList<string> Files()
    {
        var dir = DocsDirectory();
        return Directory.GetFiles(dir, "*.docx")
            .OrderBy(f => new FileInfo(f).Length)
            .ToArray();
    }
}
