using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

// Validate a WordprocessingML (.docx) file against the OOXML schema, the same check
// Microsoft Word applies before opening a file. Prints every schema error (node +
// description + part) and exits 1 if any are found, 0 otherwise. Used by the
// `ooxml-validate` CI job on the exported showcase document. See issue #195.
internal static class Program
{
    private static int Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.Error.WriteLine("usage: OoxmlValidate <path-to.docx>");
            return 2;
        }

        var path = args[0];
        if (!File.Exists(path))
        {
            Console.Error.WriteLine($"OoxmlValidate: file not found: {path}");
            return 2;
        }

        var validator = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019);
        // A ValidationErrorInfo's Node/Part lazily reference the open package, so read
        // everything we need into plain strings BEFORE the document is disposed.
        var lines = new List<string>();
        using (var doc = WordprocessingDocument.Open(path, false))
        {
            foreach (var e in validator.Validate(doc))
            {
                var n = lines.Count + 1;
                lines.Add($"#{n} [{e.ErrorType}] {e.Id}\n  {e.Description}\n  node: {e.Node?.LocalName ?? "?"}   part: {e.Part?.Uri?.ToString() ?? "?"}");
                if (n >= 100)
                {
                    lines.Add("… (further errors truncated)");
                    break;
                }
            }
        }

        if (lines.Count == 0)
        {
            Console.WriteLine($"OoxmlValidate: OK — {Path.GetFileName(path)} is schema-valid (0 errors).");
            return 0;
        }

        Console.Error.WriteLine($"OoxmlValidate: {lines.Count} schema error(s) in {Path.GetFileName(path)}:");
        foreach (var line in lines) Console.Error.WriteLine("\n" + line);
        return 1;
    }
}
