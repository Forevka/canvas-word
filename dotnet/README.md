# WordCanvas.ClearScript — .NET bindings for the headless WordCanvas pipeline

Run the canvas-word **layout / import / export** pipeline from C# — DOCX import,
page-accurate PDF export, DOCX export, and the fluent **DocumentBuilder** — without
a browser, a Node process, or any port of the engine. The exact same JavaScript that
the editor and the Node backend run is hosted inside a V8 isolate via
[ClearScript](https://github.com/microsoft/ClearScript); C# marshals only bytes
across the boundary.

```
┌──────────────── .NET process ─────────────────┐
│  WordCanvasEngine (ClearScript V8 isolate)     │
│   • loads wordcanvas.clearscript.js (esbuild)  │
│   • host-injects the bundled .ttf fonts        │
│   • DocumentBuilder / Import / Export (typed)  │
│                                                │
│   byte[] docx ─▶ [ V8: runImport ] ─▶ doc ─┐   │
│                                            │   │
│   doc ─▶ [ V8: renderPdf / writeDocx ] ─▶ byte[]│
└────────────────────────────────────────────────┘
```

The document model never crosses the boundary as data: an imported or built document
stays inside V8 as an opaque `WordDocument` handle, and only binary blobs (the docx
in, the pdf/docx out) are marshalled (as `Uint8Array`, zero base64).

## How it runs the browser engine under bare V8

The pipeline is the *isomorphic* one described in `EXPORT.md` — it already runs in a
Web Worker and on Node with no DOM. Hosting it in ClearScript's bare V8 (no Node, no
DOM, no event loop) needed three things, all confined to the JS bundle + the host:

1. **Fonts by injection, not I/O.** `installMeasureHost`'s `fs`/`fetch` font read is
   skipped because the host pre-registers the bundled metric-clone + math fonts via
   `registerFont(file, bytes)` — so measurement uses the same fontkit-over-clones path
   as the Node backend, byte-for-byte, with zero I/O.
2. **A host-owned scheduler.** Bare V8 has no `process`, timers, `TextEncoder`,
   `navigator`, `console`, etc. The esbuild banner supplies them: `process.nextTick`
   and `queueMicrotask` are true microtasks; `setTimeout`/`setImmediate` enqueue to a
   host-drainable macrotask queue. The host then drives a Node-like loop — a V8
   microtask checkpoint, then a macrotask drain, repeated — to settle the export's
   async chain (see `PumpUntilDone`). Blocking `await` would deadlock the single V8
   thread; this does not.
3. **Synchronous PDF collection.** pdfkit's `PDFDocument` is a Node `Readable` whose
   flowing-mode emission deadlocks mid-stream without a real event loop. Since pdfkit
   generates everything synchronously (`deflateSync`), the bundle entry patches it to
   capture pushed chunks and replay them as synchronous `data`+`end` — lossless, and
   it sidesteps the stream machinery entirely.

## Layout

```
dotnet/
  src/WordCanvas.ClearScript/      the binding library (net10.0, x64)
    WordCanvasEngine.cs            V8 host: load bundle, inject fonts, pump, marshal
    WordDocument.cs                opaque in-V8 doc handle + ExportPdf/ExportDocx
    Builder/                       typed DocumentBuilder/StoryBuilder/… over the JS API
    assets/wordcanvas.clearscript.js   the esbuild bundle (generated, gitignored)
    fonts/                         bundled .ttf clones (linked from frontend at build)
  bench/WordCanvas.Benchmarks/     BenchmarkDotNet suite over the real reports
frontend/
  src/clearscript/entry.ts         the bundle entry (exposes the pipeline on globalThis)
  scripts/build-clearscript.mjs    esbuild → assets/wordcanvas.clearscript.js
```

## Build

The JS bundle is produced by esbuild from the frontend workspace, then the .NET
projects copy it (and the fonts) to their output:

```sh
# 1. one-time: install the polyfill plugin (already in frontend devDependencies)
npm install

# 2. build the ClearScript bundle (→ dotnet/src/WordCanvas.ClearScript/assets/…)
node frontend/scripts/build-clearscript.mjs

# 3. build the .NET solution
dotnet build dotnet/WordCanvas.sln -c Release
```

> Requires the Windows x64 ClearScript V8 native package (referenced by the csproj)
> and .NET 10. The bundled fonts are linked from `frontend/src/export/shared/fonts`.

## Usage

### Import a .docx (incl. from a `MemoryStream`) and export

```csharp
using WordCanvas.ClearScript;

using var engine = new WordCanvasEngine();           // one V8 isolate; not thread-safe

// from bytes, a span, or any Stream (e.g. a MemoryStream / HTTP body)
WordDocument doc = engine.ImportDocx(File.ReadAllBytes("report.docx"));
// using var ms = new MemoryStream(uploadBytes); var doc = engine.ImportDocx(ms);

byte[] pdf  = doc.ExportPdf();                        // page-accurate PDF
byte[] docx = doc.ExportDocx();                       // round-tripped OOXML

Console.WriteLine($"{doc.BlockCount} blocks, {doc.MediaCount} images, {doc.Warnings.Count} warnings");
```

### Compose a document with the typed builder

```csharp
using WordCanvas.ClearScript.Builder;

WordDocument doc = engine.NewBuilder(new CreateOptions { PageSize = PageSizeName.A4 })
    .Style(new NamedStyle { Id = "Heading1", Name = "Heading 1",
                            Char = new CharStyle { Bold = true, FontSizePx = 28 } })
    .Paragraph("Quarterly Report", p => p.WithStyle("Heading1"))
    .Paragraph("Generated from C#.", p => p.Italic().Color("#555"))
    .TableOfContents(new TocOptions { MaxLevel = 2 })
    .BulletList("First", "Second", "Third")
    .Table(new[]
    {
        new CellContent[] { "Item", "Qty" },
        new CellContent[] { "Widget", "10" },
    }, new TableOptions { HeaderRow = true })
    .Footer(f => f.Paragraph(null, p => p.Text("Page ").PageField().Text(" of ").NumPagesField()))
    .Build();

byte[] pdf = doc.ExportPdf();
```

Or start from a template (its styles, lists, page setup, and bands carry over):

```csharp
var b = engine.NewBuilderFromTemplate(File.ReadAllBytes("template.docx"));
var doc = b.Paragraph("Body generated against the template's styles").Build();
```

### Headless TOC / field calculation (the Syncfusion replacement)

`TOC`, `PAGEREF`, and page-number fields can't be evaluated by OpenXML /
`WordprocessingDocument` — computing them needs a **layout engine** to know what
lands on which page. WordCanvas has one, so it can do headlessly what teams otherwise
reach for Syncfusion to do. Both operations are **drift-free**: only the field result
/ cached numbers are rewritten in the original `document.xml`; every other byte (and
every field your model doesn't represent) is preserved.

**Generate a `TOC` field's content** — for a .docx that carries a TOC field with an
empty/placeholder result (e.g. emitted by your C# pipeline). Produces the entries Word
would render on F9, with live `PAGEREF` page numbers + hyperlinks:

```csharp
var result = engine.GenerateToc(File.ReadAllBytes("report.docx"),
                                new TocOptions { MaxLevel = 3 });
File.WriteAllBytes("report.docx", result.Docx);
// result.Generated / result.Headings / result.BookmarksSynthesized
```

**Recalculate cached page numbers** — Word's F9 for `TOC`/`PAGEREF` numbers that
already exist but are stale:

```csharp
var result = engine.RecalcTocPageNumbers(stream);   // byte[] or Stream
// result.Changed = how many cached numbers were rewritten; result.Skipped = non-arabic
```

Both lay the document out with the same engine the PDF export uses, so the page numbers
match the rendered pages. Measured on the real reports: generate a 30-entry TOC in
~1.2 s; recalc page numbers in ~0.5–1.3 s.

## Benchmark

```sh
dotnet run -c Release --project dotnet/bench/WordCanvas.Benchmarks -- --filter *
```

[BenchmarkDotNet](https://benchmarkdotnet.org/) times **Import**, **Export PDF**, and
**Export DOCX** over the real reports in `frontend/` (`MemoryDiagnoser` on). The
corpus directory is auto-discovered; override with the `WORDCANVAS_DOCS` env var.

### Results

<!-- BENCHMARK_RESULTS -->
BenchmarkDotNet v0.15.8 · Windows 11 · 11th Gen Intel Core i7-11700K · .NET 10.0.9 X64
RyuJIT. `LaunchCount=1 WarmupCount=2 IterationCount=6`. Mean times; everything runs
through the V8-hosted JS engine.

| Document                          | Size   | Blocks | Import   | Export PDF | Export DOCX |
|-----------------------------------|-------:|-------:|---------:|-----------:|------------:|
| PARTY INVITATION (1 large image)  | 13 MB  |     11 |  12.5 ms |    42.4 ms |    299.9 ms |
| RenderedReport Version-5 (32408)  | 2.7 MB |  1249  |  75.6 ms |   234.8 ms |    170.7 ms |
| RenderedReport Version-5 (no TOC) | 2.7 MB |  1220  | 103.8 ms |   234.3 ms |    106.8 ms |
| AppraiseRequest Version-101       | 9.0 MB |  3353  | 269.1 ms |   406.2 ms |    533.5 ms |
| SignedReport Version-3            | 9.6 MB |  2062  | 194.2 ms |   397.2 ms |    462.3 ms |

Export benchmarks reuse a document imported once in `[GlobalSetup]`, so they measure
export alone. (The 13 MB invitation is one giant image with almost no text, hence its
fast import/PDF but heavier DOCX re-zip.)
<!-- /BENCHMARK_RESULTS -->

These are end-to-end host→V8→host times including all marshalling — a 9 MB, 3353-block
real estate report imports in ~270 ms and renders a page-accurate PDF in ~0.4 s.

## Comparison vs Syncfusion

`VsSyncfusionBenchmarks` runs the WordCanvas pipeline **head-to-head against
Syncfusion** ([DocIO](https://www.syncfusion.com/document-processing/word-framework/net) +
DocIORenderer) over the same reports and the same three operations — open a `.docx`,
export to PDF, export to DOCX — grouped by operation with WordCanvas as the baseline,
so the report shows a direct **Ratio** per document.

**1. Provide a Syncfusion license key (Essential Studio v33 — matches the pinned
`Syncfusion.*` v33.2.15 packages):**

```powershell
$env:SYNCFUSION_LICENSE_KEY = "<your v33 community/trial/commercial key>"
```

Without a key Syncfusion runs in **trial mode** (watermark + license overhead) and the
numbers are not representative — the benchmark prints a warning in that case.

**2. Run the comparison:**

```sh
dotnet run -c Release --project dotnet/bench/WordCanvas.Benchmarks -- --filter *VsSyncfusion*
# quick non-statistical Syncfusion-only timings:
dotnet run -c Release --project dotnet/bench/WordCanvas.Benchmarks -- sfsmoke
```

### Results

<!-- SYNCFUSION_RESULTS -->
BenchmarkDotNet v0.15.8 · i7-11700K · .NET 10.0.9 · Syncfusion v33.2.15 (licensed).
Mean times; "Speedup" = Syncfusion ÷ WordCanvas (>1 → WordCanvas faster).

**Export PDF** — WordCanvas wins decisively (and allocates 11–335× less):

| Document | Blocks | WordCanvas | Syncfusion | Speedup |
|---|--:|--:|--:|--:|
| SignedReport (9.4 MB)    | 2062 | **0.42 s** | 6.89 s | **16.6×** |
| PARTY INVITATION (13 MB) |   11 | **47 ms**  | 0.43 s | **9.1×**  |
| AppraiseRequest (9 MB)   | 3353 | **0.53 s** | 1.07 s | **2.0×**  |
| RenderedReport (2.7 MB)  | 1249 | **0.28 s** | 0.90 s | **3.2×**  |
| RenderedReport (2.7 MB)  | 1220 | **0.23 s** | 0.76 s | **3.2×**  |

**Import / open** — WordCanvas 1.3–1.8× faster, **~700–1500× less allocation**
(AppraiseRequest: 148 KB vs 182 MB):

| Document | WordCanvas | Syncfusion | Speedup |
|---|--:|--:|--:|
| AppraiseRequest (9 MB)  | **281 ms** | 500 ms | 1.8× |
| SignedReport (9.4 MB)   | **187 ms** | 289 ms | 1.5× |
| RenderedReport (2.6 MB) | **78 ms**  | 121 ms | 1.6× |

**Export DOCX** — Syncfusion is faster here (2–4×), though WordCanvas still allocates
10–14× less:

| Document | WordCanvas | Syncfusion | Speedup |
|---|--:|--:|--:|
| AppraiseRequest (9 MB)  | 565 ms | **136 ms** | 0.24× |
| SignedReport (9.4 MB)   | 458 ms | **142 ms** | 0.31× |
| RenderedReport (2.7 MB) | 103 ms | **49 ms**  | 0.48× |

**Takeaway:** WordCanvas dominates the layout-bound operation (PDF, where it reuses
its own engine instead of a full re-render) and import, with order-of-magnitude lower
memory throughout. Syncfusion's hand-tuned OOXML writer is faster at plain DOCX save.
<!-- /SYNCFUSION_RESULTS -->

**Fairness notes.** Both sides do the same logical work on the same machine and corpus.
Differences to keep in mind when reading the numbers:

- **Fonts/layout.** WordCanvas measures + embeds bundled metric-clone fonts (no system
  fonts; deterministic, identical to its Node backend). Syncfusion lays out with the
  machine's installed fonts. So the PDFs differ visually — this compares *throughput*,
  not pixel parity.
- **Import.** WordCanvas import = parse → in-V8 document model (crosses the C#↔V8
  boundary); Syncfusion open = parse → in-memory DOM. Both produce a reusable document.
- **Export.** Both export benchmarks reuse a document loaded once in `[GlobalSetup]`.
  WordCanvas PDF reuses its own layout engine + pdfkit; Syncfusion PDF uses
  DocIORenderer. WordCanvas timings include host↔V8 marshalling of the output bytes.

## Notes & limitations

- **One engine per thread.** A `WordCanvasEngine` owns one V8 isolate and is not
  thread-safe; create one per thread or serialize access. Engine construction loads a
  ~6 MB bundle + 25 fonts, so reuse it across documents.
- **Memory.** Converting all five reports (13 + 9.6 + 9 + 2×2.7 MB) back-to-back in one
  engine peaks at ~700 MB working set: V8 heap ~100 MB used / ~160 MB committed, .NET
  managed ~67 MB (the transient output `byte[]`s), and ~480 MB of fixed V8 + ClearScript
  native + CLR + bundle/font baseline paid once per engine. `WordCanvasEngine.V8HeapBytes()`
  reports live heap usage. The per-document marginal cost is the model + the output blob;
  drop each `WordDocument`/result between docs to keep a large batch flat.
- **Fonts.** Only the bundled Latin metric clones (+ STIX math) are embedded, matching
  the Node export. CJK / complex scripts render as tofu in PDF (same gap as the
  backend). The model keeps original family names.
- **PDF bytes are byte-identical to the Node host** (same bundle) in deterministic
  mode. `WordCanvasEngine.SetExportDate(fixedDate)` pins the `CreationDate`/`ModDate`
  and the trailer `/ID` (the latter is otherwise host-divergent — pdfkit derives it
  from `Buffer.from(md5WordArray)`, which differs between Node's native `Buffer` and
  the V8 polyfill — so deterministic mode pins it to a content hash instead). Verified
  on all five real reports via `frontend/scripts/dump-pdfs.mjs` (Node host) +
  `pdfdump` (ClearScript host) + `compare-pdfs.mjs` → **ALL IDENTICAL**. Without a
  fixed date, output is live-dated and not reproducible (same as any pdfkit export).

## Verifying Node ≡ ClearScript output

The same JS bundle runs under Node and under V8, so their output can be byte-compared
directly (this is also how a UTF-16BE font-name decode bug in the V8 `TextDecoder`
polyfill was found and fixed):

```sh
node frontend/scripts/dump-pdfs.mjs out-node                                  # Node host
dotnet run -c Release --project dotnet/bench/WordCanvas.Benchmarks -- pdfdump out-cs   # V8 host
node frontend/scripts/compare-pdfs.mjs out-node out-cs                        # → ALL IDENTICAL
```
