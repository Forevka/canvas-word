// Bundles the headless WordCanvas pipelines into ONE self-contained IIFE for a
// bare-V8 host (.NET ClearScript). Unlike build-node.mjs (platform "node", keeps
// Node builtins external), this targets an environment with NO Node and NO DOM:
//
//   - platform "browser" + esbuild-plugin-polyfill-node supplies Buffer/process/
//     stream/zlib/events/util/etc. that pdfkit + fontkit require.
//   - A banner installs the globals bare V8 lacks: timer shims (collapsed onto
//     promise microtasks — pdfkit/streams only need deferral, never real delays),
//     a minimal `process`, `queueMicrotask`, and a JSON-based `structuredClone`
//     fallback (the document model is pure serializable data, so it's lossless).
//   - Fonts are NOT read here; the host injects them via WordCanvas.installFont,
//     so measureHost's fs/fetch branch is never reached at runtime.
//
// Output: dotnet/src/WordCanvas.ClearScript/assets/wordcanvas.clearscript.js

import { build } from "esbuild";
import { polyfillNode } from "esbuild-plugin-polyfill-node";
import { fileURLToPath } from "node:url";
import { mkdir } from "node:fs/promises";

const here = (p) => fileURLToPath(new URL(p, import.meta.url));

const ENTRY = here("../src/clearscript/entry.ts");
const SHARED = here("../../shared/src/index.ts");
const OUTFILE = here("../../dotnet/src/WordCanvas.ClearScript/assets/wordcanvas.clearscript.js");

await mkdir(here("../../dotnet/src/WordCanvas.ClearScript/assets"), { recursive: true });

// Globals bare V8 doesn't provide. Runs before any bundled module code, so the
// whole runtime (process, timers, encoders) is in place before pdfkit/fontkit/
// readable-stream initialize.
const BANNER = `(function (g) {
  if (!g.globalThis) { g.globalThis = g; }
  // Progress counter: bumped whenever a queued microtask (incl. process.nextTick,
  // which delegates here) actually runs. The host pump watches it to tell ongoing
  // microtask progress (e.g. a Readable stream draining in nextTick batches) apart
  // from a genuinely stalled promise. See PumpUntilDone.
  g.__wc_microRan = 0;
  // Always wrap (even a host-provided queueMicrotask) so process.nextTick's progress
  // is visible to the pump; otherwise a host queueMicrotask would hide it.
  var realQueueMicrotask = g.queueMicrotask
    ? g.queueMicrotask.bind(g)
    : function (fn) { Promise.resolve().then(fn); };
  g.queueMicrotask = function (fn) {
    realQueueMicrotask(function () { g.__wc_microRan++; fn(); });
  };
  // Bare V8 has no event loop. We own the whole scheduler so every async path the
  // export pipeline uses routes through queues the .NET host can drain:
  //   - microtasks (process.nextTick, queueMicrotask) -> native promise jobs, which
  //     the host drains with a V8 microtask checkpoint (engine.Execute).
  //   - macrotasks (setTimeout, setImmediate) -> a host-drainable queue, run AFTER
  //     the microtask queue (Node ordering), so pdfkit's Readable stream actually
  //     emits 'end' for multi-page docs. Timer DELAYS are ignored (deferral is all
  //     the consumers need — there is no real wall-clock wait).
  var macro = [];
  var timerId = 1;
  function schedule(fn, args) {
    var t = { id: timerId++, fn: fn, args: args, cancelled: false };
    macro.push(t);
    return t.id;
  }
  if (!g.setTimeout) { g.setTimeout = function (fn) { return schedule(fn, Array.prototype.slice.call(arguments, 2)); }; }
  if (!g.setImmediate) { g.setImmediate = function (fn) { return schedule(fn, Array.prototype.slice.call(arguments, 1)); }; }
  function cancel(id) { for (var i = 0; i < macro.length; i++) { if (macro[i].id === id) macro[i].cancelled = true; } }
  if (!g.clearTimeout) { g.clearTimeout = cancel; }
  if (!g.clearImmediate) { g.clearImmediate = cancel; }
  // Host hooks: run all currently-queued macrotasks (callbacks may enqueue more,
  // which the next round picks up); report how many are pending.
  g.__wc_drainMacro = function () {
    var q = macro; macro = [];
    var ran = 0;
    for (var i = 0; i < q.length; i++) { if (!q[i].cancelled) { q[i].fn.apply(null, q[i].args); ran++; } }
    return ran;
  };
  g.__wc_macroPending = function () { return macro.length; };
  // Host-owned process (aliased as the 'process'/'node:process' module too). nextTick
  // is a TRUE microtask so stream/zlib continuations drain in a microtask checkpoint.
  if (!g.process) {
    g.process = {
      env: { NODE_ENV: "production" },
      argv: [], platform: "browser", arch: "x64", version: "", versions: { node: "" },
      browser: true,
      nextTick: function (fn) { var a = Array.prototype.slice.call(arguments, 1); g.queueMicrotask(function () { fn.apply(null, a); }); },
      cwd: function () { return "/"; },
      emitWarning: function () {}, on: function () {}, once: function () {}, off: function () {},
      removeListener: function () {}, binding: function () { throw new Error("process.binding is not supported"); },
    };
  }
  // Minimal console routed to a host-readable ring buffer so warnings/errors (and
  // diagnostics) surface. Installed UNCONDITIONALLY — ClearScript provides its own
  // console that we can't read from the host.
  {
    g.__wc_logs = [];
    var log = function (level) {
      return function () {
        var parts = [];
        for (var i = 0; i < arguments.length; i++) { try { parts.push(String(arguments[i])); } catch (e) { parts.push("?"); } }
        g.__wc_logs.push(level + ": " + parts.join(" "));
        if (g.__wc_logs.length > 1000) g.__wc_logs.shift();
      };
    };
    g.console = { log: log("log"), info: log("info"), warn: log("warn"), error: log("error"), debug: log("debug"), trace: log("trace") };
  }
  if (!g.structuredClone) {
    g.structuredClone = function (o) { return JSON.parse(JSON.stringify(o)); };
  }
  // The jspm process polyfill reads navigator.language; pretext's getEngineProfile
  // reads navigator.userAgent/vendor. An empty userAgent + non-Apple vendor yields
  // the SAME (default, non-Safari/non-Chromium) line-break profile the Node backend
  // resolves, so layout/pagination matches the existing Node export exactly.
  if (!g.navigator) {
    g.navigator = { language: "en-US", languages: ["en-US"], userAgent: "", vendor: "", platform: "" };
  }
  // Minimal Blob + URL.createObjectURL: importing a doc WITHOUT collectMediaBytes
  // (e.g. inside the headless TOC recalc functions) mints blob: URLs for images. The
  // layout engine places images by their MODEL dimensions, never the URL, so a dummy
  // is harmless — this just keeps that path from throwing under bare V8.
  if (!g.Blob) { g.Blob = function () {}; }
  if (typeof g.URL === "undefined") { g.URL = {}; }
  if (!g.URL.createObjectURL) { g.URL.createObjectURL = function () { return "blob:wc"; }; }
  if (!g.URL.revokeObjectURL) { g.URL.revokeObjectURL = function () {}; }
  // Bare V8 has no btoa/atob; the builder's bytesToDataUrl base64-encodes image bytes
  // through btoa. Self-contained base64 over binary strings (each char = one byte).
  if (!g.btoa || !g.atob) {
    var B64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    g.btoa = function (s) {
      s = String(s); var out = "", i = 0;
      while (i < s.length) {
        var c1 = s.charCodeAt(i++) & 0xff, c2 = s.charCodeAt(i++), c3 = s.charCodeAt(i++);
        var e2 = ((c1 & 3) << 4) | (isNaN(c2) ? 0 : (c2 & 0xff) >> 4);
        var e3 = isNaN(c2) ? 64 : (((c2 & 15) << 2) | (isNaN(c3) ? 0 : (c3 & 0xff) >> 6));
        var e4 = isNaN(c3) ? 64 : (c3 & 63);
        out += B64[c1 >> 2] + B64[e2] + B64[e3] + B64[e4];
      }
      return out;
    };
    g.atob = function (s) {
      s = String(s).replace(/[^A-Za-z0-9+/]/g, ""); var out = "", i = 0;
      while (i < s.length) {
        var e1 = B64.indexOf(s.charAt(i++)), e2 = B64.indexOf(s.charAt(i++));
        var e3 = B64.indexOf(s.charAt(i++)), e4 = B64.indexOf(s.charAt(i++));
        out += String.fromCharCode((e1 << 2) | (e2 >> 4));
        if (e3 >= 0 && e3 < 64) out += String.fromCharCode(((e2 & 15) << 4) | (e3 >> 2));
        if (e4 >= 0 && e4 < 64) out += String.fromCharCode(((e3 & 3) << 6) | e4);
      }
      return out;
    };
  }
  // The jspm stream polyfill reads AbortController/AbortSignal off self/window and
  // crashes if neither exists. Point self at the global and provide inert stubs
  // (pdfkit never aborts a stream, so these are never actually exercised).
  if (!g.self) { g.self = g; }
  if (!g.AbortSignal) { g.AbortSignal = function () {}; }
  if (!g.AbortController) {
    g.AbortController = function () { this.signal = new g.AbortSignal(); };
    g.AbortController.prototype.abort = function () {};
  }
  // Bare V8 has no TextEncoder/TextDecoder; fontkit/pdfkit decode font + PDF byte
  // strings through them. Minimal, correct UTF-8 implementations (constructor args
  // and stream options are ignored — none of the consumers use them).
  if (!g.TextEncoder) {
    g.TextEncoder = function () {};
    g.TextEncoder.prototype.encode = function (s) {
      s = String(s);
      var out = [];
      for (var i = 0; i < s.length; i++) {
        var c = s.charCodeAt(i);
        if (c < 0x80) { out.push(c); }
        else if (c < 0x800) { out.push(0xC0 | (c >> 6), 0x80 | (c & 0x3F)); }
        else if (c >= 0xD800 && c <= 0xDBFF) {
          var c2 = s.charCodeAt(++i);
          var cp = 0x10000 + ((c & 0x3FF) << 10) + (c2 & 0x3FF);
          out.push(0xF0 | (cp >> 18), 0x80 | ((cp >> 12) & 0x3F), 0x80 | ((cp >> 6) & 0x3F), 0x80 | (cp & 0x3F));
        } else { out.push(0xE0 | (c >> 12), 0x80 | ((c >> 6) & 0x3F), 0x80 | (c & 0x3F)); }
      }
      return new Uint8Array(out);
    };
  }
  if (!g.TextDecoder) {
    // Honors the encoding label — fontkit decodes the font 'name' table in UTF-16BE
    // (Windows platform) and Mac-Roman/Latin1 (Mac platform); decoding those as
    // UTF-8 corrupts PostScript names (null-byte interleaving) and the embedded PDF
    // BaseFont entries. Supports utf-8, utf-16be/le, and latin1/ascii/mac-roman.
    g.TextDecoder = function (label) { this._enc = label ? String(label).toLowerCase() : "utf-8"; };
    g.TextDecoder.prototype.decode = function (buf) {
      if (!buf) return "";
      // Preserve the view's exact range (byteOffset/byteLength) for DataView and other
      // typed-array views; falling back to the whole ArrayBuffer would pull stray bytes.
      var b = buf instanceof Uint8Array
        ? buf
        : ArrayBuffer.isView(buf)
          ? new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength)
          : new Uint8Array(buf);
      var enc = this._enc, out = "", i, len = b.length;
      if (enc === "utf-16le" || enc === "utf-16" || enc === "unicode" || enc === "ucs-2") {
        for (i = 0; i + 1 < len; i += 2) out += String.fromCharCode(b[i] | (b[i + 1] << 8));
        return out;
      }
      if (enc === "utf-16be") {
        for (i = 0; i + 1 < len; i += 2) out += String.fromCharCode((b[i] << 8) | b[i + 1]);
        return out;
      }
      if (enc === "ascii" || enc === "latin1" || enc === "iso-8859-1" || enc === "windows-1252" ||
          enc === "x-user-defined" || enc === "macintosh" || enc === "mac-roman" || enc === "macroman") {
        // Single-byte decode. NOTE: windows-1252 (0x80-0x9F) and mac-roman (0x80-0xFF)
        // are APPROXIMATED as Latin-1 — exact only for bytes <= 0x7F. That's sufficient
        // for the bundled fonts (their PostScript names are ASCII or UTF-16BE; the
        // Node-vs-V8 byte-identity check confirms these labels are never hit with
        // high bytes here). Add real mapping tables if non-Latin1 fonts are bundled.
        for (i = 0; i < len; i++) out += String.fromCharCode(b[i]);
        return out;
      }
      i = 0; // utf-8 (default)
      while (i < len) {
        var c = b[i++];
        if (c < 0x80) { out += String.fromCharCode(c); }
        else if (c < 0xE0) { out += String.fromCharCode(((c & 0x1F) << 6) | (b[i++] & 0x3F)); }
        else if (c < 0xF0) { out += String.fromCharCode(((c & 0x0F) << 12) | ((b[i++] & 0x3F) << 6) | (b[i++] & 0x3F)); }
        else {
          var cp = ((c & 0x07) << 18) | ((b[i++] & 0x3F) << 12) | ((b[i++] & 0x3F) << 6) | (b[i++] & 0x3F);
          cp -= 0x10000;
          out += String.fromCharCode(0xD800 + (cp >> 10), 0xDC00 + (cp & 0x3FF));
        }
      }
      return out;
    };
  }
})(typeof globalThis !== "undefined" ? globalThis : this);`;

// node:fs is never used at runtime (fonts are host-injected; PDFDocument runs with
// font:false), but importing the jspm fs polyfill drags in browser-only globals.
// Stub it to an empty module so the dead `import("node:fs/promises")` resolves cheap.
const stubNodeFs = {
  name: "stub-node-fs",
  setup(b) {
    b.onResolve({ filter: /^node:fs(\/promises)?$/ }, (a) => ({ path: a.path, namespace: "stub-node-fs" }));
    b.onLoad({ filter: /.*/, namespace: "stub-node-fs" }, () => ({
      contents: "export default {}; export const readFile = undefined; export const promises = {};",
      loader: "js",
    }));
  },
};

await build({
  entryPoints: [ENTRY],
  outfile: OUTFILE,
  bundle: true,
  format: "iife",
  globalName: "WordCanvasBundle",
  platform: "browser",
  target: "es2021",
  // @cw/shared is a raw-TS workspace package — inline it (same as build-node.mjs).
  // process/node:process resolve to the host-owned shim (banner's globalThis.process)
  // so the whole bundle shares one scheduler with a microtask nextTick.
  alias: {
    "@cw/shared": SHARED,
    process: here("../src/clearscript/process-shim.js"),
    "node:process": here("../src/clearscript/process-shim.js"),
  },
  define: {
    "process.env.NODE_ENV": '"production"',
    global: "globalThis",
  },
  plugins: [
    // First: intercept node:fs before the polyfill plugin can map it.
    stubNodeFs,
    polyfillNode({
      // pdfkit requires node:crypto at top level (its unused PDFSecurity); fs is
      // stubbed above (never called — fonts are host-injected, PDFDocument uses
      // font:false). process is NOT polyfilled here — it's the host-owned shim
      // (see alias) so the whole bundle shares one microtask-nextTick scheduler.
      polyfills: { crypto: true },
      globals: { buffer: true, process: false },
    }),
  ],
  banner: { js: BANNER },
  sourcemap: false,
  legalComments: "none",
  logLevel: "info",
  metafile: false,
});

console.log(`[build-clearscript] wrote ${OUTFILE}`);
