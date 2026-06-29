// Guards the C# ClearScript bindings against drifting from the JS builder. The C#
// wrapper is a *stringly-typed* bridge — every method is `InvokeMethod("name", …)`
// against the JS builder — so the compiler can't catch a JS method that gained no
// C# mirror (the exact gap this test was born from) or a C# call to a method the JS
// builder renamed/removed. We reconcile both sides here:
//
//   • JS surface  — reflected off the builder class prototypes (+ statics).
//   • C# surface  — the JS method names the bindings actually invoke, scraped from
//                   the binding source (InvokeMethod / Sdt / SdtList literals).
//
// Two assertions: (A) no C# call targets a missing JS method, and (B) every JS
// method is bridged or explicitly allowlisted. A new public JS method fails (B) by
// default — forcing a deliberate choice: bridge it, or document why it isn't.

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { DocumentBuilder, ParagraphBuilder, RowBuilder, StoryBuilder, TableBuilder } from "./index";

const here = (p: string): string => fileURLToPath(new URL(p, import.meta.url));

// ── JS side: method names across the builder prototype chains + statics ──────
type Ctor = { prototype: object };

const protoMethods = (ctor: Ctor): string[] => {
  const out: string[] = [];
  let proto: object | null = ctor.prototype;
  while (proto && proto !== Object.prototype) {
    for (const name of Object.getOwnPropertyNames(proto)) {
      if (name === "constructor") continue;
      const d = Object.getOwnPropertyDescriptor(proto, name);
      if (d && typeof d.value === "function") out.push(name); // methods only (getters excluded)
    }
    proto = Object.getPrototypeOf(proto);
  }
  return out;
};

const staticMethods = (ctor: object): string[] => {
  const skip = new Set(["length", "name", "prototype"]);
  return Object.getOwnPropertyNames(ctor).filter((n) => {
    if (skip.has(n)) return false;
    const d = Object.getOwnPropertyDescriptor(ctor, n);
    return !!d && typeof d.value === "function";
  });
};

const jsMethods = new Set<string>([
  ...protoMethods(DocumentBuilder as unknown as Ctor),
  ...protoMethods(StoryBuilder as unknown as Ctor),
  ...protoMethods(ParagraphBuilder as unknown as Ctor),
  ...protoMethods(TableBuilder as unknown as Ctor),
  ...protoMethods(RowBuilder as unknown as Ctor),
  ...staticMethods(DocumentBuilder),
]);

// ── C# side: the JS method names the bindings invoke ─────────────────────────
// Only the binding facades call builder methods; Specs.cs builds option objects
// (Js.Set keys) and array pushes, which are NOT builder methods — so we don't scan it.
const CS_FILES = ["StoryBuilder.cs", "DocumentBuilder.cs"];
const csDir = here("../../../dotnet/src/WordCanvas.ClearScript/Builder/");
const csSource = CS_FILES.map((f) => {
  try {
    return readFileSync(`${csDir}${f}`, "utf8");
  } catch {
    throw new Error(`C# binding source not found: ${csDir}${f} — the parity test needs the .NET bindings in the repo.`);
  }
}).join("\n");

// Method bridges: `…InvokeMethod("name"` plus the two helpers that forward a literal
// method name (`Sdt("name", …)` / `SdtList("name", …)`).
const bridged = new Set<string>(
  [...csSource.matchAll(/(?:InvokeMethod|\bSdt|\bSdtList)\(\s*"([a-zA-Z][A-Za-z0-9]*)"/g)].map((m) => m[1]!),
);

// ── Allowlists: JS methods intentionally not bridged 1:1 ─────────────────────
// Internal helpers (TS `private`/`protected` erased onto the prototype) + the
// JS-only paragraph-scope escape (`end()` — C# uses a configure lambda instead).
const JS_INTERNAL = new Set([
  "push", "flushPendingPageBreak", "band", "compileSectionPatch",
  "pushRun", "applyChar", "clearChar", "markRunCharKeys", "emitField", "emitSdt",
  "toBlock", "applyTableStyleRef", "applyTableDefaults", "cascadeBordersOntoCells", "end",
]);
// Generic JS escape hatches the C# surface exposes through TYPED variants instead
// of one stringly-typed method (a documented divergence, not drift):
//   field(spec)          → PageField()/DateField()/TimeField()/IfField()/CustomField()
//   contentControl(props)→ RichTextControl()/PlainTextControl()/Checkbox()/DropDown()/…
const JS_TYPED_VARIANTS = new Set(["field", "contentControl"]);

describe("C# binding parity with the JS builder", () => {
  it("every JS builder method is bridged in C# (or explicitly allowlisted)", () => {
    const missing = [...jsMethods]
      .filter((m) => !bridged.has(m) && !JS_INTERNAL.has(m) && !JS_TYPED_VARIANTS.has(m))
      .sort();
    // To fix a failure: add the method to the C# bindings (dotnet/src/WordCanvas.ClearScript/Builder),
    // or — if it is intentionally JS-only — add it to JS_INTERNAL / JS_TYPED_VARIANTS with a reason.
    expect(missing, `JS builder methods with no C# binding: ${missing.join(", ")}`).toEqual([]);
  });

  it("every C# InvokeMethod target exists on the JS builder (no renamed/removed calls)", () => {
    const dangling = [...bridged].filter((m) => !jsMethods.has(m)).sort();
    expect(dangling, `C# invokes JS builder methods that no longer exist: ${dangling.join(", ")}`).toEqual([]);
  });

  it("reflects a non-trivial surface on both sides (guards against an empty scrape)", () => {
    expect(jsMethods.size).toBeGreaterThan(30);
    expect(bridged.size).toBeGreaterThan(30);
  });
});
