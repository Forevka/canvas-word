// C# ↔ JS bridge parity guard. The C# bindings drive the JS runtime by NAME —
// `WordDocumentEditor` calls `_editor.InvokeMethod("setSdtText", …)` and
// `WordDocumentQuery` calls `_engine.Api.InvokeMethod("querySdts", …)`. If a JS
// method is renamed/removed, the C# call fails only at RUNTIME (the showcase
// catches it, but CI does not build .NET). This test scrapes those names from the
// C# source and asserts each targets a real JS target, so drift breaks `npm test`.

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { DocumentEditor } from "@cw/shared";
import * as bridge from "./queryBridge";

const HERE = fileURLToPath(new URL(".", import.meta.url));
const CS = (name: string): string =>
  readFileSync(fileURLToPath(new URL(`../../../dotnet/src/WordCanvas.ClearScript/${name}`, `file://${HERE}`)), "utf8");

const namesFrom = (src: string, re: RegExp): string[] => {
  const out = new Set<string>();
  for (const m of src.matchAll(re)) out.add(m[1]!);
  return [...out];
};

describe("C#↔JS bridge parity", () => {
  it("every WordDocumentEditor _editor call targets a real DocumentEditor member", () => {
    const src = CS("WordDocumentEditor.cs");
    const methods = namesFrom(src, /_editor\.InvokeMethod\("(\w+)"/g);
    const props = namesFrom(src, /_editor\.GetProperty\("(\w+)"/g);

    const proto = DocumentEditor.prototype;
    const isMethod = (n: string): boolean => typeof (proto as Record<string, unknown>)[n] === "function";
    const isGetter = (n: string): boolean => Object.getOwnPropertyDescriptor(proto, n)?.get !== undefined;

    expect(methods.length).toBeGreaterThan(0);
    for (const name of methods) expect(isMethod(name), `DocumentEditor.${name}() missing`).toBe(true);
    for (const name of props) expect(isGetter(name), `DocumentEditor.${name} getter missing`).toBe(true);
  });

  it("every WordDocumentQuery Api call targets a wired JS bridge fn", () => {
    const src = CS("WordDocumentQuery.cs");
    const names = namesFrom(src, /_engine\.Api\.InvokeMethod\("(\w+)"/g);
    // Bridge mappers are exported from queryBridge.ts; `layoutPages` is composed in
    // entry.ts (async page layout). Both must also be wired onto the entry `api`.
    const bridgeExports = new Set(Object.keys(bridge));
    const entrySrc = readFileSync(fileURLToPath(new URL("./entry.ts", import.meta.url)), "utf8");

    expect(names.length).toBeGreaterThan(0);
    for (const name of names) {
      const hasImpl = bridgeExports.has(name) || name === "layoutPages";
      expect(hasImpl, `no JS bridge fn named "${name}" (queryBridge export or layoutPages)`).toBe(true);
      expect(entrySrc.includes(name), `"${name}" is not wired onto the entry api`).toBe(true);
    }
  });
});
