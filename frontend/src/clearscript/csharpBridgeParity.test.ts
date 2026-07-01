// C# ↔ JS bridge parity guard. The C# bindings drive the JS runtime by NAME —
// `WordDocumentEditor` calls `_editor.InvokeMethod("setSdtText", …)` and
// `WordDocumentQuery` calls `_engine.Api.InvokeMethod("querySdts", …)`. If a JS
// method is renamed/removed, the C# call fails only at RUNTIME (the showcase
// catches it, but CI does not build .NET). This test scrapes those names from the
// C# source and asserts each targets a real JS target, so drift breaks `npm test`.

import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { DocumentEditor } from "@cw/shared";
import * as bridge from "./queryBridge";

const HERE = dirname(fileURLToPath(import.meta.url));
const readSrc = (relative: string): string => readFileSync(resolve(HERE, relative), "utf8");
const CS = (name: string): string => readSrc(`../../../dotnet/src/WordCanvas.ClearScript/${name}`);

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

    const proto = DocumentEditor.prototype as unknown as Record<string, unknown>;
    const isMethod = (n: string): boolean => typeof proto[n] === "function";
    const isGetter = (n: string): boolean => Object.getOwnPropertyDescriptor(proto, n)?.get !== undefined;

    expect(methods.length).toBeGreaterThan(0);
    for (const name of methods) expect(isMethod(name), `DocumentEditor.${name}() missing`).toBe(true);
    for (const name of props) expect(isGetter(name), `DocumentEditor.${name} getter missing`).toBe(true);
  });

  it("every `_engine.Api` call (query + editor) targets a wired JS bridge fn", () => {
    // Query methods and a few editor methods (e.g. newRegExp) drive the entry `api`.
    const re = /_engine\.Api\.InvokeMethod\("(\w+)"/g;
    const names = [...new Set([...namesFrom(CS("WordDocumentQuery.cs"), re), ...namesFrom(CS("WordDocumentEditor.cs"), re)])];
    // Bridge mappers are exported from queryBridge.ts; a few helpers live directly
    // on the entry `api`. Every name must have a JS impl AND be wired onto the api.
    const bridgeExports = new Set(Object.keys(bridge));
    const entryLocal = new Set(["layoutPages", "newRegExp"]);
    const entrySrc = readSrc("./entry.ts");

    expect(names.length).toBeGreaterThan(0);
    for (const name of names) {
      const hasImpl = bridgeExports.has(name) || entryLocal.has(name);
      expect(hasImpl, `no JS bridge fn named "${name}" (queryBridge export or an entry-local helper)`).toBe(true);
      // Word-boundary match so a substring hit inside a comment/other identifier
      // can't count as "wired" — the name must appear as its own token.
      const wired = new RegExp(`\\b${name}\\b`).test(entrySrc);
      expect(wired, `"${name}" is not wired onto the entry api`).toBe(true);
    }
  });
});
