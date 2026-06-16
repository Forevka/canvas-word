import { fileURLToPath } from "node:url";
import { defineConfig } from "vitest/config";

const at = (p: string): string => fileURLToPath(new URL(p, import.meta.url));

// @cw/shared is consumed straight from source (Vite doesn't read tsconfig paths,
// so mirror that one alias here). The @forevka/wordcanvas pipelines are NOT aliased
// to source any more — the backend exercises the published package exactly as an
// external consumer does, resolving them through its `exports` map. `conditions`
// pins the `node` condition so we get the Node bundle (dist-node, real fs + bundled
// pdfkit), not the browser `default` build; `deps.external` keeps that compiled
// ESM bundle on Node's own loader instead of routing it through Vite's transform.
// (`build:node` rebuilds dist-node first — see the pretest script.)
export default defineConfig({
  resolve: {
    alias: [{ find: "@cw/shared", replacement: at("../shared/src/index.ts") }],
    conditions: ["node", "import", "default"],
  },
  test: {
    environment: "node",
    server: { deps: { external: [/@forevka\/wordcanvas/] } },
  },
});
