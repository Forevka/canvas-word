import { fileURLToPath } from "node:url";
import { defineConfig } from "vitest/config";

const at = (p: string): string => fileURLToPath(new URL(p, import.meta.url));

// Vitest (Vite) does NOT read tsconfig `paths`, so mirror backend/tsconfig.json's
// paths here. The backend reuses the frontend's pure, DOM-free import/export
// pipelines and the shared document model straight from source. (`tsx` honors
// tsconfig paths at runtime and `tsc` honors them for typecheck; only the test
// runner needs this.) These subpaths are intentionally NOT in the published
// @forevka/wordcanvas `exports` — they are a workspace-internal convenience.
//
// Order matters: Vite alias entries match in order and a bare string match is a
// prefix match, so the more specific `/export/measure` MUST precede `/export`.
export default defineConfig({
  resolve: {
    alias: [
      { find: "@cw/shared", replacement: at("../shared/src/index.ts") },
      { find: "@forevka/wordcanvas/import", replacement: at("../frontend/src/import/docx/pipeline.ts") },
      { find: "@forevka/wordcanvas/export/measure", replacement: at("../frontend/src/export/shared/measureHost.ts") },
      { find: "@forevka/wordcanvas/export", replacement: at("../frontend/src/export/pipeline.ts") },
    ],
  },
});
