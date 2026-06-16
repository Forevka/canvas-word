import { defineConfig } from "vite";

// Admin dashboard. Dev server on 5174 (alongside the editor on 5173). The
// production build is served under /dashboard/ on the app's single origin, so
// assets must be base-pathed there; dev stays at root. Backend/editor origins
// come from VITE_BACKEND_URL / VITE_EDITOR_URL (see src/api.ts).
export default defineConfig(({ command }) => ({
  base: command === "build" ? "/dashboard/" : "/",
  // strictPort: fail loudly if 5174 is taken rather than silently drifting to a
  // random port — dev:online advertises "dashboard → :5174", so it must stay put.
  server: { port: 5174, strictPort: true },
}));
