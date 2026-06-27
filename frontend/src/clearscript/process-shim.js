// Re-exports the host-controlled `process` (defined in the esbuild banner so it
// exists before any module initializes). Aliased for `process` and `node:process`
// so pdfkit / readable-stream / browserify-zlib all share ONE process whose
// nextTick is a real microtask the host pump drains — see build-clearscript.mjs.
const p = globalThis.process;
export default p;
export const nextTick = p.nextTick;
export const env = p.env;
export const platform = p.platform;
export const version = p.version;
export const versions = p.versions;
export const argv = p.argv;
export const browser = true;
export const cwd = p.cwd;
export const emitWarning = p.emitWarning;
export const on = p.on;
export const once = p.once;
export const off = p.off;
export const removeListener = p.removeListener;
export const binding = p.binding;
