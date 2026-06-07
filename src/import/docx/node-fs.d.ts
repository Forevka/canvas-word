// Tests run under Node (vitest), but the project compiles with DOM-only libs.
// Declare the one Node API the tests use instead of adding @types/node globally
// (which would leak Node globals into app code and mask DOM/Node mix-ups).
declare module "node:fs" {
  export function readFileSync(path: URL | string): Uint8Array;
}
