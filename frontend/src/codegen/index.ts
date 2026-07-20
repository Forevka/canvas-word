// @forevka/wordcanvas/codegen — the reverse builder.
//
// Take a `.docx` (or an in-memory Document) and get back editable TypeScript that
// calls the fluent DocumentBuilder API to reconstruct it — so a document authored
// visually can be converted into a code TEMPLATE a developer maintains and
// parameterizes. The inverse of `@forevka/wordcanvas/builder`.
//
//   const { code, uncovered } = docxToBuilderCode(bytes);
//   // `code` is a ready-to-save .ts module exporting buildDocument(): Document
//   // `uncovered` lists any model fields the fluent API can't yet express.

import type { Document } from "@cw/shared";
import { runImport } from "../import/docx/pipeline";
import { emitBuilderCode, type CodegenOptions, type CodegenResult, type UncoveredField } from "./emit";

export { emitBuilderCode };
export type { CodegenOptions, CodegenResult, UncoveredField };

export interface DocxCodegenResult extends CodegenResult {
  /** Import diagnostics surfaced while parsing the .docx (codes only). */
  importWarnings: string[];
}

/** Import `docx` bytes and emit DocumentBuilder code that reconstructs it. */
export function docxToBuilderCode(docx: ArrayBuffer | Uint8Array, opts: CodegenOptions = {}): DocxCodegenResult {
  const bytes = docx instanceof Uint8Array ? docx : new Uint8Array(docx);
  const res = runImport(bytes, undefined, { collectMediaBytes: true });
  const out = emitBuilderCode(res.doc as Document, opts);
  return { ...out, importWarnings: res.warnings.map((w) => w.code) };
}
