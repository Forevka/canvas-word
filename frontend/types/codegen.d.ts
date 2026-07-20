// Public type surface for @forevka/wordcanvas/codegen — the REVERSE builder.
// Hand-written, self-contained (model types come from ./model, shared with
// wordcanvas.d.ts / builder.d.ts).
//
// Take a `.docx` (or an in-memory Document) and get back editable TypeScript that
// calls the fluent DocumentBuilder API to reconstruct it — turning a visually
// authored document into a code template. The inverse of ./builder.
//
//   import { docxToBuilderCode } from "@forevka/wordcanvas/codegen";
//   const { code, uncovered } = docxToBuilderCode(bytes);
//   // `code` is a ready-to-save .ts module exporting buildDocument(): Document
//   // `uncovered` lists any model fields the fluent API cannot yet express.

import type { Document } from "./model";

export type { Document } from "./model";

/** A model field the emitted code could not reproduce through the fluent API,
 *  recorded (never silently dropped) so it becomes a backlog item for growing the
 *  builder. */
export interface UncoveredField {
  /** Human-readable location, e.g. `blocks[3].runs[1]` or `section`. */
  path: string;
  /** The model field name. */
  field: string;
  /** A sample of the value (best-effort; large values are summarized). */
  value?: unknown;
  /** Why it could not be expressed / what would close the gap. */
  note?: string;
}

export interface CodegenOptions {
  /** Name of the generated function (default "buildDocument"). */
  functionName?: string;
  /** Import specifier for DocumentBuilder (default "@forevka/wordcanvas/builder"). */
  importFrom?: string;
}

export interface CodegenResult {
  /** The generated TypeScript module source. */
  code: string;
  /** Fields that could not be reproduced through the fluent API. */
  uncovered: UncoveredField[];
}

export interface DocxCodegenResult extends CodegenResult {
  /** Import diagnostics surfaced while parsing the .docx (codes only). */
  importWarnings: string[];
}

/** Generate DocumentBuilder TypeScript that reconstructs `doc`. */
export declare function emitBuilderCode(doc: Document, opts?: CodegenOptions): CodegenResult;

/** Import `docx` bytes and emit DocumentBuilder code that reconstructs it. */
export declare function docxToBuilderCode(docx: ArrayBuffer | Uint8Array, opts?: CodegenOptions): DocxCodegenResult;
