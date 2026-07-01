// Public entry for @forevka/wordcanvas/query — the document query + edit API,
// the rough analog of .NET's WordprocessingDocument access. Read a document
// (traverse, find by text, enumerate sections), edit it in place via
// DocumentEditor (over the operation engine, with undo/redo), and ask which page
// content lands on. Isomorphic model core (query/edit) + a layout-backed page
// query. Pairs with the ./import and ./export subpaths for a full round-trip.
//
//   import { findParagraphs, DocumentEditor, getPages } from "@forevka/wordcanvas/query";
//   import { importDocx } from "@forevka/wordcanvas/import";
//   const { doc } = await importDocx(bytes);
//   const ed = new DocumentEditor(doc);
//   for (const m of findParagraphs(ed.doc, "DRAFT")) ed.setParagraphText(m.paragraph.id, "FINAL");

export {
  // traversal / query
  walk,
  textOf,
  getParagraphs,
  getTables,
  getImages,
  findParagraphs,
  getBlockById,
  getParagraphById,
  getTableById,
  getImageById,
  getSections,
  resolveSections,
  // content controls (SDTs) — the primary templating surface
  getSdt,
  getSdts,
  getSdtsByTag,
  getSdtsByAlias,
  getSdtNodes,
  getSdtRoots,
  getSdtChildren,
  getSdtAncestors,
  getSdtDescendants,
  getSdtBlocks,
  sdtText,
  getSdtValue,
  // edit facade
  DocumentEditor,
} from "@cw/shared";

export type {
  BlockContext,
  BlockVisitor,
  WalkOptions,
  ParagraphMatch,
  ResolvedSection,
  InsertParagraphOptions,
  SdtMatch,
  SdtNode,
  SdtValue,
} from "@cw/shared";

// Re-export the core model types consumers need to name when using the API.
export type {
  Block,
  CharStyle,
  Document,
  DocPosition,
  EquationBlock,
  ImageBlock,
  Paragraph,
  ParaStyle,
  Run,
  SdtProps,
  SdtType,
  SectionProps,
  TableBlock,
} from "@cw/shared";

// Layout-backed page query ("what's on page N").
export { getPages, pageOfBlock } from "./layout/pages";
export type { PageInfo, GetPagesOptions } from "./layout/pages";
