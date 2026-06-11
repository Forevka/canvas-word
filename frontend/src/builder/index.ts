// @forevka/wordcanvas/builder — fluent programmatic document composition.
//
//   const b = await DocumentBuilder.fromTemplate(templateBytes);
//   const doc = b
//     .paragraph(data.title).withStyle("Heading1")
//     .paragraph("Generated " + data.date)
//     .table([["Item", "Qty"], ...data.rows], { headerRow: true })
//     .build();
//
// Pure data in, pure data out: build() returns the same Document model the
// editor renders and the exporter writes — usable in the browser (live preview
// via WordCanvas.setDocument) and server-side in Node (DOCX/PDF generation).

export type { Block, CharStyle, Document, NamedStyle, ParaStyle, Stylesheet } from "@cw/shared";
export { DocumentBuilder } from "./documentBuilder";
export type { BandOptions, CreateOptions, PageSetup, TemplateOptions } from "./documentBuilder";
export { ParagraphBuilder } from "./paragraphBuilder";
export type { IndentOptions, SpacingOptions } from "./paragraphBuilder";
export { StoryBuilder } from "./storyBuilder";
export type { ImageOptions, ListItem, ListOptions } from "./storyBuilder";
export { RowBuilder, TableBuilder } from "./tableBuilder";
export type { CellContent, CellOptions, CellSpec, TableOptions } from "./tableBuilder";
export type { BuilderWarning } from "./blockFactory";
export { bytesToDataUrl } from "./media";
export { cm, inches, PAGE_SIZES, pt, twips } from "./units";
export type { PageSize, PageSizeName } from "./units";
