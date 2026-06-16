// Public types for @forevka/wordcanvas/generate-toc — generate a Word TOC field's
// result entries (styled, with live PAGEREF page numbers + hyperlinks) and splice
// them into the field's result region in the ORIGINAL document.xml, re-zipping.
// Drift-free (no whole-document re-export). Hand-written + self-contained. Await
// installMeasureHost() (from "@forevka/wordcanvas/export/measure") first.

import type { CharStyle, ParaStyle } from "./model";

/** Per-level (1-based) TOC entry style override. */
export interface TocLevelStyle {
  char?: Partial<CharStyle>;
  para?: Partial<ParaStyle>;
}

export interface TocOptions {
  /** TOC title paragraph; null/omit to skip a title. */
  title?: { text?: string; char?: Partial<CharStyle>; para?: Partial<ParaStyle>; namedStyle?: string } | null;
  /** Base char style applied to every entry before per-level overrides. */
  baseChar?: Partial<CharStyle>;
  /** Per-level (1-based) style overrides. */
  levels?: Record<number, TocLevelStyle>;
  /** Deepest heading level to include (default 3). */
  maxLevel?: number;
  /** Per-level left indent step in px (default 20). */
  indentStepPx?: number;
  /** Dot-leader between label and page number (default "dot"). */
  leader?: "dot" | "dash" | "underscore" | "none";
  /** Show page numbers (default true). */
  includePageNumbers?: boolean;
  /** Emit entries as hyperlinks (default true). */
  hyperlink?: boolean;
}

export interface GenerateTocResult {
  bytes: Uint8Array;
  /** Entries written into the field result. */
  generated: number;
  /** Headings detected in the document. */
  headings: number;
  /** Heading bookmarks synthesized + injected (those that had none). */
  bookmarksSynthesized: number;
}

/** Generate the first TOC field's result in `docxBytes`; returns the patched .docx. */
export declare function generateTocInDocx(docxBytes: Uint8Array, opts?: TocOptions): GenerateTocResult;
