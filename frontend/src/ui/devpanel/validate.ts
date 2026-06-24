// Document validator — a pure model walk that surfaces integrity problems the
// importer/round-trip can introduce: dangling content-control / field / style /
// list / bookmark / TOC / footnote references, duplicate block ids, and unused
// content controls / fields. No DOM; unit-tested in validate.test.ts.

import type { Block, Document, Run } from "@cw/shared";
import { BAND_CONTAINERS } from "@cw/shared";

export type Severity = "error" | "warning" | "info";

export interface Diagnostic {
  severity: Severity;
  code: string;
  message: string;
  /** Block to reveal on click, when the problem is anchored to one. */
  blockId?: string | undefined;
}

export function validateDocument(doc: Document): Diagnostic[] {
  const out: Diagnostic[] = [];
  const blockIds = new Set<string>();
  const dupIds = new Set<string>();
  const usedSdt = new Set<string>();
  const usedField = new Set<string>();
  const tocTargets: { targetId: string; blockId: string }[] = [];
  const styleIds = new Set((doc.stylesheet?.styles ?? []).map((s) => s.id));

  const visitRun = (r: Run, blockId: string): void => {
    for (const id of r.style.sdtPath ?? []) usedSdt.add(id);
    if (r.style.fieldId) usedField.add(r.style.fieldId);
    if (r.style.charStyleId && !styleIds.has(r.style.charStyleId)) {
      out.push({ severity: "warning", code: "charstyle-missing", message: `Run references missing character style "${r.style.charStyleId}"`, blockId });
    }
    if (r.style.footnoteRef && !doc.footnotes?.[r.style.footnoteRef]) {
      out.push({ severity: "warning", code: "footnote-missing", message: `Run references missing footnote "${r.style.footnoteRef}"`, blockId });
    }
  };

  const visitBlock = (b: Block): void => {
    if (blockIds.has(b.id)) dupIds.add(b.id);
    else blockIds.add(b.id);
    for (const id of b.sdtPath ?? []) usedSdt.add(id);
    if (b.fieldId) usedField.add(b.fieldId);
    if (b.kind === "paragraph") {
      if (b.style.list) {
        if (!doc.lists?.[b.style.list.listId]) out.push({ severity: "error", code: "list-missing", message: `Paragraph references missing list "${b.style.list.listId}"`, blockId: b.id });
      }
      if (b.style.namedStyle && !styleIds.has(b.style.namedStyle)) {
        out.push({ severity: "warning", code: "style-missing", message: `Paragraph references missing style "${b.style.namedStyle}"`, blockId: b.id });
      }
      if (b.style.tocEntry) tocTargets.push({ targetId: b.style.tocEntry.targetId, blockId: b.id });
      for (const r of b.runs) visitRun(r, b.id);
    } else if (b.kind === "table") {
      if (b.styleId && !doc.tableStyles?.[b.styleId]) {
        out.push({ severity: "warning", code: "tablestyle-missing", message: `Table references missing table style "${b.styleId}"`, blockId: b.id });
      }
      for (const row of b.rows) for (const cell of row.cells) for (const cb of cell.blocks) visitBlock(cb);
    }
  };

  for (const b of doc.blocks) visitBlock(b);
  for (const band of BAND_CONTAINERS) for (const b of doc.section[band] ?? []) visitBlock(b);
  for (const paras of Object.values(doc.footnotes ?? {})) for (const p of paras) visitBlock(p);

  // Duplicate ids (break the layout cache + undo, which key on block id).
  for (const id of dupIds) out.push({ severity: "error", code: "duplicate-id", message: `Duplicate block id "${id}" appears more than once`, blockId: id });

  // Dangling + unused content controls.
  for (const id of usedSdt) if (!doc.sdts?.[id]) out.push({ severity: "error", code: "sdt-missing", message: `Content control "${id}" is referenced but not defined in doc.sdts` });
  for (const id of Object.keys(doc.sdts ?? {})) if (!usedSdt.has(id)) out.push({ severity: "info", code: "sdt-unused", message: `Content control "${id}" is defined but never referenced` });

  // Dangling + unused fields.
  for (const id of usedField) if (!doc.fields?.[id]) out.push({ severity: "error", code: "field-missing", message: `Field "${id}" is referenced but not defined in doc.fields` });
  for (const id of Object.keys(doc.fields ?? {})) if (!usedField.has(id)) out.push({ severity: "info", code: "field-unused", message: `Field "${id}" is defined but never referenced` });

  // Bookmarks must anchor to real blocks.
  for (const [name, range] of Object.entries(doc.bookmarks ?? {})) {
    if (!blockIds.has(range.start.blockId)) out.push({ severity: "error", code: "bookmark-start-missing", message: `Bookmark "${name}" starts in a missing block`, blockId: range.start.blockId });
    if (!blockIds.has(range.end.blockId)) out.push({ severity: "error", code: "bookmark-end-missing", message: `Bookmark "${name}" ends in a missing block`, blockId: range.end.blockId });
  }

  // TOC entry targets + the TOC anchor block.
  for (const t of tocTargets) if (!blockIds.has(t.targetId)) out.push({ severity: "warning", code: "toc-target-missing", message: `TOC entry points at missing heading "${t.targetId}"`, blockId: t.blockId });
  if (doc.tocAnchorBlockId && !blockIds.has(doc.tocAnchorBlockId)) out.push({ severity: "warning", code: "toc-anchor-missing", message: `tocAnchorBlockId "${doc.tocAnchorBlockId}" is not a block in the document` });

  return out;
}

/** Count by severity for the tab badge. */
export function countBySeverity(diags: Diagnostic[]): Record<Severity, number> {
  const c: Record<Severity, number> = { error: 0, warning: 0, info: 0 };
  for (const d of diags) c[d.severity]++;
  return c;
}
