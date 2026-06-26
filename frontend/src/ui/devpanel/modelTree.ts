// Model tree: turns the parsed `Document` into navigable tree nodes — body blocks →
// runs, tables → rows → cells, header/footer bands, footnotes, and the side-tables.
// Content-control / field membership (flat sdtPath / fieldId markers) is reconstructed
// into nested "SDT → … → runs" and "Field → run" group nodes.

import type { Block, Document, ImageBlock, Paragraph, Run, TableBlock } from "@cw/shared";
import type { MathNode } from "@cw/shared";
import { BAND_CONTAINERS, mathSlots, mathToPlainText, textOfRuns } from "@cw/shared";
import type { TreeNode } from "./types";
import { previewText, shortId } from "./types";

const runBadges = (style: Run["style"]): string[] => {
  const b: string[] = [];
  if (style.bold) b.push("B");
  if (style.italic) b.push("I");
  if (style.underline) b.push("U");
  if (style.strikethrough) b.push("S");
  if (style.hidden) b.push("hidden");
  if (style.verticalAlign) b.push(style.verticalAlign);
  if (style.link) b.push("link");
  if (style.footnoteRef) b.push("footnote");
  if (style.charStyleId) b.push(`cs:${style.charStyleId}`);
  if (style.sdtPath?.length) b.push(`sdt:${style.sdtPath.length}`);
  if (style.fieldId) b.push("field");
  return b;
};

const paraBadges = (p: Paragraph): string[] => {
  const b: string[] = [];
  if (p.style.namedStyle) b.push(p.style.namedStyle);
  if (p.style.list) b.push(`list L${p.style.list.level}`);
  if (p.style.outlineLevel !== undefined) b.push(`outline ${p.style.outlineLevel}`);
  if (p.sdtPath?.length) b.push(`sdt:${p.sdtPath.length}`);
  if (p.fieldId) b.push("field");
  return b;
};

const imageBadges = (im: ImageBlock): string[] => {
  const b = [`${Math.round(im.widthPx)}×${Math.round(im.heightPx)}`];
  if (im.wrap) b.push(im.wrap);
  if (im.anchor) b.push("anchored");
  if (im.sdtPath?.length) b.push(`sdt:${im.sdtPath.length}`);
  if (im.fieldId) b.push("field");
  return b;
};

const tableBadges = (t: TableBlock): string[] => {
  const cols = t.rows[0]?.cells.length ?? 0;
  const b = [`${t.rows.length}×${cols}`];
  if (t.styleId) b.push(`style:${t.styleId}`);
  if (t.sdtPath?.length) b.push(`sdt:${t.sdtPath.length}`);
  if (t.fieldId) b.push("field");
  return b;
};

const blockPath = (b: Block): string[] => [...(b.sdtPath ?? []), ...(b.fieldId ? [`field:${b.fieldId}`] : [])];
const runPath = (r: Run): string[] => [...(r.style.sdtPath ?? []), ...(r.style.fieldId ? [`field:${r.style.fieldId}`] : [])];

/** A group node for one ancestry id (an SDT id, or "field:<fieldId>"). */
function groupNode(doc: Document, id: string, key: string, children: TreeNode[]): TreeNode {
  if (id.startsWith("field:")) {
    const fid = id.slice(6);
    const def = doc.fields?.[fid];
    return {
      key, kind: "group",
      label: `❏ Field ${def?.name ?? shortId(fid)}`,
      preview: previewText(def?.instruction ?? ""),
      badges: def?.kind ? [def.kind] : [],
      target: { kind: "field", fieldId: fid },
      data: def ?? { fieldId: fid },
      children,
    };
  }
  const sdt = doc.sdts?.[id];
  return {
    key, kind: "group",
    label: `❑ SDT ${sdt?.alias ?? sdt?.tag ?? sdt?.type ?? shortId(id)}`,
    preview: "",
    badges: sdt?.type ? [sdt.type] : [],
    target: { kind: "sdt", sdtId: id },
    data: sdt ?? { sdtId: id },
    children,
  };
}

/** Wrap a sequence of items into nested SDT/Field group nodes by shared ancestry.
 *  Items whose path is exhausted at this depth become leaves via `makeLeaf`. */
function groupByPath<T>(
  doc: Document,
  items: T[],
  getPath: (t: T) => string[],
  keyOf: (t: T) => string,
  makeLeaf: (t: T) => TreeNode,
  depth: number,
): TreeNode[] {
  const out: TreeNode[] = [];
  let i = 0;
  while (i < items.length) {
    const path = getPath(items[i]!);
    if (path.length <= depth) { out.push(makeLeaf(items[i]!)); i++; continue; }
    const id = path[depth]!;
    let j = i;
    while (j < items.length) {
      const p = getPath(items[j]!);
      if (p.length > depth && p[depth] === id) j++;
      else break;
    }
    const slice = items.slice(i, j);
    out.push(groupNode(doc, id, `grp/${depth}/${id}/${keyOf(slice[0]!)}`, groupByPath(doc, slice, getPath, keyOf, makeLeaf, depth + 1)));
    i = j;
  }
  return out;
}

/** Build the tree node for one block, recursing into table cells, with inline runs
 *  grouped under their content-control / field membership. */
/** A MathML AST node as a tree node (recursive). Atoms show their text; every
 *  node highlights the equation it belongs to (math nodes have no own range). */
function mathNode(n: MathNode, key: string, target: TreeNode["target"]): TreeNode {
  const slots = mathSlots(n);
  const text = n.type === "ident" || n.type === "number" || n.type === "op" || n.type === "text" ? n.text : "";
  const badges: string[] = [];
  if (n.type === "ident" && n.variant) badges.push(n.variant);
  if (n.type === "op" && n.stretchy) badges.push("stretchy");
  if (n.type === "frac" && n.thickness === "0") badges.push("no-bar");
  if (n.type === "nary") badges.push(`op ${n.op}`);
  if (n.type === "fenced") badges.push(`${n.open || "·"} ${n.close || "·"}`);
  return {
    key,
    kind: "tag",
    label: text ? `${n.type} "${text}"` : n.type,
    preview: slots.length ? "" : mathToPlainText(n),
    badges,
    target,
    data: n,
    children: slots.map((c, i) => mathNode(c, `${key}.${i}`, target)),
  };
}

function blockNode(doc: Document, block: Block): TreeNode {
  if (block.kind === "paragraph") {
    type RunItem = { run: Run; index: number; start: number; end: number };
    const items: RunItem[] = [];
    let off = 0;
    block.runs.forEach((r, i) => { const start = off; off += r.text.length; items.push({ run: r, index: i, start, end: off }); });
    const runLeaf = (it: RunItem): TreeNode => {
      const target = { kind: "run" as const, blockId: block.id, start: it.start, end: it.end };
      const eq = it.run.style.equation;
      // An inline equation run (a single U+FFFC) shows as "∑ equation" and expands
      // into its MathML AST, instead of an opaque object-replacement char.
      if (eq) {
        return {
          key: `${block.id}#run${it.index}`,
          kind: "run",
          label: "∑ equation (inline)",
          preview: mathToPlainText(eq.root) || "∅",
          badges: ["inline"],
          target,
          data: it.run,
          children: eq.root.children.map((n, i) => mathNode(n, `${block.id}#run${it.index}.m${i}`, target)),
        };
      }
      return {
        key: `${block.id}#run${it.index}`,
        kind: "run",
        label: "run",
        preview: previewText(it.run.text) || "∅",
        badges: runBadges(it.run.style),
        target,
        data: it.run,
        children: [],
      };
    };
    return {
      key: block.id, kind: "block",
      label: "¶ paragraph",
      preview: previewText(textOfRuns(block.runs)) || "(empty)",
      badges: paraBadges(block),
      blockId: block.id,
      target: { kind: "block", blockId: block.id },
      data: block,
      children: groupByPath(doc, items, (it) => runPath(it.run), (it) => `${block.id}#run${it.index}`, runLeaf, 0),
    };
  }
  if (block.kind === "image") {
    return {
      key: block.id, kind: "block",
      label: "🖼 image",
      preview: block.mediaId ? `media ${shortId(block.mediaId)}` : "(inline)",
      badges: imageBadges(block),
      blockId: block.id,
      target: { kind: "block", blockId: block.id },
      data: block,
      children: [],
    };
  }
  if (block.kind === "equation") {
    return {
      key: block.id, kind: "block",
      label: "∑ equation",
      preview: mathToPlainText(block.equation.root),
      badges: [block.equation.display ? "display" : "inline"],
      blockId: block.id,
      target: { kind: "block", blockId: block.id },
      data: block,
      // Expand into the MathML AST so the structure (fractions, scripts, …) is
      // inspectable; every math node highlights the whole equation.
      children: block.equation.root.children.map((n, i) => mathNode(n, `${block.id}#m${i}`, { kind: "block", blockId: block.id })),
    };
  }
  // table → rows → cells → (grouped) blocks
  const rows: TreeNode[] = block.rows.map((row, ri) => ({
    key: `${block.id}#r${ri}`, kind: "tag",
    label: `row ${ri}`,
    preview: "",
    badges: [`${row.cells.length} cells`],
    target: { kind: "row", tableId: block.id, ri },
    data: row,
    children: row.cells.map((cell, ci) => ({
      key: cell.id, kind: "tag",
      label: `cell ${ri},${ci}`,
      preview: previewText(cell.blocks.map((b) => (b.kind === "paragraph" ? textOfRuns(b.runs) : "")).join(" ")),
      badges: [
        ...(cell.colSpan && cell.colSpan > 1 ? [`colSpan ${cell.colSpan}`] : []),
        ...(cell.rowSpan && cell.rowSpan > 1 ? [`rowSpan ${cell.rowSpan}`] : []),
        ...(cell.shading ? ["shaded"] : []),
      ],
      target: { kind: "cell", tableId: block.id, ri, ci },
      data: cell,
      children: groupByPath(doc, cell.blocks, blockPath, (b) => b.id, (b) => blockNode(doc, b), 0),
    })),
  }));
  return {
    key: block.id, kind: "block",
    label: "▦ table",
    preview: "",
    badges: tableBadges(block),
    blockId: block.id,
    target: { kind: "block", blockId: block.id },
    data: block,
    children: rows,
  };
}

/** A non-empty record/array section as a tree node, or null when there's nothing. */
function recordNode(key: string, label: string, obj: Record<string, unknown> | undefined): TreeNode | null {
  if (!obj) return null;
  const entries = Object.entries(obj);
  if (entries.length === 0) return null;
  return {
    key, kind: "tag", label,
    preview: `${entries.length}`,
    badges: [],
    data: obj,
    children: entries.map(([k, v]) => ({
      key: `${key}/${k}`, kind: "tag",
      label: k,
      preview: previewText(typeof v === "string" ? v : JSON.stringify(v)),
      badges: [],
      data: v,
      children: [],
    })),
  };
}

/** Whole-document model tree: Body, then Section bands, Footnotes, and Side-tables. */
export function buildModelTree(doc: Document): TreeNode[] {
  const roots: TreeNode[] = [];
  const blockChildren = (blocks: Block[]): TreeNode[] =>
    groupByPath(doc, blocks, blockPath, (b) => b.id, (b) => blockNode(doc, b), 0);

  roots.push({
    key: "$body", kind: "tag", label: "Body",
    preview: `${doc.blocks.length} blocks`,
    badges: [],
    data: { blocks: doc.blocks.length },
    children: blockChildren(doc.blocks),
  });

  const bandChildren: TreeNode[] = [];
  for (const band of BAND_CONTAINERS) {
    const blocks = doc.section[band];
    if (blocks && blocks.length > 0) {
      bandChildren.push({
        key: `$band/${band}`, kind: "tag", label: band,
        preview: `${blocks.length} blocks`,
        badges: [],
        data: { band, blocks: blocks.length },
        children: blockChildren(blocks),
      });
    }
  }
  if (bandChildren.length > 0) {
    roots.push({ key: "$bands", kind: "tag", label: "Section bands", preview: `${bandChildren.length}`, badges: [], data: doc.section, children: bandChildren });
  }

  const footnotes = doc.footnotes;
  if (footnotes && Object.keys(footnotes).length > 0) {
    roots.push({
      key: "$footnotes", kind: "tag", label: "Footnotes",
      preview: `${Object.keys(footnotes).length}`,
      badges: [],
      data: footnotes,
      children: Object.entries(footnotes).map(([noteId, paras]) => ({
        key: `$fn/${noteId}`, kind: "tag",
        label: `note ${noteId}`,
        preview: previewText(paras.map((p) => textOfRuns(p.runs)).join(" ")),
        badges: [`${paras.length} ¶`],
        blockId: paras[0]?.id,
        data: paras,
        children: blockChildren(paras),
      })),
    });
  }

  const sideKids = [
    recordNode("$styles", "Stylesheet", doc.stylesheet as unknown as Record<string, unknown> | undefined),
    recordNode("$lists", "Lists", doc.lists),
    recordNode("$tableStyles", "Table styles", doc.tableStyles),
    recordNode("$sdts", "Content controls (SDT)", doc.sdts),
    recordNode("$fields", "Fields", doc.fields),
    recordNode("$bookmarks", "Bookmarks", doc.bookmarks),
  ].filter((n): n is TreeNode => n !== null);
  if (sideKids.length > 0) {
    roots.push({ key: "$side", kind: "tag", label: "Side tables", preview: `${sideKids.length}`, badges: [], data: {}, children: sideKids });
  }

  return roots;
}

/** Cheap change-signature for the model tree (block revisions + side-table sizes). */
export function modelSignature(doc: Document): string {
  const blocks = doc.blocks.map((b) => `${b.id}.${b.revision}`).join(",");
  const counts = `${Object.keys(doc.sdts ?? {}).length},${Object.keys(doc.fields ?? {}).length},${Object.keys(doc.bookmarks ?? {}).length},${Object.keys(doc.footnotes ?? {}).length}`;
  return `${doc.blocks.length}:${blocks}|${counts}`;
}
