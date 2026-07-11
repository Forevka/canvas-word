// 3-way merge editor — the Phase 2 resolution surface for merge3. The document
// currently open is MINE; the user picks a common BASE and the other version
// (THEIRS). merge3 auto-merges every non-conflicting change and lists the
// conflicts; each conflict is resolved by choosing Mine / Theirs / Both / Base,
// and a live preview re-assembles on every choice. Applying swaps the assembled
// document into the editor.
//
// Reuses the render stack (layout engine + paint) like compareView, and the
// isomorphic merge3/assembleMerge core.

import {
  assembleMerge,
  blockById,
  merge3,
  textOfRuns,
  type Document,
  type MergeChoice,
  type Merge3Result,
  type MergeSegment,
} from "@cw/shared";
import { createDialogShell } from "./dialogShell";
import { importDocx } from "../import/docx/importDocx";
import { createLayoutEngine } from "../layout/engine";
import { createPaintLayer, type PaintScheduler } from "../paint/renderer";
import { caretRect } from "../layout/geometry";
import type { LayoutTree } from "../layout/layoutTree";
import type { ResolvedTheme } from "../config";
import type { CustomFontRegistry } from "../fonts/customRegistry";

export interface Merge3ViewOptions {
  /** The document open in the editor — MINE. */
  mine: Document;
  mineName?: string;
  /** Preset the common ancestor. When both `base` and `theirs` are supplied the
   *  setup panel is skipped and the merge editor opens straight away. */
  base?: Document;
  /** Preset the other version (see `base`). */
  theirs?: Document;
  fontRegistry?: CustomFontRegistry;
  theme?: ResolvedTheme;
  /** Apply the assembled merge result back to the editor. */
  onApply(doc: Document): void;
  onClose?(): void;
}

export interface Merge3ViewHandle {
  close(): void;
}

const CHOICES: { key: MergeChoice; label: string }[] = [
  { key: "mine", label: "Mine" },
  { key: "theirs", label: "Theirs" },
  { key: "both", label: "Both" },
  { key: "base", label: "Base" },
];

const CSS = `
.cw-mrg-backdrop{position:fixed;inset:0;z-index:2147483000;pointer-events:none;}
.cw-mrg-modal{position:fixed;pointer-events:auto;width:min(96vw,1200px);height:86vh;display:flex;flex-direction:column;
  background:#fff;color:#201f1e;border:1px solid #c8c6c4;border-radius:8px;box-shadow:0 12px 40px rgba(0,0,0,.3);
  font:13px 'Segoe UI',Roboto,sans-serif;}
.cw-mrg-head{display:flex;align-items:center;gap:10px;padding:10px 12px;border-bottom:1px solid #edebe9;cursor:move;}
.cw-mrg-head h2{margin:0;font-size:15px;font-weight:600;flex:1;}
.cw-mrg-count{font-size:12px;color:#605e5c;background:#f3f2f1;border-radius:10px;padding:2px 10px;}
.cw-mrg-x{border:0;background:transparent;font-size:20px;line-height:1;cursor:pointer;color:#605e5c;padding:2px 6px;border-radius:4px;}
.cw-mrg-x:hover{background:#f3f2f1;color:#201f1e;}
.cw-mrg-body{display:flex;flex:1 1 auto;min-height:0;}
.cw-mrg-list{flex:0 0 360px;overflow:auto;border-right:1px solid #edebe9;background:#faf9f8;}
.cw-mrg-preview{flex:1 1 auto;min-width:0;overflow:auto;background:#e6e6e6;padding:16px;}
.cw-mrg-setup{padding:22px;display:flex;flex-direction:column;gap:16px;max-width:560px;}
.cw-mrg-setup h3{margin:0 0 2px;font-size:14px;}
.cw-mrg-setup p{margin:0;color:#605e5c;}
.cw-mrg-field{display:flex;flex-direction:column;gap:6px;}
.cw-mrg-field label{font-weight:600;}
.cw-mrg-row{padding:8px 10px;border-bottom:1px solid #efedeb;}
.cw-mrg-row.unresolved{background:#fff9e6;}
.cw-mrg-row.active{background:#eef0ff;}
.cw-mrg-rowtop{display:flex;align-items:center;gap:6px;margin-bottom:6px;cursor:pointer;}
.cw-mrg-badge{font-size:10px;font-weight:700;text-transform:uppercase;padding:1px 6px;border-radius:4px;color:#fff;background:#c19c00;}
.cw-mrg-badge.done{background:#107c41;}
.cw-mrg-title{flex:1;font-weight:600;}
.cw-mrg-sides{display:flex;flex-direction:column;gap:3px;margin-bottom:6px;font-size:12px;}
.cw-mrg-side{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.cw-mrg-side b{display:inline-block;width:52px;color:#605e5c;font-weight:600;}
.cw-mrg-choices{display:flex;gap:5px;}
.cw-mrg-choice{border:1px solid #c8c6c4;background:#fff;border-radius:4px;padding:3px 10px;font-size:12px;cursor:pointer;}
.cw-mrg-choice:hover{background:#f3f2f1;}
.cw-mrg-choice.sel{background:#6b3fa0;border-color:#6b3fa0;color:#fff;}
.cw-mrg-empty{padding:26px 16px;text-align:center;color:#605e5c;}
.cw-mrg-foot{display:flex;align-items:center;gap:8px;padding:10px 14px;border-top:1px solid #edebe9;}
.cw-mrg-foot .spacer{flex:1;}
.cw-mrg-note{color:#8a6d00;font-size:12px;}
.cw-mrg-btn{border:1px solid #c8c6c4;background:#fff;border-radius:4px;padding:7px 16px;font-size:13px;cursor:pointer;}
.cw-mrg-btn:hover{background:#f3f2f1;}
.cw-mrg-btn:disabled{opacity:.5;cursor:default;}
.cw-mrg-btn.primary{background:#6b3fa0;border-color:#6b3fa0;color:#fff;}
.cw-mrg-btn.primary:hover{background:#5a3488;}
`;

const truncate = (s: string, n = 60): string => (s.length > n ? s.slice(0, n - 1) + "…" : s);
const blocksText = (blocks: { kind: string } []): string =>
  blocks.map((b) => ("runs" in b ? textOfRuns((b as { runs: Parameters<typeof textOfRuns>[0] }).runs) : `<${b.kind}>`)).join(" ¶ ") || "(nothing)";

export function showMerge3View(opts: Merge3ViewOptions): Merge3ViewHandle {
  const shell = createDialogShell({
    prefix: "cw-mrg",
    cssId: "cw-mrg-css",
    css: CSS,
    title: `3-way Merge — ${opts.mineName ?? "current document"} is “Mine”`,
    extraNoDrag: ".cw-mrg-body, .cw-mrg-foot",
    onClose: () => {
      painter?.destroy();
      opts.onClose?.();
    },
  });

  let base: Document | null = null;
  let theirs: Document | null = null;
  let res: Merge3Result | null = null;
  const choices: Record<string, MergeChoice> = {};
  let activeId: string | null = null;

  const count = document.createElement("span");
  count.className = "cw-mrg-count";
  shell.head.insertBefore(count, shell.closeBtn);

  shell.body.className = "cw-mrg-body";
  let painter: PaintScheduler | null = null;
  let engine: ReturnType<typeof createLayoutEngine> | null = null;
  let tree: LayoutTree | null = null;

  // ---- setup panel: choose the base + other version ----
  const renderSetup = (): void => {
    shell.body.innerHTML = "";
    shell.foot.innerHTML = "";
    count.textContent = "Choose versions";
    const setup = document.createElement("div");
    setup.className = "cw-mrg-setup";
    const intro = document.createElement("div");
    intro.innerHTML = `<h3>Pick the two other versions</h3><p>The document open in the editor is <b>Mine</b>. Choose the common ancestor (<b>Base</b>) and the other edited copy (<b>Theirs</b>).</p>`;
    const mkField = (labelText: string, onFile: (d: Document | null) => void): HTMLElement => {
      const field = document.createElement("div");
      field.className = "cw-mrg-field";
      const label = document.createElement("label");
      label.textContent = labelText;
      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document";
      input.addEventListener("change", async () => {
        const file = input.files?.[0];
        if (!file) return onFile(null);
        try {
          onFile((await importDocx(file)).doc);
          refreshSetupButton();
        } catch (e) {
          alert(`Could not read "${file.name}": ${e instanceof Error ? e.message : String(e)}`);
          onFile(null);
        }
      });
      field.append(label, input);
      return field;
    };
    setup.append(intro, mkField("Base (common ancestor) .docx", (d) => (base = d)), mkField("Theirs (other version) .docx", (d) => (theirs = d)));
    shell.body.appendChild(setup);

    const cancel = mkBtn("Cancel", false, () => shell.close());
    const start = mkBtn("Start merge", true, () => {
      if (base && theirs) {
        res = merge3(base, opts.mine, theirs, { docId: "merge" });
        renderMerge();
      }
    });
    start.disabled = true;
    const spacer = document.createElement("div");
    spacer.className = "spacer";
    shell.foot.append(spacer, cancel, start);
    refreshSetupButton = () => (start.disabled = !(base && theirs));
  };
  let refreshSetupButton = (): void => {};

  // ---- merge editor ----
  const list = document.createElement("div");
  list.className = "cw-mrg-list";
  const preview = document.createElement("div");
  preview.className = "cw-mrg-preview";

  const firstBlockIdOfChoice = (seg: MergeSegment, choice: MergeChoice): string | null => {
    const blocks = choice === "theirs" ? seg.theirs : choice === "base" ? seg.base : seg.mine;
    return blocks && blocks[0] ? blocks[0].id : null;
  };

  const scrollToConflict = (seg: MergeSegment): void => {
    if (!tree || !painter) return;
    const id = firstBlockIdOfChoice(seg, choices[seg.id!] ?? "mine");
    if (!id || !blockById(assembled(), id)) return;
    const cr = caretRect(tree, { blockId: id, offset: 0 });
    if (cr) {
      const at = painter.caretToContainer(cr);
      if (at) preview.scrollTop = Math.max(0, at.top - 90);
    }
  };

  const assembled = (): Document => assembleMerge(res!, choices);

  const renderPreview = (): void => {
    if (!engine || !painter) return;
    engine.reset();
    tree = engine.layout(assembled());
    painter.setTree(tree);
  };

  const unresolvedCount = (): number => res!.conflicts.filter((c) => !choices[c.id!]).length;

  const renderList = (): void => {
    list.innerHTML = "";
    if (res!.conflicts.length === 0) {
      const empty = document.createElement("div");
      empty.className = "cw-mrg-empty";
      empty.textContent = "No conflicts — every change merged automatically. Apply to finish.";
      list.appendChild(empty);
      return;
    }
    res!.conflicts.forEach((seg, i) => {
      const chosen = choices[seg.id!];
      const row = document.createElement("div");
      row.className = "cw-mrg-row" + (chosen ? "" : " unresolved") + (seg.id === activeId ? " active" : "");
      const top = document.createElement("div");
      top.className = "cw-mrg-rowtop";
      const badge = document.createElement("span");
      badge.className = "cw-mrg-badge" + (chosen ? " done" : "");
      badge.textContent = chosen ? "Resolved" : "Conflict";
      const title = document.createElement("span");
      title.className = "cw-mrg-title";
      title.textContent = `Conflict ${i + 1}`;
      top.append(badge, title);
      top.addEventListener("click", () => {
        activeId = seg.id!;
        scrollToConflict(seg);
        renderList();
      });
      const sides = document.createElement("div");
      sides.className = "cw-mrg-sides";
      const mkSide = (name: string, blocks: MergeSegment["mine"]): HTMLElement => {
        const d = document.createElement("div");
        d.className = "cw-mrg-side";
        d.innerHTML = `<b>${name}</b>`;
        d.append(document.createTextNode(truncate(blocksText(blocks ?? []))));
        return d;
      };
      sides.append(mkSide("Base", seg.base), mkSide("Mine", seg.mine), mkSide("Theirs", seg.theirs));
      const ch = document.createElement("div");
      ch.className = "cw-mrg-choices";
      for (const c of CHOICES) {
        const btn = document.createElement("button");
        btn.className = "cw-mrg-choice" + (chosen === c.key ? " sel" : "");
        btn.textContent = c.label;
        btn.addEventListener("click", () => {
          choices[seg.id!] = c.key;
          activeId = seg.id!;
          renderPreview();
          renderList();
          updateFoot();
          scrollToConflict(seg);
        });
        ch.appendChild(btn);
      }
      row.append(top, sides, ch);
      list.appendChild(row);
    });
  };

  let applyBtn: HTMLButtonElement | null = null;
  let note: HTMLElement | null = null;
  const updateFoot = (): void => {
    const n = res!.conflicts.length;
    const left = unresolvedCount();
    count.textContent = `${res!.stats.autoMerged} auto-merged · ${n} conflict${n === 1 ? "" : "s"}${n ? ` · ${n - left} resolved` : ""}`;
    if (note) note.textContent = left > 0 ? `${left} unresolved conflict${left === 1 ? "" : "s"} will default to Mine` : "";
  };

  const renderMerge = (): void => {
    shell.body.innerHTML = "";
    shell.foot.innerHTML = "";
    shell.body.append(list, preview);
    engine = createLayoutEngine(opts.fontRegistry);
    painter = createPaintLayer(preview, {
      chrome: false,
      ...(opts.theme ? { theme: opts.theme } : {}),
      ...(opts.fontRegistry ? { fontRegistry: opts.fontRegistry } : {}),
    });
    note = document.createElement("span");
    note.className = "cw-mrg-note";
    const spacer = document.createElement("div");
    spacer.className = "spacer";
    applyBtn = mkBtn("Apply merged result", true, () => {
      opts.onApply(assembled());
      shell.close();
    });
    shell.foot.append(note, spacer, mkBtn("Cancel", false, () => shell.close()), applyBtn);
    renderPreview();
    renderList();
    updateFoot();
  };

  function mkBtn(label: string, primary: boolean, onClick: () => void): HTMLButtonElement {
    const b = document.createElement("button");
    b.className = "cw-mrg-btn" + (primary ? " primary" : "");
    b.textContent = label;
    b.addEventListener("click", onClick);
    return b;
  }

  // Presets supplied (programmatic launch) → straight to the merge editor;
  // otherwise show the file-picking setup panel.
  if (opts.base && opts.theirs) {
    base = opts.base;
    theirs = opts.theirs;
    res = merge3(base, opts.mine, theirs, { docId: "merge" });
    renderMerge();
  } else {
    renderSetup();
  }
  return { close: () => shell.close() };
}
