// Layer 6: accessibility — the canvas text mirror (critique A1, "milestone 6").
//
// Canvas-rendered text is invisible to assistive technology: the editor exposes
// only the IME proxy (role=textbox "Document editor"), so a screen-reader user
// could not read a single word. This module maintains an OFF-SCREEN DOM mirror
// of the whole document — one node per paragraph, with heading roles/levels and
// table structure — kept in sync with the model. The editable proxy `aria-owns`
// the mirror and points `aria-activedescendant` at the paragraph under the caret,
// so NVDA / Narrator can (a) read the body, (b) report the caret's paragraph,
// and (c) perceive the selection (aria-selected + a live-region announcement).
// It mirrors the discipline of the F6 drawing-object layer (see shapeAnnounce).

import type { EditorState } from "../editor/state";
import type { Block, Document } from "@cw/shared";
import { isCollapsed, styleById, textOfBlock, textOfRuns } from "@cw/shared";

export interface A11yMirror {
  sync(state: EditorState): void;
  announce(message: string): void;
  destroy(): void;
}

// sr-only: present for assistive tech, invisible and layout-neutral. Pinned to
// the origin at 1px + clip (the broadly-supported idiom) so it never adds scroll
// height. `user-select:none` keeps it out of mouse selections.
const SR_ONLY =
  "position:absolute;top:0;left:0;width:1px;height:1px;margin:0;padding:0;overflow:hidden;" +
  "clip:rect(0,0,0,0);white-space:normal;pointer-events:none;user-select:none;color:#000;";

const NAME_CAP = 80; // cap a selection announcement so it isn't a paragraph long

let mirrorSeq = 0;

/** "Heading 3" / "heading3" → 3; "Title" → 0; else null. Mirrors editorApp's
 *  labelLevel so the mirror's heading roles match the outline. */
function labelLevel(label: string | undefined): number | null {
  if (!label) return null;
  const m = /(?:^|\s)heading\s*([1-9])/i.exec(label);
  if (m) return Number(m[1]);
  if (/^\s*title\s*$/i.test(label)) return 0;
  if (/(?:^|\s)heading(?:\s|$)/i.test(label)) return 1;
  return null;
}

/** Heading level of a paragraph's named style, walking the basedOn chain (a
 *  custom "Section title" based on Heading 2 still reads as level 2). */
function headingLevelOf(doc: Document, styleId: string | undefined): number | null {
  if (!styleId) return null;
  const sheet = doc.stylesheet;
  let level = labelLevel(styleId);
  const seen = new Set<string>();
  for (
    let cur = sheet ? styleById(sheet, styleId) : undefined;
    cur && level === null && !seen.has(cur.id);
    cur = cur.basedOn && sheet ? styleById(sheet, cur.basedOn) : undefined
  ) {
    seen.add(cur.id);
    level = labelLevel(cur.name) ?? labelLevel(cur.id);
  }
  return level;
}

/** A flat, ordered description of the accessible nodes to render. `blockId` is
 *  present for anything that can hold the caret (paragraphs), so the proxy's
 *  aria-activedescendant / selection can resolve to it. */
interface Node {
  blockId: string | null;
  role: "paragraph" | "heading" | "object" | "table-open" | "table-close" | "row-open" | "row-close" | "cell-open" | "cell-close";
  level?: number | undefined; // heading level (0 = Title → aria-level 1)
  text?: string | undefined;
}

/** Walk the block tree (recursing tables) into an ordered node list. */
function describe(doc: Document, blocks: Block[], out: Node[]): void {
  for (const b of blocks) {
    if (b.kind === "paragraph") {
      const level = headingLevelOf(doc, b.style.namedStyle);
      out.push({ blockId: b.id, role: level === null ? "paragraph" : "heading", level: level ?? undefined, text: textOfRuns(b.runs) });
    } else if (b.kind === "table") {
      out.push({ blockId: null, role: "table-open" });
      for (const row of b.rows) {
        out.push({ blockId: null, role: "row-open" });
        for (const cell of row.cells) {
          out.push({ blockId: null, role: "cell-open" });
          describe(doc, cell.blocks, out);
          out.push({ blockId: null, role: "cell-close" });
        }
        out.push({ blockId: null, role: "row-close" });
      }
      out.push({ blockId: null, role: "table-close" });
    } else if (b.kind === "image") {
      out.push({ blockId: b.id, role: "object", text: "Image" });
    } else if (b.kind === "equation") {
      out.push({ blockId: b.id, role: "object", text: "Equation" });
    }
    // Shapes (floating, anchored) and custom blocks are announced through the
    // F6 drawing-object layer, not the in-flow reading mirror.
  }
}

export function createA11yMirror(proxyEl: HTMLElement): A11yMirror {
  const container = proxyEl.parentElement ?? document.body;
  const seq = mirrorSeq++;

  // ARIA live region for editor announcements (formatting toggles, shape moves,
  // selection summaries). aria-live=polite so it never interrupts reading.
  const live = document.createElement("div");
  live.setAttribute("aria-live", "polite");
  live.setAttribute("aria-atomic", "true");
  live.style.cssText = SR_ONLY;
  container.appendChild(live);

  // The document mirror: a role=document region the proxy owns.
  const root = document.createElement("div");
  root.id = `cw-a11y-doc-${seq}`;
  root.setAttribute("role", "document");
  root.setAttribute("aria-label", "Document contents");
  root.style.cssText = SR_ONLY;
  container.appendChild(root);
  proxyEl.setAttribute("aria-owns", root.id);

  const idFor = (blockId: string): string => `cw-a11y-${seq}-${blockId}`;

  let lastStructSig = "\0";
  let orderedIds: string[] = []; // paragraph ids in document order (for selection ranges)
  let nodeById = new Map<string, HTMLElement>();
  let lastSelKey = "\0";
  let announceTimer: ReturnType<typeof setTimeout> | null = null;

  /** Keep the mirror in sync with the document. The full DOM is rebuilt only when
   *  the STRUCTURE changes (blocks added/removed/reordered, heading level changed);
   *  a plain text edit patches just the affected node, so typing doesn't recreate
   *  hundreds of nodes per keystroke on a large document. */
  const rebuild = (doc: Document): void => {
    const nodes: Node[] = [];
    describe(doc, doc.blocks, nodes);
    // Structure signature excludes text — a text edit patches in place below.
    const structSig = nodes.map((n) => `${n.role}:${n.blockId ?? ""}:${n.level ?? ""}`).join("\n");
    if (structSig === lastStructSig) {
      // Same structure: update only the text nodes whose content changed.
      for (const n of nodes) {
        if (!n.blockId || n.text === undefined) continue;
        const el = nodeById.get(n.blockId);
        if (!el) continue;
        const want = n.text.length > 0 ? n.text : " ";
        if (el.textContent !== want) el.textContent = want;
      }
      return;
    }
    lastStructSig = structSig;

    root.textContent = "";
    orderedIds = [];
    nodeById = new Map();
    // A tiny stack lets table/row/cell wrappers nest their children.
    let cursor: HTMLElement = root;
    const stack: HTMLElement[] = [];
    const open = (tag: string, role: string): void => {
      const el = document.createElement(tag);
      el.setAttribute("role", role);
      cursor.appendChild(el);
      stack.push(cursor);
      cursor = el;
    };
    const close = (): void => {
      cursor = stack.pop() ?? root;
    };
    for (const n of nodes) {
      switch (n.role) {
        case "table-open": open("div", "table"); break;
        case "table-close": close(); break;
        case "row-open": open("div", "row"); break;
        case "row-close": close(); break;
        case "cell-open": open("div", "cell"); break;
        case "cell-close": close(); break;
        default: {
          const el = document.createElement(n.role === "heading" ? "div" : "p");
          if (n.blockId) {
            el.id = idFor(n.blockId);
            nodeById.set(n.blockId, el);
            orderedIds.push(n.blockId);
          }
          if (n.role === "heading") {
            el.setAttribute("role", "heading");
            el.setAttribute("aria-level", String(Math.max(1, n.level ?? 1)));
          }
          const t = n.text ?? "";
          el.textContent = t.length > 0 ? t : " "; // nbsp so an empty paragraph is still a node
          cursor.appendChild(el);
        }
      }
    }
  };

  /** Reflect the caret + selection onto the proxy and the mirror nodes. */
  const reflect = (state: EditorState): void => {
    const sel = state.selection;
    const focus = sel?.focus;
    // Caret → aria-activedescendant + the proxy's own name (immediate context).
    if (focus && nodeById.has(focus.blockId)) {
      proxyEl.setAttribute("aria-activedescendant", idFor(focus.blockId));
    } else {
      proxyEl.removeAttribute("aria-activedescendant");
    }

    // Selection → aria-selected on covered nodes + a debounced announcement.
    const collapsed = !sel || isCollapsed(sel);
    const ai = sel ? orderedIds.indexOf(sel.anchor.blockId) : -1;
    const fi = sel ? orderedIds.indexOf(sel.focus.blockId) : -1;
    const lo = Math.min(ai, fi);
    const hi = Math.max(ai, fi);
    orderedIds.forEach((id, i) => {
      const node = nodeById.get(id);
      if (!node) return;
      const on = !collapsed && ai >= 0 && fi >= 0 && i >= lo && i <= hi;
      if (on) node.setAttribute("aria-selected", "true");
      else node.removeAttribute("aria-selected");
    });

    const selKey = sel ? `${sel.anchor.blockId}:${sel.anchor.offset}-${sel.focus.blockId}:${sel.focus.offset}` : "";
    if (selKey !== lastSelKey) {
      lastSelKey = selKey;
      if (!collapsed && sel) announceSelection(state.doc, sel);
      else clearAnnouncement(); // collapsing must retire the last "Selected: …"
    }
  };

  /** Drop any announcement, pending or already rendered. Clearing the text matters
   *  as much as cancelling the timer: once the 300ms has elapsed the message is in
   *  the accessibility tree, and a review cursor reaching it later would read a
   *  selection the user no longer has. */
  const clearAnnouncement = (): void => {
    if (announceTimer) {
      clearTimeout(announceTimer);
      announceTimer = null;
    }
    if (live.textContent !== "") live.textContent = "";
  };

  /** Debounced "what is selected" summary (kept off the reading flow so a drag
   *  doesn't chatter). Single-paragraph → the selected text; multi → a count. */
  const announceSelection = (doc: Document, sel: NonNullable<EditorState["selection"]>): void => {
    if (announceTimer) clearTimeout(announceTimer);
    announceTimer = setTimeout(() => {
      let msg: string;
      if (sel.anchor.blockId === sel.focus.blockId) {
        const text = textOfBlock(doc, sel.anchor.blockId);
        const a = Math.min(sel.anchor.offset, sel.focus.offset);
        const b = Math.max(sel.anchor.offset, sel.focus.offset);
        const excerpt = text.slice(a, b);
        if (excerpt.trim() === "") return;
        msg = `Selected: ${excerpt.length > NAME_CAP ? excerpt.slice(0, NAME_CAP) + "…" : excerpt}`;
      } else {
        const ai = orderedIds.indexOf(sel.anchor.blockId);
        const fi = orderedIds.indexOf(sel.focus.blockId);
        const n = ai >= 0 && fi >= 0 ? Math.abs(fi - ai) + 1 : 2;
        msg = `Selected ${n} paragraphs`;
      }
      live.textContent = msg;
    }, 300);
  };

  return {
    sync(state: EditorState): void {
      rebuild(state.doc);
      reflect(state);
    },
    announce(message: string): void {
      live.textContent = message;
    },
    destroy(): void {
      if (announceTimer) clearTimeout(announceTimer);
      proxyEl.removeAttribute("aria-owns");
      proxyEl.removeAttribute("aria-activedescendant");
      live.remove();
      root.remove();
    },
  };
}
