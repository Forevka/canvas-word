// WebMCP agent bridge: exposes the live editor to AI agents as Model-Context tools
// (the W3C `navigator.modelContext` / `document.modelContext` surface, polyfilled by
// @mcp-b/global). Tools wrap the rich internal Editor — read & inspect (incl. a
// layout-geometry dump for debugging "renders weirdly" reports), suggest & comment,
// and direct edits. Registration is gated per capability and torn down via one
// AbortSignal. The polyfill itself is initialized by the caller (editorApp); this
// module only registers tools, so it stays trivially testable with a fake context.

import type { CharStyle, Document, DocSelection, Fragment, ParaStyle } from "@cw/shared";
import { blockById, DEFAULT_CHAR_STYLE, paragraphsOf, textOfBlock, textOfRuns } from "@cw/shared";
import type { Editor } from "../index";
import { dumpLayout, type LayoutDumpOptions } from "./layoutDump";

// ---- minimal Web Model Context shape (no dependency on the polyfill's global
//      type augmentation, so this module type-checks standalone and in tests) ----

interface AgentToolResult {
  content: { type: "text"; text: string }[];
  isError?: boolean;
}
interface AgentToolDef {
  name: string;
  description: string;
  inputSchema?: { type: "object"; properties?: Record<string, unknown>; required?: string[] };
  execute: (args: Record<string, unknown>) => AgentToolResult | Promise<AgentToolResult>;
}
export interface ModelContextLike {
  registerTool: (tool: AgentToolDef, options?: { signal?: AbortSignal }) => void;
}

export type AgentCapability = "read" | "suggest" | "edit";

export interface AgentToolsConfig {
  /** Which tool buckets to register. Default: all three. */
  capabilities?: AgentCapability[];
  /** Optional prefix so multiple editors on a page don't collide, e.g. "doc1" →
   *  "doc1_get_document". Default: no prefix. */
  name?: string;
  /** Inject the model-context object (tests / custom runtimes). Default: resolve
   *  `document.modelContext` then `navigator.modelContext` from the global. */
  modelContext?: ModelContextLike;
}

/** Host-supplied capabilities the core Editor doesn't expose directly. */
export interface AgentToolsContext {
  /** Current collaboration doc id, if any (surfaced by get_document_stats). */
  docId?: () => string | null;
  /** Replace the whole document (wraps the host's setDocument path). Enables the
   *  `set_document` tool when present. */
  setDocument?: (doc: Document) => void;
}

// ---- helpers -------------------------------------------------------------------

const COMMENT_STYLE: CharStyle = DEFAULT_CHAR_STYLE;

const ok = (text: string): AgentToolResult => ({ content: [{ type: "text", text }] });
const okJson = (value: unknown): AgentToolResult => ok(JSON.stringify(value, null, 2));
const fail = (message: string): AgentToolResult => ({ content: [{ type: "text", text: message }], isError: true });

const str = (v: unknown): string | undefined => (typeof v === "string" ? v : undefined);
const bool = (v: unknown): boolean | undefined => (typeof v === "boolean" ? v : undefined);
const num = (v: unknown): number | undefined => (typeof v === "number" && Number.isFinite(v) ? v : undefined);

function textFragment(text: string): Fragment {
  return [{ text, style: COMMENT_STYLE }];
}

/** Build search options, omitting absent keys (exactOptionalPropertyTypes). */
function searchOpts(args: Record<string, unknown>): { matchCase?: boolean; wholeWord?: boolean } {
  const o: { matchCase?: boolean; wholeWord?: boolean } = {};
  const mc = bool(args.matchCase);
  if (mc !== undefined) o.matchCase = mc;
  const ww = bool(args.wholeWord);
  if (ww !== undefined) o.wholeWord = ww;
  return o;
}

/** Plain text of the whole body (one line per paragraph, incl. table cells). */
function documentText(doc: Document): string {
  return paragraphsOf(doc)
    .map((p) => textOfRuns(p.runs))
    .join("\n");
}

/** Best-effort plain text of the current selection. */
function selectionText(doc: Document, sel: DocSelection): string {
  const paras = paragraphsOf(doc);
  const idxOf = (id: string): number => paras.findIndex((p) => p.id === id);
  let a = sel.anchor;
  let f = sel.focus;
  let aIdx = idxOf(a.blockId);
  let fIdx = idxOf(f.blockId);
  if (aIdx > fIdx || (aIdx === fIdx && a.offset > f.offset)) {
    [a, f] = [f, a];
    [aIdx, fIdx] = [fIdx, aIdx];
  }
  if (aIdx < 0 || fIdx < 0) {
    if (sel.anchor.blockId === sel.focus.blockId) {
      const p = paras.find((pp) => pp.id === sel.anchor.blockId);
      if (p) {
        const t = textOfRuns(p.runs);
        return t.slice(Math.min(sel.anchor.offset, sel.focus.offset), Math.max(sel.anchor.offset, sel.focus.offset));
      }
    }
    return "";
  }
  if (aIdx === fIdx) return textOfRuns(paras[aIdx]!.runs).slice(a.offset, f.offset);
  const parts: string[] = [textOfRuns(paras[aIdx]!.runs).slice(a.offset)];
  for (let i = aIdx + 1; i < fIdx; i++) parts.push(textOfRuns(paras[i]!.runs));
  parts.push(textOfRuns(paras[fIdx]!.runs).slice(0, f.offset));
  return parts.join("\n");
}

function resolveModelContext(): ModelContextLike | null {
  const g = globalThis as unknown as {
    document?: { modelContext?: ModelContextLike };
    navigator?: { modelContext?: ModelContextLike };
  };
  return g.document?.modelContext ?? g.navigator?.modelContext ?? null;
}

// ---- registration --------------------------------------------------------------

/** Register the editor's agent tools on the Web Model Context. Returns a disposer
 *  that unregisters every tool (via the shared AbortSignal). No-op disposer when no
 *  model-context runtime is available. */
export function registerAgentTools(
  editor: Editor,
  ctx: AgentToolsContext = {},
  config: AgentToolsConfig = {},
): () => void {
  const mc = config.modelContext ?? resolveModelContext();
  if (!mc) {
    console.warn("[wordcanvas] agentTools enabled but no navigator.modelContext is available; tools not registered.");
    return () => {};
  }

  const caps = new Set<AgentCapability>(config.capabilities ?? ["read", "suggest", "edit"]);
  const prefix = config.name ? `${config.name}_` : "";
  const ac = new AbortController();

  const register = (tool: AgentToolDef): void => {
    const wrapped: AgentToolDef = {
      ...tool,
      name: prefix + tool.name,
      execute: async (args) => {
        try {
          return await tool.execute(args ?? {});
        } catch (e) {
          return fail(`${tool.name} failed: ${e instanceof Error ? e.message : String(e)}`);
        }
      },
    };
    mc.registerTool(wrapped, { signal: ac.signal });
  };

  // ---- read & inspect (always) -------------------------------------------------

  register({
    name: "get_document",
    description:
      "Read the open document. format='text' returns plain text (one line per paragraph); format='json' returns the full document model (blocks, runs, styles, tables).",
    inputSchema: { type: "object", properties: { format: { type: "string", enum: ["text", "json"], description: "text (default) or json" } } },
    execute: (args) => {
      const doc = editor.getDocument();
      return (str(args.format) ?? "text") === "json" ? okJson(doc) : ok(documentText(doc));
    },
  });

  register({
    name: "get_selection",
    description: "Get the current caret/selection (anchor + focus positions) and the selected text.",
    inputSchema: { type: "object", properties: {} },
    execute: () => {
      const sel = editor.getSelection();
      if (!sel) return okJson({ selection: null });
      return okJson({ selection: sel, text: selectionText(editor.getDocument(), sel) });
    },
  });

  register({
    name: "search_document",
    description:
      "Find text in the document. Highlights matches and moves the selection to the first match. Returns the match count and current 1-based index.",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Text to find" },
        matchCase: { type: "boolean" },
        wholeWord: { type: "boolean" },
      },
      required: ["query"],
    },
    execute: (args) => {
      const query = str(args.query);
      if (!query) return fail("query is required");
      const state = editor.search(query, searchOpts(args));
      return okJson({ total: state.total, current: state.index });
    },
  });

  register({
    name: "inspect_layout",
    description:
      "Dump the laid-out geometry (pages, blocks, line boxes, text fragment positions in page-local CSS px) for debugging rendering / text-placement issues. Scope with page (0-based) and/or blockId to bound the payload.",
    inputSchema: {
      type: "object",
      properties: {
        page: { type: "number", description: "0-based page index to restrict to" },
        blockId: { type: "string", description: "Restrict to a single model block id" },
        includeText: { type: "boolean", description: "Include rendered fragment text (default true)" },
        maxFragmentsPerLine: { type: "number" },
      },
    },
    execute: (args) => {
      const opts: LayoutDumpOptions = {};
      const page = num(args.page);
      if (page !== undefined) opts.page = page;
      const blockId = str(args.blockId);
      if (blockId) opts.blockId = blockId;
      const includeText = bool(args.includeText);
      if (includeText !== undefined) opts.includeText = includeText;
      const maxFrags = num(args.maxFragmentsPerLine);
      if (maxFrags !== undefined) opts.maxFragmentsPerLine = maxFrags;
      return okJson(dumpLayout(editor.getLayoutTree(), opts));
    },
  });

  register({
    name: "get_document_stats",
    description: "Summary stats: page count, current page, paragraph count, and the collaboration doc id (if online).",
    inputSchema: { type: "object", properties: {} },
    execute: () => {
      const info = editor.getLayoutInfo();
      const doc = editor.getDocument();
      return okJson({
        pageCount: info.pageCount,
        currentPage: info.currentPage,
        blockCount: doc.blocks.length,
        paragraphCount: paragraphsOf(doc).length,
        mode: editor.getMode(),
        docId: ctx.docId?.() ?? null,
      });
    },
  });

  // ---- suggest & comment -------------------------------------------------------

  if (caps.has("suggest")) {
    register({
      name: "set_mode",
      description:
        "Switch editor mode: 'edit' (direct edits), 'suggest' (edits become tracked changes), or 'view' (read-only). Returns whether the switch was allowed.",
      inputSchema: { type: "object", properties: { mode: { type: "string", enum: ["edit", "suggest", "view"] } }, required: ["mode"] },
      execute: (args) => {
        const mode = str(args.mode);
        if (mode !== "edit" && mode !== "suggest" && mode !== "view") return fail("mode must be edit|suggest|view");
        return okJson({ ok: editor.setMode(mode), mode: editor.getMode() });
      },
    });

    register({
      name: "get_review",
      description: "Read the review overlay: tracked-change suggestions and comment threads.",
      inputSchema: { type: "object", properties: {} },
      execute: () => okJson(editor.getReview()),
    });

    register({
      name: "add_comment",
      description:
        "Add a comment anchored to a text range. Provide 'find' to locate+select the text first, or comment on the current selection. Returns the new thread id.",
      inputSchema: {
        type: "object",
        properties: {
          body: { type: "string", description: "Comment text" },
          find: { type: "string", description: "Optional: text to find and anchor the comment to" },
        },
        required: ["body"],
      },
      execute: (args) => {
        const body = str(args.body);
        if (!body) return fail("body is required");
        const find = str(args.find);
        if (find) {
          const state = editor.search(find);
          if (state.total === 0) return fail(`text not found: ${find}`);
        }
        if (!editor.getSelection()) return fail("no selection to anchor the comment (pass 'find' or select first)");
        const id = editor.addComment(textFragment(body));
        if (find) editor.searchClear();
        return id ? okJson({ threadId: id }) : fail("could not add comment (view mode or empty selection)");
      },
    });

    register({
      name: "reply_to_comment",
      description: "Reply to an existing comment thread by id.",
      inputSchema: { type: "object", properties: { threadId: { type: "string" }, body: { type: "string" } }, required: ["threadId", "body"] },
      execute: (args) => {
        const threadId = str(args.threadId);
        const body = str(args.body);
        if (!threadId || !body) return fail("threadId and body are required");
        editor.replyToComment(threadId, textFragment(body));
        return ok("replied");
      },
    });

    register({
      name: "resolve_thread",
      description: "Resolve (or reopen) a comment thread by id.",
      inputSchema: { type: "object", properties: { threadId: { type: "string" }, resolved: { type: "boolean", description: "default true" } }, required: ["threadId"] },
      execute: (args) => {
        const threadId = str(args.threadId);
        if (!threadId) return fail("threadId is required");
        editor.resolveThread(threadId, bool(args.resolved) ?? true);
        return ok("done");
      },
    });

    register({
      name: "accept_suggestion",
      description: "Accept a tracked-change suggestion by id, or all suggestions when id is omitted.",
      inputSchema: { type: "object", properties: { id: { type: "string", description: "Suggestion id; omit to accept all" } } },
      execute: (args) => {
        const id = str(args.id);
        if (id) editor.acceptSuggestion(id);
        else editor.acceptAllSuggestions();
        return ok(id ? `accepted ${id}` : "accepted all");
      },
    });

    register({
      name: "reject_suggestion",
      description: "Reject a tracked-change suggestion by id, or all suggestions when id is omitted.",
      inputSchema: { type: "object", properties: { id: { type: "string", description: "Suggestion id; omit to reject all" } } },
      execute: (args) => {
        const id = str(args.id);
        if (id) editor.rejectSuggestion(id);
        else editor.rejectAllSuggestions();
        return ok(id ? `rejected ${id}` : "rejected all");
      },
    });
  }

  // ---- direct edits ------------------------------------------------------------

  if (caps.has("edit")) {
    register({
      name: "replace_text",
      description:
        "Find text and replace it. all=true replaces every match (returns the count); otherwise replaces the first match.",
      inputSchema: {
        type: "object",
        properties: {
          find: { type: "string" },
          replaceWith: { type: "string" },
          all: { type: "boolean" },
          matchCase: { type: "boolean" },
          wholeWord: { type: "boolean" },
        },
        required: ["find", "replaceWith"],
      },
      execute: (args) => {
        const find = str(args.find);
        const replaceWith = str(args.replaceWith);
        if (find === undefined || replaceWith === undefined) return fail("find and replaceWith are required");
        const state = editor.search(find, searchOpts(args));
        if (state.total === 0) {
          editor.searchClear();
          return fail(`text not found: ${find}`);
        }
        let count: number;
        if (bool(args.all)) count = editor.searchReplaceAll(replaceWith);
        else {
          editor.searchReplaceCurrent(replaceWith);
          count = 1;
        }
        editor.searchClear();
        return okJson({ replaced: count });
      },
    });

    register({
      name: "insert_text",
      description:
        "Insert text at the current selection (replacing it if it's a range). Pass 'find' to first select a target — note this REPLACES the found text; use replace_text for plain replacement.",
      inputSchema: { type: "object", properties: { text: { type: "string" }, find: { type: "string" } }, required: ["text"] },
      execute: async (args) => {
        const text = str(args.text);
        if (text === undefined) return fail("text is required");
        const find = str(args.find);
        if (find) {
          const state = editor.search(find);
          if (state.total === 0) {
            editor.searchClear();
            return fail(`text not found: ${find}`);
          }
        }
        if (!editor.getSelection()) return fail("no selection (pass 'find' or select_range first)");
        const { insertText } = await import("../editor/commands");
        editor.dispatch(insertText(text));
        if (find) editor.searchClear();
        return ok("inserted");
      },
    });

    register({
      name: "format_text",
      description:
        "Apply character formatting to a range. Pass 'find' to target text, or format the current selection. clear=true resets bold/italic/underline/strikethrough/highlight.",
      inputSchema: {
        type: "object",
        properties: {
          find: { type: "string" },
          bold: { type: "boolean" },
          italic: { type: "boolean" },
          underline: { type: "boolean" },
          strikethrough: { type: "boolean" },
          color: { type: "string", description: "CSS color, e.g. #c00" },
          fontFamily: { type: "string" },
          fontSizePx: { type: "number" },
          highlightColor: { type: "string" },
          clear: { type: "boolean" },
        },
      },
      execute: (args) => {
        const find = str(args.find);
        if (find) {
          const state = editor.search(find);
          if (state.total === 0) {
            editor.searchClear();
            return fail(`text not found: ${find}`);
          }
        }
        if (!editor.getSelection()) return fail("no selection (pass 'find' or select_range first)");
        let patch: Partial<CharStyle>;
        if (bool(args.clear)) {
          patch = { bold: false, italic: false, underline: false, strikethrough: false, highlightColor: undefined };
        } else {
          patch = {};
          const b = bool(args.bold);
          if (b !== undefined) patch.bold = b;
          const i = bool(args.italic);
          if (i !== undefined) patch.italic = i;
          const u = bool(args.underline);
          if (u !== undefined) patch.underline = u;
          const s = bool(args.strikethrough);
          if (s !== undefined) patch.strikethrough = s;
          const color = str(args.color);
          if (color) patch.color = color;
          const fontFamily = str(args.fontFamily);
          if (fontFamily) patch.fontFamily = fontFamily;
          const size = num(args.fontSizePx);
          if (size !== undefined) patch.fontSizePx = size;
          const hl = str(args.highlightColor);
          if (hl) patch.highlightColor = hl;
        }
        editor.setCharStyle(patch);
        if (find) editor.searchClear();
        return ok("formatted");
      },
    });

    register({
      name: "set_alignment",
      description: "Set paragraph alignment of the current selection: left | center | right | justify.",
      inputSchema: { type: "object", properties: { align: { type: "string", enum: ["left", "center", "right", "justify"] } }, required: ["align"] },
      execute: (args) => {
        const align = str(args.align);
        if (align !== "left" && align !== "center" && align !== "right" && align !== "justify") return fail("invalid align");
        editor.align(align as ParaStyle["align"]);
        return ok("aligned");
      },
    });

    register({
      name: "select_range",
      description:
        "Set the selection to an explicit range using model positions (block id + UTF-16 offset). Use get_document(json) / inspect_layout to discover block ids.",
      inputSchema: {
        type: "object",
        properties: {
          anchorBlockId: { type: "string" },
          anchorOffset: { type: "number" },
          focusBlockId: { type: "string" },
          focusOffset: { type: "number" },
        },
        required: ["anchorBlockId", "anchorOffset", "focusBlockId", "focusOffset"],
      },
      execute: (args) => {
        const ab = str(args.anchorBlockId);
        const ao = num(args.anchorOffset);
        const fb = str(args.focusBlockId);
        const fo = num(args.focusOffset);
        if (ab === undefined || ao === undefined || fb === undefined || fo === undefined) return fail("anchor/focus block id + offset are required");
        // Validate against the live model — agents can hallucinate ids/offsets, and
        // an out-of-range selection breaks layout/paint downstream.
        const doc = editor.getDocument();
        if (!blockById(doc, ab)) return fail(`anchor block '${ab}' not found`);
        if (!blockById(doc, fb)) return fail(`focus block '${fb}' not found`);
        const clamp = (o: number, bid: string): number => Math.max(0, Math.min(o, textOfBlock(doc, bid).length));
        editor.setSelection({ anchor: { blockId: ab, offset: clamp(ao, ab) }, focus: { blockId: fb, offset: clamp(fo, fb) } });
        return ok("selection set");
      },
    });

    register({
      name: "undo",
      description: "Undo the last edit.",
      inputSchema: { type: "object", properties: {} },
      execute: () => {
        editor.undo();
        return ok("undone");
      },
    });

    register({
      name: "redo",
      description: "Redo the last undone edit.",
      inputSchema: { type: "object", properties: {} },
      execute: () => {
        editor.redo();
        return ok("redone");
      },
    });

    if (ctx.setDocument) {
      register({
        name: "set_document",
        description:
          "Replace the entire document with a new document model (JSON, same shape as get_document(json) returns). Drops undo history — use sparingly.",
        inputSchema: { type: "object", properties: { json: { type: "string", description: "Document model as a JSON string" } }, required: ["json"] },
        execute: (args) => {
          const json = str(args.json);
          if (!json) return fail("json is required");
          let doc: Document;
          try {
            doc = JSON.parse(json) as Document;
          } catch (e) {
            return fail(`invalid JSON: ${e instanceof Error ? e.message : String(e)}`);
          }
          ctx.setDocument!(doc);
          return ok("document replaced");
        },
      });
    }
  }

  return () => ac.abort();
}
