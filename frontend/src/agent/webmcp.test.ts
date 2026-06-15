import { describe, expect, it, vi } from "vitest";
import type { CharStyle, Document, DocSelection, Paragraph } from "@cw/shared";
import type { Editor } from "../index";
import type { LayoutTree } from "../layout/layoutTree";
import { registerAgentTools, type ModelContextLike } from "./webmcp";

// ---- fakes ---------------------------------------------------------------------

const CHAR: CharStyle = {
  fontFamily: "Georgia, serif",
  fontSizePx: 16,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: "#202124",
};

const para = (id: string, text: string): Paragraph => ({
  kind: "paragraph",
  id,
  revision: 0,
  runs: [{ text, style: CHAR }],
  style: { align: "left", lineHeight: 1.5, spaceBeforePx: 0, spaceAfterPx: 0, indentFirstLinePx: 0, indentLeftPx: 0 },
});

function fakeDoc(): Document {
  return {
    // section just needs to be an object — band lookups are `doc.section[band] ?? []`.
    section: {} as Document["section"],
    blocks: [para("p1", "Hello colour world"), para("p2", "More colour here")],
  } as Document;
}

function fakeLayoutTree(): LayoutTree {
  return {
    pageWidthPx: 816,
    pageHeightPx: 1056,
    marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
    pages: [
      {
        index: 0,
        number: 1,
        widthPx: 816,
        heightPx: 1056,
        marginPx: { top: 96, right: 96, bottom: 96, left: 96 },
        contentTopPx: 96,
        contentBottomPx: 960,
        blocks: [
          {
            blockId: "p1",
            x: 96,
            y: 96,
            firstLineIndex: 0,
            lines: [
              {
                y: 0,
                height: 24,
                ascent: 19,
                fragments: [
                  { blockId: "p1", startOffset: 0, endOffset: 18, text: "Hello colour world", style: CHAR, x: 96, width: 200 },
                ],
              },
            ],
          },
        ],
      },
    ],
  };
}

interface Recorder {
  setCharStyle: ReturnType<typeof vi.fn>;
  addComment: ReturnType<typeof vi.fn>;
  searchReplaceAll: ReturnType<typeof vi.fn>;
  searchReplaceCurrent: ReturnType<typeof vi.fn>;
  dispatch: ReturnType<typeof vi.fn>;
  setMode: ReturnType<typeof vi.fn>;
}

function makeEditor(): { editor: Editor; rec: Recorder; getSelection: () => DocSelection | null } {
  const doc = fakeDoc();
  let selection: DocSelection | null = null;
  const rec: Recorder = {
    setCharStyle: vi.fn(),
    addComment: vi.fn(() => "thread1"),
    searchReplaceAll: vi.fn(() => 2),
    searchReplaceCurrent: vi.fn(() => ({ index: 1, total: 2 })),
    dispatch: vi.fn(),
    setMode: vi.fn(() => true),
  };
  const editor = {
    getDocument: () => doc,
    getSelection: () => selection,
    setSelection: (s: DocSelection | null) => {
      selection = s;
    },
    getLayoutTree: fakeLayoutTree,
    getLayoutInfo: () => ({ pageCount: 1, currentPage: 1 }),
    getMode: () => "edit" as const,
    setMode: rec.setMode,
    getReview: () => ({ suggestions: [], threads: [] }),
    // search: find first occurrence, set selection to it, report count.
    search: (query: string) => {
      let total = 0;
      let first: DocSelection | null = null;
      for (const p of doc.blocks as Paragraph[]) {
        const t = p.runs[0]!.text;
        let i = t.indexOf(query);
        while (i >= 0) {
          total++;
          if (!first) first = { anchor: { blockId: p.id, offset: i }, focus: { blockId: p.id, offset: i + query.length } };
          i = t.indexOf(query, i + 1);
        }
      }
      selection = first;
      return { index: total > 0 ? 1 : 0, total };
    },
    searchNav: () => ({ index: 1, total: 2 }),
    searchClear: () => {},
    searchReplaceCurrent: rec.searchReplaceCurrent,
    searchReplaceAll: rec.searchReplaceAll,
    setCharStyle: rec.setCharStyle,
    align: vi.fn(),
    dispatch: rec.dispatch,
    addComment: rec.addComment,
    replyToComment: vi.fn(),
    resolveThread: vi.fn(),
    acceptSuggestion: vi.fn(),
    rejectSuggestion: vi.fn(),
    acceptAllSuggestions: vi.fn(),
    rejectAllSuggestions: vi.fn(),
    undo: vi.fn(),
    redo: vi.fn(),
  } as unknown as Editor;
  return { editor, rec, getSelection: () => selection };
}

interface CapturedTool {
  name: string;
  execute: (args: Record<string, unknown>) => Promise<{ content: { type: "text"; text: string }[]; isError?: boolean }>;
}

function fakeContext(): { mc: ModelContextLike; tools: Map<string, CapturedTool>; signals: (AbortSignal | undefined)[] } {
  const tools = new Map<string, CapturedTool>();
  const signals: (AbortSignal | undefined)[] = [];
  const mc: ModelContextLike = {
    registerTool: (tool, opts) => {
      tools.set(tool.name, tool as unknown as CapturedTool);
      signals.push(opts?.signal);
    },
  };
  return { mc, tools, signals };
}

const parse = (r: { content: { text: string }[] }): unknown => JSON.parse(r.content[0]!.text);

// ---- tests ---------------------------------------------------------------------

describe("registerAgentTools", () => {
  it("registers all buckets by default and gates by capability", () => {
    const all = fakeContext();
    registerAgentTools(makeEditor().editor, {}, { modelContext: all.mc });
    expect(all.tools.has("get_document")).toBe(true);
    expect(all.tools.has("add_comment")).toBe(true);
    expect(all.tools.has("replace_text")).toBe(true);

    const readOnly = fakeContext();
    registerAgentTools(makeEditor().editor, {}, { modelContext: readOnly.mc, capabilities: ["read"] });
    expect(readOnly.tools.has("get_document")).toBe(true);
    expect(readOnly.tools.has("inspect_layout")).toBe(true);
    expect(readOnly.tools.has("add_comment")).toBe(false);
    expect(readOnly.tools.has("replace_text")).toBe(false);
  });

  it("namespaces tool names with the config name", () => {
    const { mc, tools } = fakeContext();
    registerAgentTools(makeEditor().editor, {}, { modelContext: mc, name: "doc1" });
    expect(tools.has("doc1_get_document")).toBe(true);
    expect(tools.has("get_document")).toBe(false);
  });

  it("get_document returns text or json", async () => {
    const { mc, tools } = fakeContext();
    registerAgentTools(makeEditor().editor, {}, { modelContext: mc });
    const text = await tools.get("get_document")!.execute({ format: "text" });
    expect(text.content[0]!.text).toBe("Hello colour world\nMore colour here");
    const json = await tools.get("get_document")!.execute({ format: "json" });
    expect((parse(json) as Document).blocks).toHaveLength(2);
  });

  it("inspect_layout dumps page/line/fragment geometry", async () => {
    const { mc, tools } = fakeContext();
    registerAgentTools(makeEditor().editor, {}, { modelContext: mc });
    const out = parse(await tools.get("inspect_layout")!.execute({})) as {
      pageCount: number;
      pages: { blocks: { lines: { fragments: { text: string; x: number }[] }[] }[] }[];
    };
    expect(out.pageCount).toBe(1);
    const frag = out.pages[0]!.blocks[0]!.lines[0]!.fragments[0]!;
    expect(frag.text).toBe("Hello colour world");
    expect(frag.x).toBe(96);
  });

  it("inspect_layout scopes to a block id", async () => {
    const { mc, tools } = fakeContext();
    registerAgentTools(makeEditor().editor, {}, { modelContext: mc });
    const out = parse(await tools.get("inspect_layout")!.execute({ blockId: "p1" })) as { pages: { blocks: unknown[] }[] };
    expect(out.pages[0]!.blocks).toHaveLength(1);
    const none = parse(await tools.get("inspect_layout")!.execute({ blockId: "nope" })) as { pages: { blocks: unknown[] }[] };
    expect(none.pages[0]!.blocks).toHaveLength(0);
  });

  it("replace_text searches then replaces all", async () => {
    const { editor, rec } = makeEditor();
    const { mc, tools } = fakeContext();
    registerAgentTools(editor, {}, { modelContext: mc });
    const out = parse(await tools.get("replace_text")!.execute({ find: "colour", replaceWith: "color", all: true })) as {
      replaced: number;
    };
    expect(out.replaced).toBe(2);
    expect(rec.searchReplaceAll).toHaveBeenCalledWith("color");
  });

  it("replace_text reports when text is not found", async () => {
    const { editor } = makeEditor();
    const { mc, tools } = fakeContext();
    registerAgentTools(editor, {}, { modelContext: mc });
    const out = await tools.get("replace_text")!.execute({ find: "zzz", replaceWith: "x" });
    expect(out.isError).toBe(true);
  });

  it("format_text finds text then applies a char-style patch", async () => {
    const { editor, rec } = makeEditor();
    const { mc, tools } = fakeContext();
    registerAgentTools(editor, {}, { modelContext: mc });
    await tools.get("format_text")!.execute({ find: "Hello", bold: true });
    expect(rec.setCharStyle).toHaveBeenCalledWith({ bold: true });
  });

  it("insert_text dispatches an insert command at the selection", async () => {
    const { editor, rec } = makeEditor();
    const { mc, tools } = fakeContext();
    registerAgentTools(editor, {}, { modelContext: mc });
    const out = await tools.get("insert_text")!.execute({ text: "X", find: "world" });
    expect(out.isError).toBeUndefined();
    expect(rec.dispatch).toHaveBeenCalledTimes(1);
  });

  it("add_comment anchors to found text and returns a thread id", async () => {
    const { editor, rec } = makeEditor();
    const { mc, tools } = fakeContext();
    registerAgentTools(editor, {}, { modelContext: mc });
    const out = parse(await tools.get("add_comment")!.execute({ body: "fix this", find: "world" })) as { threadId: string };
    expect(out.threadId).toBe("thread1");
    expect(rec.addComment).toHaveBeenCalledTimes(1);
  });

  it("add_comment fails without a selection", async () => {
    const { editor } = makeEditor();
    const { mc, tools } = fakeContext();
    registerAgentTools(editor, {}, { modelContext: mc });
    const out = await tools.get("add_comment")!.execute({ body: "orphan" });
    expect(out.isError).toBe(true);
  });

  it("set_document is only present when the host provides a setDocument hook", () => {
    const without = fakeContext();
    registerAgentTools(makeEditor().editor, {}, { modelContext: without.mc });
    expect(without.tools.has("set_document")).toBe(false);

    const setDocument = vi.fn();
    const withHook = fakeContext();
    registerAgentTools(makeEditor().editor, { setDocument }, { modelContext: withHook.mc });
    expect(withHook.tools.has("set_document")).toBe(true);
  });

  it("the disposer aborts the shared signal", () => {
    const { mc, signals } = fakeContext();
    const dispose = registerAgentTools(makeEditor().editor, {}, { modelContext: mc });
    expect(signals[0]?.aborted).toBe(false);
    dispose();
    expect(signals[0]?.aborted).toBe(true);
  });

  it("returns a no-op disposer when no model context is available", () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    const dispose = registerAgentTools(makeEditor().editor, {}, { modelContext: null as unknown as ModelContextLike });
    expect(typeof dispose).toBe("function");
    expect(() => dispose()).not.toThrow();
    warn.mockRestore();
  });
});
