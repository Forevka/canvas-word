// E2 / E3 — keyboard routing for the drawing-object layer in selectionController:
//   F6 / Shift+F6            → cycleObjectFocus (enter + step the object cycle)
//   Tab / Shift+Tab (object) → cycleObjectFocus (walk between objects)
//   Arrows (object selected) → nudgeSelectedObject (1px, 10px with Shift)
//
// The controller wires real DOM/window listeners and geometry.ts creates a
// measuring <canvas> at import; this repo's vitest runs in node, so we stub a
// minimal document + window and a fake container that captures the keydown
// handler, then invoke it with synthetic events and assert the deps it calls.

import { beforeAll, describe, expect, it, vi } from "vitest";

interface FakeContainer {
  tabIndex: number;
  style: Record<string, string>;
  handlers: Record<string, (ev: unknown) => void>;
  addEventListener(type: string, fn: (ev: unknown) => void): void;
  removeEventListener(): void;
}

beforeAll(() => {
  const g = globalThis as unknown as { document: unknown; window: unknown };
  g.document = {
    createElement: () => ({ getContext: () => ({ measureText: () => ({ width: 0 }), font: "" }), style: {} }),
  };
  g.window = { addEventListener: () => {}, removeEventListener: () => {} };
});

/** Build the controller with stub deps and return the captured keydown handler
 *  plus the spied object-layer deps. */
async function setup(over: Record<string, unknown> = {}) {
  const { createSelectionController } = await import("./selectionController");
  const container: FakeContainer = {
    tabIndex: 0,
    style: {},
    handlers: {},
    addEventListener(type, fn) { this.handlers[type] = fn; },
    removeEventListener() {},
  };
  const cycleObjectFocus = vi.fn(() => true);
  const nudgeSelectedObject = vi.fn(() => true);
  const onTab = vi.fn(() => false);
  const deps = {
    container,
    getTree: () => ({ pages: [] }),
    getDoc: () => ({ section: {}, blocks: [] }),
    getSelection: () => null,
    setSelection: () => {},
    getCellSelection: () => null,
    setCellSelection: () => {},
    clientToPage: () => null,
    getGridSpacing: () => 0,
    isSnapToGrid: () => false,
    focusProxy: () => {},
    onDeleteSelection: () => {},
    getStory: () => null,
    setStory: () => {},
    decorationAt: () => null,
    onDecorationClick: () => false,
    selectObject: () => {},
    hasSelectedObject: () => false,
    cycleObjectFocus,
    nudgeSelectedObject,
    deleteSelectedObject: () => {},
    startColumnDrag: () => {},
    setColumnGuide: () => {},
    startRowDrag: () => {},
    setRowGuide: () => {},
    applyObjectMove: () => {},
    onTab,
    jumpToBlock: () => {},
    onAnchorJump: () => {},
    onSdtPress: () => false,
    getActiveShapeText: () => null,
    getShapeScope: () => null,
    getShapeTextParagraphs: () => [],
    shapeTextRect: () => null,
    shapeHasText: () => false,
    enterShapeText: () => {},
    exitShapeText: () => {},
    enterSelectedShapeText: () => false,
    startShapeText: () => false,
    ...over,
  };
  createSelectionController(deps as never);
  const keydown = container.handlers["keydown"]!;
  return { keydown, cycleObjectFocus, nudgeSelectedObject, onTab };
}

const key = (k: string, mods: Partial<{ shiftKey: boolean; ctrlKey: boolean; altKey: boolean; metaKey: boolean }> = {}) => {
  const ev = { key: k, shiftKey: false, ctrlKey: false, altKey: false, metaKey: false, ...mods, preventDefault: vi.fn() };
  return ev;
};

describe("object keyboard routing (E2/E3)", () => {
  it("F6 enters the object cycle forward; Shift+F6 backward", async () => {
    const { keydown, cycleObjectFocus } = await setup();
    const f6 = key("F6");
    keydown(f6);
    expect(cycleObjectFocus).toHaveBeenCalledWith(false);
    expect(f6.preventDefault).toHaveBeenCalled();

    const sf6 = key("F6", { shiftKey: true });
    keydown(sf6);
    expect(cycleObjectFocus).toHaveBeenLastCalledWith(true);
  });

  it("F6 does not fire while editing a shape's text box", async () => {
    const { keydown, cycleObjectFocus } = await setup({ getActiveShapeText: () => "shp1" });
    const f6 = key("F6");
    keydown(f6);
    expect(cycleObjectFocus).not.toHaveBeenCalled();
    expect(f6.preventDefault).not.toHaveBeenCalled();
  });

  it("Tab / Shift+Tab cycle objects when one is selected (not cell nav)", async () => {
    const { keydown, cycleObjectFocus, onTab } = await setup({ hasSelectedObject: () => true });
    const tab = key("Tab");
    keydown(tab);
    expect(cycleObjectFocus).toHaveBeenCalledWith(false);
    expect(onTab).not.toHaveBeenCalled();
    expect(tab.preventDefault).toHaveBeenCalled();

    keydown(key("Tab", { shiftKey: true }));
    expect(cycleObjectFocus).toHaveBeenLastCalledWith(true);
  });

  it("Tab falls through to cell navigation when no object is selected", async () => {
    const { keydown, cycleObjectFocus, onTab } = await setup({ hasSelectedObject: () => false });
    keydown(key("Tab"));
    expect(cycleObjectFocus).not.toHaveBeenCalled();
    expect(onTab).toHaveBeenCalledWith(false);
  });

  it("arrows nudge a selected object by 1px, 10px with Shift", async () => {
    const { keydown, nudgeSelectedObject } = await setup({ hasSelectedObject: () => true });
    keydown(key("ArrowRight"));
    expect(nudgeSelectedObject).toHaveBeenCalledWith(1, 0);
    keydown(key("ArrowLeft"));
    expect(nudgeSelectedObject).toHaveBeenLastCalledWith(-1, 0);
    keydown(key("ArrowUp", { shiftKey: true }));
    expect(nudgeSelectedObject).toHaveBeenLastCalledWith(0, -10);
    keydown(key("ArrowDown", { shiftKey: true }));
    expect(nudgeSelectedObject).toHaveBeenLastCalledWith(0, 10);
  });

  it("arrows do not nudge when no object is selected", async () => {
    const { keydown, nudgeSelectedObject } = await setup({ hasSelectedObject: () => false });
    keydown(key("ArrowRight"));
    expect(nudgeSelectedObject).not.toHaveBeenCalled();
  });

  it("Ctrl+Arrow does not nudge an object (leaves word navigation alone)", async () => {
    const { keydown, nudgeSelectedObject } = await setup({ hasSelectedObject: () => true });
    keydown(key("ArrowRight", { ctrlKey: true }));
    expect(nudgeSelectedObject).not.toHaveBeenCalled();
  });
});

// Ctrl+A select-all must work under non-Latin layouts, where the physical A key
// emits a non-Latin character (e.g. 'ф' on a Ukrainian layout). The handler matches
// via the physical code, so the Ctrl+A branch is entered and consumes the event
// (calls preventDefault) regardless of the produced character. (setSelection needs a
// real layout tree to fire, which the stub tree doesn't provide, so we assert the
// branch was reached via preventDefault — the observable that was broken before.)
describe("Ctrl+A select-all under non-Latin layouts", () => {
  const kev = (
    k: string,
    code: string,
    mods: Partial<{ shiftKey: boolean; ctrlKey: boolean; altKey: boolean; metaKey: boolean }> = {},
  ) => ({ key: k, code, shiftKey: false, ctrlKey: false, altKey: false, metaKey: false, ...mods, preventDefault: vi.fn() });

  it("enters the select-all branch when the physical A emits 'ф' (Cyrillic)", async () => {
    const { keydown } = await setup();
    const e = kev("ф", "KeyA", { ctrlKey: true });
    keydown(e);
    expect(e.preventDefault).toHaveBeenCalled();
  });

  it("still works for a plain Latin Ctrl+A", async () => {
    const { keydown } = await setup();
    const e = kev("a", "KeyA", { ctrlKey: true });
    keydown(e);
    expect(e.preventDefault).toHaveBeenCalled();
  });

  it("ignores a Cyrillic Ctrl-chord on a different physical key", async () => {
    const { keydown } = await setup();
    const e = kev("ф", "KeyS", { ctrlKey: true }); // not the physical A position
    keydown(e);
    expect(e.preventDefault).not.toHaveBeenCalled();
  });

  // Windows reports AltGr as Ctrl+Alt, so without an altKey guard the physical-code
  // match above would read AltGr+A ('ą' on a Polish layout) as select-all and eat
  // the character.
  it("leaves AltGr+A alone so 'ą' can be typed", async () => {
    const { keydown } = await setup();
    const e = kev("ą", "KeyA", { ctrlKey: true, altKey: true });
    keydown(e);
    expect(e.preventDefault).not.toHaveBeenCalled();
  });
});
