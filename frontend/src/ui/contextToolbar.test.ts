import { describe, expect, it } from "vitest";
import { pickActive, type ContextToolbar } from "./contextToolbar";
import type { AnchorRect } from "./floatingBarPosition";

const VP = { width: 1000, height: 800 };
const RECT: AnchorRect = { left: 100, top: 100, width: 50, height: 20 };

/** A fake toolbar whose `resolve()` returns a fixed anchor (or null). */
function fake(id: string, priority: number, anchor: AnchorRect | null): ContextToolbar {
  return {
    id,
    priority,
    resolve: () => anchor,
    show: () => {},
    hide: () => {},
    destroy: () => {},
  };
}

describe("pickActive", () => {
  it("returns null when no toolbar's context is active", () => {
    const picked = pickActive([fake("a", 10, null), fake("b", 20, null)], VP);
    expect(picked).toBeNull();
  });

  it("picks the highest-priority active toolbar regardless of registration order", () => {
    const low = fake("low", 10, RECT);
    const high = fake("high", 30, RECT);
    const mid = fake("mid", 20, RECT);
    expect(pickActive([low, high, mid], VP)?.toolbar.id).toBe("high");
    expect(pickActive([high, low, mid], VP)?.toolbar.id).toBe("high");
  });

  it("skips a higher-priority toolbar that is inactive and falls to the next", () => {
    const picked = pickActive([fake("img", 30, null), fake("text", 20, RECT)], VP);
    expect(picked?.toolbar.id).toBe("text");
  });

  it("returns the resolved anchor alongside the toolbar", () => {
    const picked = pickActive([fake("a", 10, RECT)], VP);
    expect(picked?.anchor).toEqual(RECT);
  });

  it("skips a toolbar whose anchor is scrolled out of view", () => {
    const offscreen: AnchorRect = { left: 100, top: -500, width: 50, height: 20 };
    const picked = pickActive([fake("img", 30, offscreen), fake("text", 20, RECT)], VP);
    expect(picked?.toolbar.id).toBe("text");
  });

  it("keeps registration order for equal priorities (stable)", () => {
    const first = fake("first", 15, RECT);
    const second = fake("second", 15, RECT);
    expect(pickActive([first, second], VP)?.toolbar.id).toBe("first");
  });
});
