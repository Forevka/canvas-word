import { describe, expect, it } from "vitest";
import { configureIds, createIdGenerator, currentSiteId, freshId } from "./ids";

describe("id generation", () => {
  it("mints siteId-namespaced, monotonic, base36 ids", () => {
    const g = createIdGenerator("a3f9c1");
    expect(g.next()).toBe("a3f9c1-0");
    expect(g.next()).toBe("a3f9c1-1");
    expect(g.count()).toBe(2);
  });

  it("resumes from a starting counter", () => {
    const g = createIdGenerator("zz", 40);
    expect(g.next()).toBe("zz-14"); // 40 in base36
  });

  it("two sites never collide", () => {
    const a = createIdGenerator("siteA");
    const b = createIdGenerator("siteB");
    const ids = new Set<string>();
    for (let i = 0; i < 1000; i++) {
      ids.add(a.next());
      ids.add(b.next());
    }
    expect(ids.size).toBe(2000);
  });

  it("the default generator is reconfigurable", () => {
    configureIds("sess1");
    expect(currentSiteId()).toBe("sess1");
    expect(freshId()).toBe("sess1-0");
    expect(freshId()).toBe("sess1-1");
    configureIds("sess2", 5);
    expect(freshId()).toBe("sess2-5");
  });
});
