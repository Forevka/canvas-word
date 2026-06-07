import { describe, expect, it } from "vitest";
import { strFromU8 } from "fflate";
import { openArchive } from "./zip";
import { ImportError } from "./types";
import { makeDocx } from "./fixture";

const errCode = (fn: () => unknown): string => {
  try {
    fn();
  } catch (e) {
    if (e instanceof ImportError) return e.code;
    throw e;
  }
  throw new Error("expected ImportError");
};

describe("openArchive", () => {
  it("reads parts from a real zip", () => {
    const archive = openArchive(makeDocx({ "word/document.xml": "<w:document/>" }));
    expect(archive.names()).toContain("word/document.xml");
    expect(archive.text("word/document.xml")).toBe("<w:document/>");
    expect(archive.text("nope.xml")).toBeUndefined();
  });

  it("decodes UTF-8 part text", () => {
    const archive = openArchive(makeDocx({ "a.xml": "<t>héllo — 你好</t>" }));
    expect(archive.text("a.xml")).toBe("<t>héllo — 你好</t>");
    expect(strFromU8(archive.part("a.xml")!)).toBe("<t>héllo — 你好</t>");
  });

  it("rejects non-zip bytes with NOT_ZIP", () => {
    expect(errCode(() => openArchive(new Uint8Array([1, 2, 3, 4, 5])))).toBe("NOT_ZIP");
  });

  it("detects password-protected (OLE) files as ENCRYPTED", () => {
    const ole = new Uint8Array(64);
    ole.set([0xd0, 0xcf, 0x11, 0xe0]);
    expect(errCode(() => openArchive(ole))).toBe("ENCRYPTED");
  });
});
