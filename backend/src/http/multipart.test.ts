import { describe, expect, it } from "vitest";
import { parseMultipart } from "./multipart";

const CRLF = "\r\n";

/** Build a multipart/form-data body Buffer from parts. Each part may carry a
 *  filename (file part) or none (plain field). Data may be a string or Buffer so
 *  we can exercise binary-safe handling. */
function buildBody(
  boundary: string,
  parts: Array<{ name: string; filename?: string; data: string | Buffer }>,
): Buffer {
  const chunks: Buffer[] = [];
  for (const p of parts) {
    let header = `--${boundary}${CRLF}Content-Disposition: form-data; name="${p.name}"`;
    if (p.filename !== undefined) header += `; filename="${p.filename}"`;
    header += CRLF + CRLF;
    chunks.push(Buffer.from(header, "utf8"));
    chunks.push(typeof p.data === "string" ? Buffer.from(p.data, "utf8") : p.data);
    chunks.push(Buffer.from(CRLF, "utf8"));
  }
  chunks.push(Buffer.from(`--${boundary}--${CRLF}`, "utf8"));
  return Buffer.concat(chunks);
}

describe("parseMultipart", () => {
  it("returns {} when the content-type has no boundary", () => {
    expect(parseMultipart("multipart/form-data", Buffer.from("anything"))).toEqual({});
  });

  it("returns {} for a totally unrelated content-type", () => {
    expect(parseMultipart("application/json", Buffer.from("{}"))).toEqual({});
  });

  it("parses a single file part with its filename", () => {
    const body = buildBody("X", [{ name: "file", filename: "doc.docx", data: "hello world" }]);
    const out = parseMultipart('multipart/form-data; boundary=X', body);
    expect(Object.keys(out)).toEqual(["file"]);
    expect(out.file!.filename).toBe("doc.docx");
    expect(out.file!.data.toString("utf8")).toBe("hello world");
  });

  it("parses a quoted boundary value", () => {
    const body = buildBody("Bound-42", [{ name: "file", filename: "a.bin", data: "data" }]);
    const out = parseMultipart('multipart/form-data; boundary="Bound-42"', body);
    expect(out.file!.data.toString("utf8")).toBe("data");
  });

  it("parses a JSON field part with no filename (no filename key on the result)", () => {
    const body = buildBody("X", [{ name: "meta", data: '{"title":"hi"}' }]);
    const out = parseMultipart("multipart/form-data; boundary=X", body);
    expect(out.meta!.data.toString("utf8")).toBe('{"title":"hi"}');
    expect("filename" in out.meta!).toBe(false);
  });

  it("includes an empty filename when filename=\"\" is present", () => {
    const body = buildBody("X", [{ name: "f", filename: "", data: "x" }]);
    const out = parseMultipart("multipart/form-data; boundary=X", body);
    expect(out.f!.filename).toBe("");
  });

  it("parses multiple parts (a field plus a file)", () => {
    const body = buildBody("X", [
      { name: "meta", data: '{"k":1}' },
      { name: "file", filename: "f.docx", data: "BODYBYTES" },
    ]);
    const out = parseMultipart("multipart/form-data; boundary=X", body);
    expect(Object.keys(out).sort()).toEqual(["file", "meta"]);
    expect(out.meta!.data.toString("utf8")).toBe('{"k":1}');
    expect(out.file!.filename).toBe("f.docx");
    expect(out.file!.data.toString("utf8")).toBe("BODYBYTES");
  });

  it("preserves binary data byte-for-byte (including bytes that look like CR/LF/dashes)", () => {
    const bin = Buffer.from([0x00, 0x0d, 0x0a, 0x2d, 0x2d, 0xff, 0x10, 0x00]);
    const body = buildBody("X", [{ name: "file", filename: "b.bin", data: bin }]);
    const out = parseMultipart("multipart/form-data; boundary=X", body);
    expect([...out.file!.data]).toEqual([...bin]);
  });

  it("stops at the closing boundary and ignores trailing bytes", () => {
    const body = buildBody("X", [{ name: "file", filename: "f", data: "abc" }]);
    const withTrailer = Buffer.concat([body, Buffer.from("garbage after closing boundary")]);
    const out = parseMultipart("multipart/form-data; boundary=X", withTrailer);
    expect(Object.keys(out)).toEqual(["file"]);
    expect(out.file!.data.toString("utf8")).toBe("abc");
  });

  it("ignores a part whose Content-Disposition has no name", () => {
    // Manually craft a nameless part; it should be skipped entirely.
    const body = Buffer.concat([
      Buffer.from(`--X${CRLF}Content-Disposition: form-data${CRLF}${CRLF}`, "utf8"),
      Buffer.from("orphan"),
      Buffer.from(CRLF, "utf8"),
      Buffer.from(`--X--${CRLF}`, "utf8"),
    ]);
    const out = parseMultipart("multipart/form-data; boundary=X", body);
    expect(out).toEqual({});
  });
});
