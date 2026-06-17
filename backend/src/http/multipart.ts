/** Minimal multipart/form-data parser (binary-safe) — enough for a file part plus
 *  small JSON fields. Returns each part's name → { data, filename? }. */
export function parseMultipart(contentType: string, body: Buffer): Record<string, { data: Buffer; filename?: string }> {
  const m = /boundary=(?:"([^"]+)"|([^;]+))/i.exec(contentType);
  if (!m) return {};
  const boundary = Buffer.from("--" + (m[1] ?? m[2] ?? "").trim());
  const out: Record<string, { data: Buffer; filename?: string }> = {};
  let idx = body.indexOf(boundary);
  while (idx >= 0) {
    const start = idx + boundary.length;
    if (body[start] === 0x2d && body[start + 1] === 0x2d) break; // closing "--"
    const headerEnd = body.indexOf("\r\n\r\n", start);
    if (headerEnd < 0) break;
    const headers = body.toString("utf8", start, headerEnd);
    const next = body.indexOf(boundary, headerEnd + 4);
    if (next < 0) break;
    const data = body.subarray(headerEnd + 4, next - 2); // strip trailing CRLF
    const name = /name="([^"]+)"/i.exec(headers)?.[1];
    const filename = /filename="([^"]*)"/i.exec(headers)?.[1];
    if (name) out[name] = filename !== undefined ? { data, filename } : { data };
    idx = next;
  }
  return out;
}
