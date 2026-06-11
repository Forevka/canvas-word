// Image bytes → data: URL. The builder never mints blob: URLs — a blob: URL is
// tied to one browsing context and means nothing in Node, while a data: URL is
// self-contained, so built documents are portable across the editor, the
// browser export worker, and server-side Node export.

/** Base64-encode bytes without Buffer (btoa is global in browsers and Node 16+).
 *  Chunked so large images don't blow the call-stack via String.fromCharCode. */
export function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  const CHUNK = 0x8000;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
  }
  return btoa(binary);
}

export function bytesToDataUrl(bytes: Uint8Array | ArrayBuffer, mime: string): string {
  const u8 = bytes instanceof ArrayBuffer ? new Uint8Array(bytes) : bytes;
  return `data:${mime};base64,${bytesToBase64(u8)}`;
}
