// Loads the bundled metric-clone fonts as FontFaces on the editor's main thread,
// registered under their clone names (Carlito, Arimo, …). charStyleToFont maps
// every document family to a clone name, so once these are loaded the editor
// measures & paints with the exact fonts the exporters embed — identical layout
// in the editor, the browser export, and the Node backend. Call (and await)
// before the first layout.

import { CLONE_FAMILIES, FONT_STYLES } from "../../fonts/clones";

const STYLE_ATTRS: Record<string, { weight: string; style: string }> = {
  Regular: { weight: "400", style: "normal" },
  Bold: { weight: "700", style: "normal" },
  Italic: { weight: "400", style: "italic" },
  BoldItalic: { weight: "700", style: "italic" },
};

let started: Promise<void> | null = null;

/** Fetch + register the bundled fonts (idempotent). `onProgress(loaded, total)`
 *  reports streamed bytes for a smooth loading bar; on a memoized second call it
 *  fires once at 100% so a late-mounting editor's bar still completes. */
export function loadEditorFonts(onProgress?: (loaded: number, total: number) => void): Promise<void> {
  // Already loading/loaded (e.g. a second WordCanvas instance): don't re-fetch,
  // but still complete the caller's bar once the shared load settles.
  if (started) {
    if (onProgress) void started.then(() => onProgress(1, 1));
    return started;
  }
  if (typeof document === "undefined" || !("fonts" in document)) return Promise.resolve();
  started = (async () => {
    const specs = CLONE_FAMILIES.flatMap((fam) =>
      FONT_STYLES.map((style) => ({ fam, style, url: new URL(`./fonts/${fam}-${style}.ttf`, import.meta.url) })),
    );
    // Kick off every fetch, then read content-length up front so `total` (and thus
    // the percentage) is known before any body finishes streaming.
    const responses = await Promise.all(specs.map((s) => fetch(s.url)));
    const total = responses.reduce((sum, r) => sum + (Number(r.headers.get("content-length")) || 0), 0);
    let loaded = 0;
    const buffers = await Promise.all(
      responses.map(async (r) => {
        // No streamable body or unknown size: fall back to a single chunk so the
        // bar still advances (just less smoothly) by the buffer's byte length.
        if (!r.body || total === 0) {
          const buf = await r.arrayBuffer();
          loaded += buf.byteLength;
          onProgress?.(loaded, total || loaded);
          return buf;
        }
        const reader = r.body.getReader();
        const chunks: Uint8Array[] = [];
        let size = 0;
        for (;;) {
          const { done, value } = await reader.read();
          if (done) break;
          chunks.push(value);
          size += value.byteLength;
          loaded += value.byteLength;
          onProgress?.(loaded, total);
        }
        const buf = new Uint8Array(size);
        let off = 0;
        for (const c of chunks) {
          buf.set(c, off);
          off += c.byteLength;
        }
        return buf.buffer;
      }),
    );
    const faces = specs.map((s, i) => {
      const a = STYLE_ATTRS[s.style]!;
      return new FontFace(s.fam, buffers[i]!, { weight: a.weight, style: a.style });
    });
    await Promise.all(faces.map((f) => f.load()));
    for (const f of faces) (document as Document & { fonts: FontFaceSet }).fonts.add(f);
  })();
  return started;
}
