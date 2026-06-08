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

export function loadEditorFonts(): Promise<void> {
  if (started) return started;
  if (typeof document === "undefined" || !("fonts" in document)) return Promise.resolve();
  started = (async () => {
    const faces = await Promise.all(
      CLONE_FAMILIES.flatMap((fam) =>
        FONT_STYLES.map(async (style) => {
          const url = new URL(`./fonts/${fam}-${style}.ttf`, import.meta.url);
          const buf = await (await fetch(url)).arrayBuffer();
          const a = STYLE_ATTRS[style]!;
          return new FontFace(fam, buf, { weight: a.weight, style: a.style });
        }),
      ),
    );
    await Promise.all(faces.map((f) => f.load()));
    for (const f of faces) (document as Document & { fonts: FontFaceSet }).fonts.add(f);
  })();
  return started;
}
