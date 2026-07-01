// Regression: a CROPPED image whose bytes can't be decoded must not leak its clip.
// The crop path clips to the crop window, then draws the image; if the draw throws
// (e.g. an SVG or otherwise-undecodable image), the clip used to stay on the graphics
// stack and blanked the rest of the page (body text, header, footer rendered
// white-on-white). The fix restores the clip even on a decode error, so save/restore
// stay balanced and only a placeholder box is drawn.
import { describe, expect, it } from "vitest";
import type { PlacedBlock } from "../../layout/layoutTree";
import { paintBlock, type PaintCtx } from "./paintBlock";

/** A minimal pdfkit stand-in that records the graphics ops we care about and throws
 *  on `.image()` (simulating an undecodable image). Only the methods paintBlock's
 *  image path touches are implemented. */
function makeRecordingDoc(): { doc: PaintCtx["doc"]; ops: string[] } {
  const ops: string[] = [];
  const chain = {
    clip: () => (ops.push("clip"), doc),
    fill: () => (ops.push("fill"), doc),
  };
  const doc = {
    save: () => (ops.push("save"), doc),
    restore: () => (ops.push("restore"), doc),
    rect: () => (ops.push("rect"), chain),
    image: () => {
      ops.push("image");
      throw new Error("image-format-unsupported");
    },
  } as unknown as PaintCtx["doc"];
  return { doc, ops };
}

/** A placed cropped-image block (garbage bytes decode-fail through ctx.image). */
const croppedImageBlock = (): PlacedBlock =>
  ({
    blockId: "img0",
    x: 100,
    y: 100,
    image: {
      src: "data:image/svg+xml,<svg/>",
      width: 150,
      height: 110,
      crop: { left: 0.25, top: 0.2, right: 0.25, bottom: 0.2 },
    },
    lines: [],
  }) as unknown as PlacedBlock;

describe("PDF image crop — clip never leaks on a decode error", () => {
  it("restores the clip (balanced save/restore) and warns when a cropped image fails to decode", () => {
    const { doc, ops } = makeRecordingDoc();
    let warned: string | undefined;
    const ctx: PaintCtx = {
      doc,
      font: () => "Arimo-Regular.ttf",
      image: () => Uint8Array.from([1, 2, 3, 4]), // not a valid PNG/JPEG → drawImage throws
      warn: (code) => {
        warned = code;
      },
    };

    expect(() => paintBlock(ctx, croppedImageBlock())).not.toThrow();

    // The failed decode was reported…
    expect(warned).toBe("image-format-unsupported");
    // …the image draw was attempted…
    expect(ops).toContain("image");
    // …and every save was matched by a restore (the clip did NOT leak).
    expect(ops.filter((o) => o === "save").length).toBe(ops.filter((o) => o === "restore").length);
    // The restore happens AFTER the throwing image draw (finally), before the placeholder.
    expect(ops.indexOf("restore")).toBeGreaterThan(ops.indexOf("image"));
    // A placeholder box is still drawn (rect + fill) after the clip is restored.
    expect(ops).toContain("fill");
  });
});
