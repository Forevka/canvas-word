// Regression guard for the full-page background-image bug: PARTY INVITATION.docx
// renders on ONE page in Word (a decorative wp:anchor + wp:wrapNone + behindDoc
// image sits behind the text), but the engine used to treat it as a flow block
// that consumed a whole page and pushed the text onto page 2.
//
// Gated on the file being present at the frontend root (it is large and not
// committed), so CI without it still passes — same pattern as export/e2e.test.ts.
import "./test-canvas-setup";

import { existsSync, readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { runImport } from "../import/docx/pipeline";
import { createLayoutEngine } from "./engine";

const FILE = fileURLToPath(new URL("../../PARTY INVITATION.docx", import.meta.url));
const maybe = existsSync(FILE) ? describe : describe.skip;

maybe("PARTY INVITATION.docx (full-page background)", () => {
  it("lays out on a single page with the background behind the text", () => {
    const { doc } = runImport(new Uint8Array(readFileSync(FILE)));
    // The decorative background imports as a behind-text anchored image.
    const bg = doc.blocks.find((b) => b.kind === "image" && b.anchor?.behind);
    expect(bg).toBeDefined();

    const tree = createLayoutEngine().layout(doc);
    expect(tree.pages).toHaveLength(1);

    // The background is painted first (under the text) and is the page's first block.
    const first = tree.pages[0]!.blocks[0]!;
    expect(first.image?.behind).toBe(true);
    // The title text is present on the same page, not pushed below a page of image.
    const hasText = tree.pages[0]!.blocks.some((b) => b.lines.some((l) => l.fragments.length > 0));
    expect(hasText).toBe(true);
  });
});
