// A header/footer part (w:hdr / w:ftr) — same block content as the body.

import type { Block } from "@cw/shared";
import { blockXml, type PartCtx } from "./documentXml";
import { el, WML_NS, XML_DECL } from "./xmlWrite";

export function headerFooterXml(blocks: Block[], kind: "header" | "footer", ctx: PartCtx): string {
  const root = kind === "header" ? "w:hdr" : "w:ftr";
  const body = blocks.length > 0 ? blocks.map((b) => blockXml(b, ctx)).join("") : el("w:p");
  return XML_DECL + el(root, WML_NS, body);
}
