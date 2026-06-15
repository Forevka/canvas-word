// Parse a fragment of body OOXML (what a custom-field resolver returns) into model
// blocks. The fragment is a sequence of w:p / w:tbl (or a whole w:document); we
// wrap it in a minimal docx and run the FULL import pipeline, so styling, tables,
// lists and the rest decode exactly as they would from a real file.
//
// v1 limitation: a fragment referencing external media/relationships (images by
// r:id, hyperlinks via rels) won't resolve those parts — it has none. Inline text,
// formatting, tables and lists work.

import { strToU8, zipSync } from "fflate";
import type { Block } from "@cw/shared";
import { runImport } from "./pipeline";

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** Wrap a body OOXML fragment into a full document.xml part. */
function wrapDocumentXml(ooxml: string): string {
  const t = ooxml.trim();
  if (/<w:document[\s>]/.test(t)) return t.startsWith("<?xml") ? t : `${XML_DECL}\n${t}`;
  return `${XML_DECL}\n<w:document xmlns:w="${WML_NS}"><w:body>${t}</w:body></w:document>`;
}

/** Parse resolver OOXML into model blocks (empty array if it has no blocks). */
export function parseOoxmlFragment(ooxml: string): Block[] {
  const bytes = zipSync({
    "[Content_Types].xml": strToU8(CONTENT_TYPES),
    "word/document.xml": strToU8(wrapDocumentXml(ooxml)),
  });
  return runImport(bytes).doc.blocks;
}
