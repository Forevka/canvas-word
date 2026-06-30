// Test-only helpers: synthesize minimal-but-real .docx archives in memory.
// Real zips via fflate — the same bytes Word would produce structurally, so the
// whole pipeline (magic sniffing, unzip, content types, parse) is exercised.

import { strToU8, zipSync } from "fflate";

export const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

export function documentXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
}

export function makeDocx(parts: Record<string, string | Uint8Array>): Uint8Array {
  return zipSync(
    Object.fromEntries(
      Object.entries(parts).map(([name, content]) => [
        name,
        typeof content === "string" ? strToU8(content) : content,
      ]),
    ),
  );
}

/** Valid 1×1 transparent PNG. */
export const PNG_1PX = Uint8Array.from(
  atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="),
  (c) => c.charCodeAt(0),
);

export function headerPartXml(content: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${content}</w:hdr>`;
}

export function footerPartXml(content: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${content}</w:ftr>`;
}

export const REL_TYPES = {
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  header: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
  footer: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
  styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  numbering: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
  hyperlink: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
  footnotes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
  endnotes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
  settings: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
} as const;

export function relsXml(rels: Array<{ id: string; type: string; target: string; external?: boolean }>): string {
  const entries = rels
    .map(
      (r) =>
        `<Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}"${r.external ? ' TargetMode="External"' : ""}/>`,
    )
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;
}

/** Minimal DrawingML inline/anchored image (namespaces left undeclared — txml
 *  matches qualified names directly, same as Word output parsing). */
export function drawingXml(relId: string, cxEmu: number, cyEmu: number, anchored = false): string {
  const tag = anchored ? "wp:anchor" : "wp:inline";
  return (
    `<w:drawing><${tag}><wp:extent cx="${cxEmu}" cy="${cyEmu}"/>` +
    `<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="${relId}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>` +
    `</${tag}></w:drawing>`
  );
}

/** Standard two-part docx around the given w:body content. */
export function simpleDocx(body: string): Uint8Array {
  return makeDocx({
    "[Content_Types].xml": CONTENT_TYPES_XML,
    "word/document.xml": documentXml(body),
  });
}

export function stylesPartXml(styles: string, docDefaults = ""): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${docDefaults}${styles}</w:styles>`;
}

/** Office-default-ish theme: minor=Calibri, major=Calibri Light, accent1=4472C4. */
export const THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`;

/** docx with styles.xml / theme1.xml wired through a real rels part. */
export function styledDocx(body: string, stylesXml?: string, themeXml: string = THEME_XML): Uint8Array {
  const parts: Record<string, string> = {
    "[Content_Types].xml": CONTENT_TYPES_XML,
    "word/document.xml": documentXml(body),
  };
  const rels: string[] = [];
  if (stylesXml !== undefined) {
    parts["word/styles.xml"] = stylesXml;
    rels.push(
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
    );
  }
  parts["word/theme/theme1.xml"] = themeXml;
  rels.push(
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`,
  );
  parts["word/_rels/document.xml.rels"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`;
  return makeDocx(parts);
}
