// Generic OOXML field-instruction parsing. Pure and layout-free, so import,
// export, and the editor command layer all share it.
//
// A "field" in OOXML is a region delimited by w:fldChar begin/separate/end runs
// with a w:instrText instruction (e.g. ` TOC \o "1-3" \h `, ` PAGEREF _Toc1 \h `,
// ` MYCHART "sales-2026" `). Word ships a fixed set of BUILT-IN field types it
// knows how to render; everything else is a developer-defined CUSTOM field whose
// result the host computes (see Document.fields + the editor's resolveField hook).
// This module only classifies an instruction and splits out its name + arguments;
// TOC has its own richer parser in ./toc.

/** A field instruction split into its keyword and arguments. */
export interface ParsedFieldInstruction {
  /** Field keyword, uppercased (e.g. "TOC", "MYCHART"). "" when the instruction
   *  is blank. */
  name: string;
  /** Whitespace/quote-delimited tokens after the keyword (surrounding double
   *  quotes stripped). Switches like `\o` appear here verbatim. */
  args: string[];
  /** The instruction exactly as given (untrimmed). */
  raw: string;
}

/** Field keywords Word renders itself — NOT treated as host-resolvable custom
 *  fields. TOC is built-in here too (the editor handles it on its own tocEntry
 *  path), so it is never captured as a generic custom field. */
const BUILTIN_FIELDS = new Set([
  "TOC", "PAGE", "NUMPAGES", "PAGEREF", "HYPERLINK", "REF", "SEQ", "STYLEREF",
  "DATE", "TIME", "CREATEDATE", "SAVEDATE", "PRINTDATE", "EDITTIME", "FILENAME",
  "FILESIZE", "AUTHOR", "LASTSAVEDBY", "TITLE", "SUBJECT", "KEYWORDS", "COMMENTS",
  "NUMWORDS", "NUMCHARS", "TOA", "INDEX", "XE", "TC", "TA", "RD",
  "FORMTEXT", "FORMCHECKBOX", "FORMDROPDOWN", "SYMBOL", "QUOTE", "IF",
  "DOCPROPERTY", "DOCVARIABLE", "MERGEFIELD", "ADDRESSBLOCK", "GREETINGLINE",
  "FILLIN", "ASK", "SET", "GOTOBUTTON", "MACROBUTTON", "EQ", "LISTNUM",
  "BIBLIOGRAPHY", "CITATION", "SECTION", "SECTIONPAGES", "AUTONUM", "AUTONUMLGL",
  "AUTONUMOUT", "INCLUDETEXT", "INCLUDEPICTURE", "LINK", "NOTEREF",
]);

/** Is `name` one of Word's built-in field types (so NOT a custom field)? */
export function isBuiltinField(name: string): boolean {
  return BUILTIN_FIELDS.has(name.trim().toUpperCase());
}

/** Tokenize a field instruction, honoring double-quoted arguments. */
function tokenizeField(instr: string): string[] {
  const out: string[] = [];
  const re = /"([^"]*)"|(\S+)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(instr)) !== null) out.push(m[1] !== undefined ? m[1] : m[2]!);
  return out;
}

/** Parse a field instruction into its keyword + arguments. */
export function parseFieldInstruction(instr: string): ParsedFieldInstruction {
  const tokens = tokenizeField(instr);
  return { name: tokens.length > 0 ? tokens[0]!.toUpperCase() : "", args: tokens.slice(1), raw: instr };
}

/** Is this instruction a host-resolvable custom field (has a keyword and isn't a
 *  Word built-in)? Drives import capture + the editor's "Update Field" menu. */
export function isCustomFieldInstruction(instr: string): boolean {
  const name = parseFieldInstruction(instr).name;
  return name.length > 0 && !isBuiltinField(name);
}
