// Text tokenizer for intra-paragraph diffing. Splits a string into tokens whose
// UTF-16 [start, end) offsets are LOAD-BEARING: diffDocuments turns token hunks
// into suggestion anchors, which are (blockId, UTF-16 offset) positions into the
// live document. Word granularity keeps the redline readable (whole words move
// as a unit); char granularity is available for a finer diff.

export interface Token {
  text: string;
  /** UTF-16 offset of the token's first code unit within the source string. */
  start: number;
  /** UTF-16 offset one past the token's last code unit. */
  end: number;
}

export type Granularity = "word" | "char";

// A word run (\w+), a whitespace run (\s+), or any single other character
// (punctuation/symbol). Unicode-aware; each match is one token.
const WORD_RE = /\p{L}[\p{L}\p{M}\p{N}_']*|\s+|[^\p{L}\p{M}\p{N}\s]/gu;

/** Split `text` into tokens carrying their UTF-16 offsets. */
export function tokenize(text: string, granularity: Granularity = "word"): Token[] {
  return granularity === "char" ? tokenizeChars(text) : tokenizeWords(text);
}

function tokenizeWords(text: string): Token[] {
  const out: Token[] = [];
  WORD_RE.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = WORD_RE.exec(text)) !== null) {
    out.push({ text: match[0], start: match.index, end: match.index + match[0].length });
    if (match[0].length === 0) WORD_RE.lastIndex++; // defensive: never loop on a zero-width match
  }
  return out;
}

/** One token per Unicode code point, but with UTF-16 offsets (so a surrogate
 *  pair stays one token spanning two code units). */
function tokenizeChars(text: string): Token[] {
  const out: Token[] = [];
  let start = 0;
  for (const ch of text) {
    out.push({ text: ch, start, end: start + ch.length });
    start += ch.length;
  }
  return out;
}
