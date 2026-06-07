// OPC container access: a .docx is a zip of "parts". fflate does the inflation;
// this module adds magic-byte sniffing (the difference between "not a docx" and
// "password-protected docx" is a real UX distinction) and UTF-8 part reads.

import { strFromU8, unzipSync } from "fflate";
import { ImportError } from "./types";

const ZIP_MAGIC = [0x50, 0x4b]; // "PK"
/** OLE compound file — what Word writes for password-protected documents. */
const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0];

const startsWith = (bytes: Uint8Array, magic: number[]): boolean =>
  magic.every((b, i) => bytes[i] === b);

export interface Archive {
  /** Raw bytes of a part by exact name, e.g. "word/document.xml". */
  part(name: string): Uint8Array | undefined;
  /** Part decoded as UTF-8 text. */
  text(name: string): string | undefined;
  names(): string[];
}

export function openArchive(bytes: Uint8Array): Archive {
  if (startsWith(bytes, OLE_MAGIC)) {
    throw new ImportError("ENCRYPTED", "Password-protected documents are not supported.");
  }
  if (!startsWith(bytes, ZIP_MAGIC)) {
    throw new ImportError("NOT_ZIP", "Not a .docx file (zip container signature missing).");
  }
  // Media is the bulk of a rich docx (often megabytes of JPEGs) but only the
  // referenced parts are ever read — inflate everything EXCEPT media up front,
  // and pull media parts on demand.
  const isMedia = (name: string): boolean => /^word\/media\//.test(name);
  let entries: Record<string, Uint8Array>;
  let mediaNames: string[];
  try {
    mediaNames = [];
    entries = unzipSync(bytes, {
      filter: (f) => {
        if (!isMedia(f.name)) return true;
        mediaNames.push(f.name);
        return false;
      },
    });
  } catch (e) {
    throw new ImportError("NOT_ZIP", `Corrupt zip container: ${e instanceof Error ? e.message : String(e)}`);
  }
  const texts = new Map<string, string>();
  const lazyExtract = (name: string): Uint8Array | undefined => {
    const extracted = unzipSync(bytes, { filter: (f) => f.name === name })[name];
    if (extracted) entries[name] = extracted; // cache for repeated references
    return extracted;
  };
  return {
    part: (name) => entries[name] ?? (isMedia(name) ? lazyExtract(name) : undefined),
    text(name) {
      const cached = texts.get(name);
      if (cached !== undefined) return cached;
      const raw = entries[name];
      if (!raw) return undefined;
      const decoded = strFromU8(raw);
      texts.set(name, decoded);
      return decoded;
    },
    names: () => [...Object.keys(entries), ...mediaNames.filter((n) => !(n in entries))],
  };
}
