// Deduplicating warning collector — mirrors the importer's WarningSink. One entry
// per code; repeats bump `count` so the UI can show "image-format-unsupported ×4".

import type { ExportWarning } from "./types";

export class WarningSink {
  private readonly map = new Map<string, ExportWarning>();

  warn(code: string, detail?: string): void {
    const existing = this.map.get(code);
    if (existing) {
      existing.count++;
      return;
    }
    this.map.set(code, detail === undefined ? { code, count: 1 } : { code, detail, count: 1 });
  }

  list(): ExportWarning[] {
    return [...this.map.values()];
  }
}
