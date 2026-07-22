// A tiny most-recently-used string list persisted in localStorage — shared by the
// shape gallery's "recently used" row (issue #244 A3) and the colour popover's
// "recent colours" row (issue #244 D1). Mirrors the symbolPicker.ts recent pattern:
// newest-first, de-duplicated, capped, and resilient to a disabled/absent
// localStorage (private mode, SSR) so callers never have to guard it.

/** Read the recent list for `key` (newest first), clamped to `max`. Returns [] on
 *  any storage error or malformed value. */
export function readRecent(key: string, max: number): string[] {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return [];
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr.filter((v): v is string => typeof v === "string").slice(0, max) : [];
  } catch {
    return [];
  }
}

/** Push `value` to the front of the recent list for `key` (de-duplicating and
 *  clamping to `max`), and return the new list. Swallows storage errors. */
export function pushRecent(key: string, value: string, max: number): string[] {
  const next = [value, ...readRecent(key, max).filter((v) => v !== value)].slice(0, max);
  try {
    localStorage.setItem(key, JSON.stringify(next));
  } catch {
    /* storage unavailable — recents are a nicety, not load-bearing */
  }
  return next;
}
