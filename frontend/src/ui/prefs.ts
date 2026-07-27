// Tiny scalar application-preference store persisted in localStorage — the same
// convention as recentList.ts / the inspector-collapse overrides, but for a single
// enum value rather than an MRU list. Used by the Settings surface (theme, chrome
// preset). Resilient to a disabled/absent localStorage (private mode, SSR) so
// callers never have to guard it; an unknown/absent value returns null.

/** Read the preference at `key`, but only if it is one of `allowed`. Returns null
 *  on any storage error, an absent value, or a value outside `allowed`. */
export function readPref<T extends string>(key: string, allowed: readonly T[]): T | null {
  try {
    const raw = localStorage.getItem(key);
    return raw && (allowed as readonly string[]).includes(raw) ? (raw as T) : null;
  } catch {
    return null;
  }
}

/** Persist `value` at `key`. Swallows storage errors (prefs are a nicety). */
export function writePref(key: string, value: string): void {
  try {
    localStorage.setItem(key, value);
  } catch {
    /* storage unavailable — the preference just won't persist across reloads */
  }
}
