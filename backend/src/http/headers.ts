// Small header / value helpers shared across HTTP routes.
import type { UserInfo } from "@cw/shared";

/** First value of a possibly-multi-valued header. */
export function headerStr(h: string | string[] | undefined): string | undefined {
  return Array.isArray(h) ? h[0] : h;
}

/** Parse the X-User header into a UserInfo (ignores malformed input). */
export function safeParseUser(raw: string): UserInfo | undefined {
  try {
    const u = JSON.parse(raw) as Partial<UserInfo>;
    if (u && typeof u.id === "string") {
      return { id: u.id, firstName: u.firstName ?? "", lastName: u.lastName ?? "" };
    }
  } catch {
    // ignore malformed header
  }
  return undefined;
}

/** Make a title safe for a Content-Disposition filename (no quotes/paths/control). */
export function sanitizeFilename(name: string): string {
  const cleaned = name.replace(/[-"\\/:*?<>|]+/g, "_").trim();
  return cleaned.length > 0 ? cleaned.slice(0, 120) : "document";
}
