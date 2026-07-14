// A small, self-dismissing toast for non-blocking notices — e.g. "some linked
// images couldn't be embedded in the export". Deliberately self-contained
// (inline styles, no dependency on ui/styles.ts) so it renders identically no
// matter how the editor is embedded, and stacks bottom-right without stealing focus.

export type NoticeKind = "warning" | "info";

const COLORS: Record<NoticeKind, { bar: string; icon: string }> = {
  warning: { bar: "#d9822b", icon: "⚠" },
  info: { bar: "#2b6cb0", icon: "ℹ" },
};

let host: HTMLDivElement | undefined;

function ensureHost(): HTMLDivElement {
  if (host && host.isConnected) return host;
  host = document.createElement("div");
  Object.assign(host.style, {
    position: "fixed",
    right: "16px",
    bottom: "16px",
    zIndex: "2147483000", // above the busy overlay and app chrome
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    maxWidth: "min(420px, calc(100vw - 32px))",
    pointerEvents: "none",
  } satisfies Partial<CSSStyleDeclaration>);
  document.body.appendChild(host);
  return host;
}

/** Show a toast that auto-dismisses after `timeoutMs` (0 = sticky until clicked).
 *  Returns a dismiss function. */
export function showNotice(message: string, kind: NoticeKind = "warning", timeoutMs = 8000): () => void {
  const c = COLORS[kind];
  const card = document.createElement("div");
  Object.assign(card.style, {
    pointerEvents: "auto",
    display: "flex",
    alignItems: "flex-start",
    gap: "10px",
    padding: "12px 14px",
    borderRadius: "8px",
    borderLeft: `4px solid ${c.bar}`,
    background: "#1f2430",
    color: "#e7ebf0",
    font: "13px/1.45 -apple-system, Segoe UI, Roboto, sans-serif",
    boxShadow: "0 6px 24px rgba(0,0,0,0.28)",
    cursor: "pointer",
  } satisfies Partial<CSSStyleDeclaration>);
  card.setAttribute("role", "status");
  card.title = "Dismiss";

  const icon = document.createElement("span");
  icon.textContent = c.icon;
  icon.style.color = c.bar;
  icon.style.flex = "0 0 auto";
  icon.style.fontSize = "15px";

  const text = document.createElement("div");
  text.textContent = message;
  text.style.flex = "1 1 auto";

  card.append(icon, text);
  ensureHost().appendChild(card);

  let removed = false;
  const dismiss = (): void => {
    if (removed) return;
    removed = true;
    card.remove();
    if (host && host.childElementCount === 0) {
      host.remove();
      host = undefined;
    }
  };
  card.addEventListener("click", dismiss);
  if (timeoutMs > 0) setTimeout(dismiss, timeoutMs);
  return dismiss;
}
