// @-mention autocomplete for comment composers. Attaches to a <textarea>: typing
// "@" + a query pops a dropdown of the editor's knownUsers; picking one inserts
// "@Display Name " and records the user. Reusable by both the floating comment
// composer (index.ts) and the inline reply box (the Review panel). The body text
// carries the "@Name" spans; getMentions() returns the structured UserInfo list
// (filtered to mentions still present in the text) for Comment.mentions.

import { colorForId, userDisplayName, type UserInfo } from "@cw/shared";

export interface MentionAutocomplete {
  /** Users currently @-mentioned in the textarea (resolved identities). */
  getMentions(): UserInfo[];
  destroy(): void;
}

const initials = (u: UserInfo): string => ((u.firstName[0] ?? "?") + (u.lastName[0] ?? "")).toUpperCase();

export function attachMentionAutocomplete(ta: HTMLTextAreaElement, getUsers: () => UserInfo[]): MentionAutocomplete {
  const chosen = new Map<string, UserInfo>(); // everyone ever picked (id → user)
  let menu: HTMLDivElement | null = null;
  let items: UserInfo[] = [];
  let active = 0;
  let tokenStart = -1; // index of the '@' that opened the current token

  const close = (): void => {
    menu?.remove();
    menu = null;
    items = [];
    tokenStart = -1;
  };

  /** The "@query" token immediately before the caret (or null). '@' must start
   *  the value or follow whitespace; the query runs to the caret with no space. */
  const tokenAtCaret = (): { start: number; query: string } | null => {
    const v = ta.value;
    const caret = ta.selectionStart;
    let i = caret - 1;
    while (i >= 0 && !/\s/.test(v[i]!) && v[i] !== "@") i--;
    if (i < 0 || v[i] !== "@") return null;
    if (i > 0 && !/\s/.test(v[i - 1]!)) return null;
    return { start: i, query: v.slice(i + 1, caret) };
  };

  const render = (): void => {
    if (!menu) {
      menu = document.createElement("div");
      menu.className = "cw-mention-menu";
      menu.addEventListener("mousedown", (e) => e.preventDefault()); // keep textarea focus
      document.body.appendChild(menu);
    }
    menu.textContent = "";
    items.forEach((u, idx) => {
      const row = document.createElement("div");
      row.className = "cw-mention-item" + (idx === active ? " active" : "");
      const av = document.createElement("span");
      av.className = "cw-mention-av";
      av.style.background = colorForId(u.id);
      av.textContent = initials(u);
      const name = document.createElement("span");
      name.textContent = userDisplayName(u);
      row.append(av, name);
      row.addEventListener("click", () => pick(u));
      menu!.appendChild(row);
    });
    const r = ta.getBoundingClientRect();
    menu.style.left = `${r.left}px`;
    menu.style.top = `${r.bottom + 4}px`;
    menu.style.width = `${Math.max(200, r.width)}px`;
  };

  const update = (): void => {
    const tok = tokenAtCaret();
    if (!tok) return close();
    const q = tok.query.toLowerCase();
    items = getUsers()
      .filter((u) => userDisplayName(u).toLowerCase().includes(q))
      .slice(0, 6);
    if (items.length === 0) return close();
    tokenStart = tok.start;
    active = 0;
    render();
  };

  const pick = (u: UserInfo): void => {
    if (tokenStart < 0) return;
    const v = ta.value;
    const caret = ta.selectionStart;
    const before = v.slice(0, tokenStart);
    const insert = `@${userDisplayName(u)} `;
    ta.value = before + insert + v.slice(caret);
    const pos = (before + insert).length;
    ta.setSelectionRange(pos, pos);
    chosen.set(u.id, u);
    close();
    ta.focus();
    ta.dispatchEvent(new Event("input")); // re-enable the composer's submit button
  };

  const onInput = (): void => update();
  // Capture phase so an open menu pre-empts the composer's own Enter handler.
  const onKeydown = (e: KeyboardEvent): void => {
    if (!menu || items.length === 0) return;
    if (e.key === "ArrowDown") {
      e.preventDefault();
      active = (active + 1) % items.length;
      render();
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      active = (active - 1 + items.length) % items.length;
      render();
    } else if (e.key === "Enter" || e.key === "Tab") {
      e.preventDefault();
      e.stopPropagation();
      pick(items[active]!);
    } else if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      close();
    }
  };
  const onBlur = (): void => {
    setTimeout(close, 120); // allow a click on a menu row to land before closing
  };

  ta.addEventListener("input", onInput);
  ta.addEventListener("keydown", onKeydown, true);
  ta.addEventListener("blur", onBlur);

  return {
    getMentions: () => [...chosen.values()].filter((u) => ta.value.includes(`@${userDisplayName(u)}`)),
    destroy: () => {
      close();
      ta.removeEventListener("input", onInput);
      ta.removeEventListener("keydown", onKeydown, true);
      ta.removeEventListener("blur", onBlur);
    },
  };
}
