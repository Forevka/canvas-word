// A self-contained identity prompt for the demo. The package deliberately does
// NOT ship one: a real embedder already knows its user and passes `user` to the
// WordCanvas constructor. This mirrors that contract — it just collects a name
// and hands a `UserInfo` back, the same shape the constructor expects.

import type { UserInfo } from "@forevka/wordcanvas";

// Stable name → id hash (FNV-1a), so the same person gets the same id across
// reloads and the backend recognizes a returning collaborator.
function deterministicUserId(name: string): string {
  const s = name.trim().toLowerCase().replace(/\s+/g, " ");
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return "u" + (h >>> 0).toString(16).padStart(8, "0");
}

export function promptIdentity(): Promise<UserInfo> {
  return new Promise((resolve) => {
    const back = document.createElement("div");
    back.style.cssText =
      "position:fixed;inset:0;z-index:9999;background:rgba(0,0,0,0.35);display:flex;align-items:center;" +
      "justify-content:center;font-family:'Segoe UI',Roboto,Arial,sans-serif;";

    const box = document.createElement("div");
    box.style.cssText =
      "background:#fff;border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,0.3);padding:22px 24px;min-width:320px;";
    box.innerHTML =
      `<div style="font-size:16px;font-weight:600;color:#323130;margin-bottom:4px;">Who are you?</div>` +
      `<div style="font-size:12px;color:#605e5c;margin-bottom:14px;">Your name labels your caret and edits for collaborators.</div>`;

    const mkField = (label: string): HTMLInputElement => {
      const wrap = document.createElement("label");
      wrap.style.cssText = "display:flex;flex-direction:column;gap:3px;font-size:12px;color:#605e5c;margin-bottom:10px;";
      wrap.textContent = label;
      const input = document.createElement("input");
      input.type = "text";
      input.style.cssText =
        "height:30px;border:1px solid #c8c6c4;border-radius:5px;padding:0 9px;font:inherit;font-size:14px;color:#323130;";
      wrap.appendChild(input);
      box.appendChild(wrap);
      return input;
    };

    const first = mkField("First name");
    const last = mkField("Last name");

    const row = document.createElement("div");
    row.style.cssText = "display:flex;justify-content:flex-end;margin-top:6px;";
    const go = document.createElement("button");
    go.textContent = "Continue";
    go.style.cssText =
      "height:32px;padding:0 16px;border:none;border-radius:5px;background:#2b579a;color:#fff;font:inherit;" +
      "font-size:14px;font-weight:600;cursor:pointer;";
    row.appendChild(go);
    box.appendChild(row);
    back.appendChild(box);
    document.body.appendChild(back);
    setTimeout(() => first.focus(), 0);

    const submit = (): void => {
      const f = first.value.trim() || "Anonymous";
      const l = last.value.trim();
      back.remove();
      resolve({ id: deterministicUserId(`${f} ${l}`), firstName: f, lastName: l });
    };
    go.addEventListener("click", submit);
    box.addEventListener("keydown", (e) => {
      if (e.key === "Enter") submit();
    });
  });
}
