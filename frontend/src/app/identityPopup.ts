// POC identity popup: when a backend is configured but the embedder didn't pass
// a `user`, ask for a First/Last name and derive a DETERMINISTIC id (same name →
// same id, so the backend recognizes a returning person). A real embedder skips
// this by passing `user` directly. Self-contained inline styles since it renders
// before the editor (and its stylesheet) mounts. POC: identity is NOT persisted —
// every page entry without a provided `user` prompts again.

import { deterministicUserId, type UserInfo } from "@cw/shared";

export function showIdentityPopup(): Promise<UserInfo> {
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

    const mkField = (label: string, value: string): HTMLInputElement => {
      const wrap = document.createElement("label");
      wrap.style.cssText = "display:flex;flex-direction:column;gap:3px;font-size:12px;color:#605e5c;margin-bottom:10px;";
      wrap.textContent = label;
      const input = document.createElement("input");
      input.type = "text";
      input.value = value;
      input.style.cssText =
        "height:30px;border:1px solid #c8c6c4;border-radius:5px;padding:0 9px;font:inherit;font-size:14px;color:#323130;";
      wrap.appendChild(input);
      box.appendChild(wrap);
      return input;
    };

    const first = mkField("First name", "");
    const last = mkField("Last name", "");

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
      const user: UserInfo = { id: deterministicUserId(`${f} ${l}`), firstName: f, lastName: l };
      back.remove();
      resolve(user);
    };
    go.addEventListener("click", submit);
    box.addEventListener("keydown", (e) => {
      if (e.key === "Enter") submit();
    });
  });
}
