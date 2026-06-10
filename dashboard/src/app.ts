// Admin dashboard shell: login gate + tabbed views (Documents, Users, Webhooks,
// Upload). Vanilla TS; each view fetches and renders into the main panel.

import { api, AuthError, BACKEND_URL, downloadUrl, editorUrl, getToken, login } from "./api";
import { clear, el, fmtDate, fmtName } from "./dom";
import type { ApiTokenCreated, DeliveryRecord, WebhookRecord } from "./types";

const root = document.getElementById("app")!;

function mount(): void {
  if (getToken()) renderShell();
  else renderLogin();
}

// --- login ------------------------------------------------------------------

function renderLogin(): void {
  const err = el("div", { class: "err" });
  const username = el("input", { placeholder: "Username", autofocus: true });
  const password = el("input", { type: "password", placeholder: "Password" });

  const submit = async (e: Event): Promise<void> => {
    e.preventDefault();
    err.textContent = "";
    try {
      await login(username.value.trim(), password.value);
      renderShell();
    } catch (ex) {
      err.textContent = ex instanceof Error ? ex.message : String(ex);
    }
  };

  const form = el(
    "form",
    { class: "card stack", onsubmit: submit },
    el("h1", {}, "canvas-word · admin"),
    el("label", { class: "muted" }, "Username"),
    username,
    el("label", { class: "muted" }, "Password"),
    password,
    err,
    el("button", { class: "primary", type: "submit" }, "Sign in"),
  );
  clear(root);
  root.append(el("div", { class: "login" }, form));
}

// --- shell ------------------------------------------------------------------

type Tab = { id: string; label: string; render: (main: HTMLElement) => void };

const TABS: Tab[] = [
  { id: "documents", label: "Documents", render: renderDocuments },
  { id: "users", label: "Users", render: renderUsers },
  { id: "webhooks", label: "Webhooks", render: renderWebhooks },
  { id: "tokens", label: "Integrations", render: renderTokens },
  { id: "upload", label: "Upload", render: renderUpload },
];

function renderShell(activeId = "documents"): void {
  const main = el("main");
  const tabButtons = TABS.map((t) =>
    el(
      "button",
      {
        class: `tab${t.id === activeId ? " active" : ""}`,
        onclick: () => renderShell(t.id),
      },
      t.label,
    ),
  );
  const header = el(
    "header",
    { class: "bar" },
    el("h1", {}, "canvas-word"),
    el("nav", { class: "tabs" }, ...tabButtons),
    el("div", { class: "spacer" }),
    el(
      "button",
      {
        onclick: () => {
          sessionStorage.removeItem("cw-admin-token");
          renderLogin();
        },
      },
      "Sign out",
    ),
  );
  clear(root);
  root.append(header, main);
  TABS.find((t) => t.id === activeId)!.render(main);
}

/** Run an async view body, bouncing to login on auth failure. */
function guard(main: HTMLElement, body: () => Promise<void>): void {
  main.append(el("p", { class: "muted" }, "Loading…"));
  body().catch((ex) => {
    if (ex instanceof AuthError) return renderLogin();
    clear(main);
    main.append(el("p", { class: "err" }, ex instanceof Error ? ex.message : String(ex)));
  });
}

// --- documents --------------------------------------------------------------

function renderDocuments(main: HTMLElement): void {
  guard(main, async () => {
    const docs = await api.listDocuments();
    clear(main);
    if (docs.length === 0) {
      main.append(el("p", { class: "muted" }, "No documents yet. Upload one from the Upload tab."));
      return;
    }
    const rows = docs.map((d) =>
      el(
        "tr",
        {},
        el("td", {}, el("div", {}, d.title || "Untitled"), el("code", { class: "mono muted" }, d.docId)),
        el("td", {}, fmtName(d.creatorFirstName, d.creatorLastName)),
        el("td", {}, fmtDate(d.createdAt)),
        el("td", {}, fmtDate(d.lastEditedAt)),
        el("td", {}, String(d.contributorCount)),
        el("td", {}, String(d.headVersion)),
        el(
          "td",
          { class: "actions" },
          el("a", { href: editorUrl(d.docId), target: "_blank", rel: "noreferrer" }, "Open"),
          el("a", { href: downloadUrl(d.docId, "docx") }, ".docx"),
          el("a", { href: downloadUrl(d.docId, "pdf") }, ".pdf"),
          el("a", { href: "#", onclick: (e: Event) => { e.preventDefault(); showActivity(d.docId, d.title); } }, "Activity"),
        ),
      ),
    );
    main.append(
      table(
        ["Document", "Created by", "Created", "Last edited", "Contributors", "Version", ""],
        rows,
      ),
    );
  });
}

function showActivity(docId: string, title: string | null): void {
  const main = document.querySelector("main")!;
  clear(main);
  main.append(
    el("div", { class: "row" },
      el("button", { onclick: () => renderShell("documents") }, "← Back"),
      el("h2", {}, title || "Untitled"),
    ),
  );
  guard(main, async () => {
    const activity = await api.activity(docId);
    // guard() left a "Loading…" note; rebuild cleanly below the header.
    const back = main.firstChild;
    clear(main);
    if (back) main.append(back);
    const rows = activity.entries.map((e) =>
      el(
        "tr",
        {},
        el("td", {}, String(e.seq)),
        el("td", {}, fmtName(e.firstName, e.lastName, e.userId ?? "—")),
        el("td", {}, e.origin),
        el("td", {}, String(e.opCount)),
        el("td", {}, fmtDate(e.ts)),
      ),
    );
    main.append(
      el("p", { class: "muted" }, `Created by ${fmtName(activity.creatorFirstName, activity.creatorLastName)} · ${fmtDate(activity.createdAt)}`),
      table(["Seq", "Author", "Origin", "Ops", "When"], rows),
    );
  });
}

// --- users ------------------------------------------------------------------

function renderUsers(main: HTMLElement): void {
  guard(main, async () => {
    const users = await api.listUsers();
    clear(main);
    if (users.length === 0) {
      main.append(el("p", { class: "muted" }, "No users have edited a document yet."));
      return;
    }
    const rows = users.map((u) =>
      el(
        "tr",
        {},
        el("td", {}, fmtName(u.firstName, u.lastName), el("div", {}, el("code", { class: "mono muted" }, u.id))),
        el("td", {}, String(u.docsCreated)),
        el("td", {}, String(u.changeCount)),
        el("td", {}, fmtDate(u.lastActive)),
      ),
    );
    main.append(table(["User", "Docs created", "Changes", "Last active"], rows));
  });
}

// --- webhooks ---------------------------------------------------------------

function renderWebhooks(main: HTMLElement): void {
  guard(main, async () => {
    const hooks = await api.listWebhooks();
    clear(main);
    main.append(webhookForm(() => renderShell("webhooks")));
    if (hooks.length === 0) {
      main.append(el("p", { class: "muted" }, "No webhooks registered."));
      return;
    }
    main.append(...hooks.map(webhookCard));
  });
}

function webhookForm(onCreated: () => void): HTMLElement {
  const url = el("input", { placeholder: "https://example.com/hooks/canvas-word", type: "url" });
  const err = el("div", { class: "err" });
  const submit = async (e: Event): Promise<void> => {
    e.preventDefault();
    err.textContent = "";
    if (!url.value.trim()) {
      err.textContent = "URL is required";
      return;
    }
    try {
      const hook = await api.createWebhook({ url: url.value.trim() });
      alert(`Webhook created. Signing secret (shown once):\n\n${hook.secret}`);
      onCreated();
    } catch (ex) {
      if (ex instanceof AuthError) return renderLogin();
      err.textContent = ex instanceof Error ? ex.message : String(ex);
    }
  };
  return el(
    "form",
    { class: "card stack", onsubmit: submit, style: "margin-bottom:16px" },
    el("h2", {}, "Add webhook"),
    el("p", { class: "muted" }, "Fired when the last editor of a document disconnects (document.session_ended). The signing secret is generated and shown once."),
    el("div", { class: "row" }, url, el("button", { class: "primary", type: "submit" }, "Create")),
    err,
  );
}

function webhookCard(h: WebhookRecord): HTMLElement {
  const body = el("div", { class: "card stack", style: "margin-bottom:12px" });
  const deliveriesBox = el("div");
  body.append(
    el(
      "div",
      { class: "row" },
      el("code", { class: "mono" }, h.url),
      el("span", { class: `pill ${h.active ? "ok" : "off"}` }, h.active ? "active" : "paused"),
      el("div", { class: "spacer" }),
      el(
        "button",
        {
          onclick: async () => {
            try {
              await api.setWebhookActive(h.id, !h.active);
              renderShell("webhooks");
            } catch (ex) {
              if (ex instanceof AuthError) renderLogin();
            }
          },
        },
        h.active ? "Pause" : "Resume",
      ),
      el(
        "button",
        {
          onclick: () => {
            clear(deliveriesBox);
            void loadDeliveries(h.id, deliveriesBox);
          },
        },
        "Deliveries",
      ),
      el(
        "button",
        {
          class: "danger",
          onclick: async () => {
            if (!confirm("Delete this webhook?")) return;
            try {
              await api.deleteWebhook(h.id);
              renderShell("webhooks");
            } catch (ex) {
              if (ex instanceof AuthError) renderLogin();
            }
          },
        },
        "Delete",
      ),
    ),
    el("div", { class: "muted" }, `Events: ${h.events.join(", ")} · created ${fmtDate(h.createdAt)}`),
    deliveriesBox,
  );
  return body;
}

async function loadDeliveries(webhookId: string, box: HTMLElement): Promise<void> {
  box.append(el("p", { class: "muted" }, "Loading deliveries…"));
  let deliveries: DeliveryRecord[];
  try {
    deliveries = await api.deliveries(webhookId);
  } catch (ex) {
    if (ex instanceof AuthError) return renderLogin();
    clear(box);
    box.append(el("p", { class: "err" }, ex instanceof Error ? ex.message : String(ex)));
    return;
  }
  clear(box);
  if (deliveries.length === 0) {
    box.append(el("p", { class: "muted" }, "No deliveries yet."));
    return;
  }
  const rows = deliveries.map((d) =>
    el(
      "tr",
      {},
      el("td", {}, el("span", { class: `pill ${d.status === "success" ? "ok" : "fail"}` }, d.status)),
      el("td", {}, d.responseCode != null ? String(d.responseCode) : "—"),
      el("td", {}, String(d.attempts)),
      el("td", {}, el("code", { class: "mono muted" }, d.docId ?? "—")),
      el("td", {}, fmtDate(d.createdAt)),
      el("td", { class: "muted" }, d.lastError ?? ""),
    ),
  );
  box.append(table(["Status", "HTTP", "Attempts", "Doc", "When", "Error"], rows));
}

// --- integrations (API tokens) ----------------------------------------------

function renderTokens(main: HTMLElement): void {
  clear(main);
  const reveal = el("div"); // one-time token box appears here after create
  const listBox = el("div");
  main.append(tokenForm(reveal, listBox), reveal, listBox);
  void refreshTokenList(listBox);
}

function tokenForm(reveal: HTMLElement, listBox: HTMLElement): HTMLElement {
  const name = el("input", { placeholder: "e.g. partner-crm" });
  const err = el("div", { class: "err" });
  const submit = async (e: Event): Promise<void> => {
    e.preventDefault();
    err.textContent = "";
    try {
      const created = await api.createToken(name.value.trim() || "integration");
      name.value = "";
      clear(reveal);
      reveal.append(tokenRevealBox(created));
      await refreshTokenList(listBox);
    } catch (ex) {
      if (ex instanceof AuthError) return renderLogin();
      err.textContent = ex instanceof Error ? ex.message : String(ex);
    }
  };
  return el(
    "form",
    { class: "card stack", onsubmit: submit, style: "margin-bottom:16px" },
    el("h2", {}, "Create integration token"),
    el("p", { class: "muted" }, "A long-lived API key a 3rd-party system uses to POST /upload and get back a docId. Shown once; only its hash is stored."),
    el("div", { class: "row" }, name, el("button", { class: "primary", type: "submit" }, "Create token")),
    err,
  );
}

function tokenRevealBox(created: ApiTokenCreated): HTMLElement {
  const usage = `curl -X POST ${BACKEND_URL}/upload \\
  -H "X-API-Key: ${created.token}" \\
  -H "X-Filename: report.docx" \\
  --data-binary @report.docx
# -> { "docId": "...", "version": 0, "warnings": [] }`;
  return el(
    "div",
    { class: "card stack", style: "border-color:var(--accent);margin-bottom:16px" },
    el("strong", {}, "Copy this token now — it won't be shown again."),
    el("code", { class: "mono" }, created.token),
    el("button", { onclick: () => void navigator.clipboard?.writeText(created.token) }, "Copy"),
    el(
      "details",
      {},
      el("summary", { class: "muted" }, "How a 3rd-party system uses it"),
      el("pre", { class: "mono muted", style: "white-space:pre-wrap" }, usage),
    ),
  );
}

async function refreshTokenList(listBox: HTMLElement): Promise<void> {
  clear(listBox);
  let tokens;
  try {
    tokens = await api.listTokens();
  } catch (ex) {
    if (ex instanceof AuthError) return renderLogin();
    listBox.append(el("p", { class: "err" }, ex instanceof Error ? ex.message : String(ex)));
    return;
  }
  if (tokens.length === 0) {
    listBox.append(el("p", { class: "muted" }, "No integration tokens yet."));
    return;
  }
  const rows = tokens.map((t) =>
    el(
      "tr",
      {},
      el("td", {}, el("div", {}, t.name), el("code", { class: "mono muted" }, `${t.prefix}…`)),
      el("td", {}, el("span", { class: `pill ${t.active ? "ok" : "off"}` }, t.active ? "active" : "revoked")),
      el("td", {}, fmtDate(t.createdAt)),
      el("td", {}, fmtDate(t.lastUsedAt)),
      el(
        "td",
        { class: "actions" },
        el(
          "button",
          {
            class: "danger",
            onclick: async () => {
              if (!confirm(`Revoke "${t.name}"? Integrations using it will stop working.`)) return;
              try {
                await api.revokeToken(t.id);
                await refreshTokenList(listBox);
              } catch (ex) {
                if (ex instanceof AuthError) renderLogin();
              }
            },
          },
          "Revoke",
        ),
      ),
    ),
  );
  listBox.append(table(["Token", "Status", "Created", "Last used", ""], rows));
}

// --- upload -----------------------------------------------------------------

function renderUpload(main: HTMLElement): void {
  clear(main);
  const file = el("input", { type: "file", accept: ".docx" });
  const out = el("div", { class: "stack" });
  const err = el("div", { class: "err" });

  const submit = async (e: Event): Promise<void> => {
    e.preventDefault();
    err.textContent = "";
    clear(out);
    const f = file.files?.[0];
    if (!f) {
      err.textContent = "Choose a .docx file first";
      return;
    }
    out.append(el("p", { class: "muted" }, "Uploading & parsing…"));
    try {
      const result = await api.upload(f);
      clear(out);
      out.append(
        el("div", { class: "card stack" },
          el("div", {}, "Uploaded. Document id:"),
          el("code", { class: "mono" }, result.docId),
          el("div", { class: "actions" },
            el("a", { class: "primary", href: editorUrl(result.docId), target: "_blank", rel: "noreferrer", style: "padding:6px 12px;border-radius:8px;color:#fff;background:var(--accent)" }, "Open in editor"),
            el("a", { href: downloadUrl(result.docId, "docx") }, "Download .docx"),
            el("a", { href: downloadUrl(result.docId, "pdf") }, "Download .pdf"),
          ),
          result.warnings.length > 0
            ? el("details", {}, el("summary", { class: "muted" }, `${result.warnings.length} import warning(s)`),
                el("ul", {}, ...result.warnings.map((w) => el("li", { class: "muted" }, w.message))))
            : el("div", { class: "muted" }, "No import warnings."),
        ),
      );
    } catch (ex) {
      if (ex instanceof AuthError) return renderLogin();
      clear(out);
      err.textContent = ex instanceof Error ? ex.message : String(ex);
    }
  };

  main.append(
    el("form", { class: "card stack", onsubmit: submit },
      el("h2", {}, "Upload a document"),
      el("p", { class: "muted" }, "Upload a .docx — the backend parses it and returns a docId you can open for review."),
      el("div", { class: "row" }, file, el("button", { class: "primary", type: "submit" }, "Upload")),
      err,
    ),
    out,
  );
}

// --- shared -----------------------------------------------------------------

function table(headers: string[], rows: HTMLElement[]): HTMLElement {
  return el(
    "table",
    {},
    el("thead", {}, el("tr", {}, ...headers.map((h) => el("th", {}, h)))),
    el("tbody", {}, ...rows),
  );
}

mount();
