// Admin API client. The bearer token lives in sessionStorage; a 401 anywhere
// clears it and surfaces as AuthError so the shell can bounce back to login.

import type {
  ApiTokenCreated,
  ApiTokenRecord,
  DeliveryRecord,
  DocumentActivity,
  DocumentSummary,
  UploadResult,
  UserSummary,
  WebhookRecord,
} from "./types";

// Default to the page's own origin: in production the dashboard is served at
// <origin>/dashboard/ and the backend API (/admin/*, /upload, /docs/*) is
// reverse-proxied on the same origin. Local dev overrides both via
// .env.development (the editor + backend run on different ports there).
export const BACKEND_URL = import.meta.env.VITE_BACKEND_URL ?? location.origin;
export const EDITOR_URL = import.meta.env.VITE_EDITOR_URL ?? location.origin;

const TOKEN_KEY = "cw-admin-token";
export const getToken = (): string | null => sessionStorage.getItem(TOKEN_KEY);
const setToken = (t: string): void => sessionStorage.setItem(TOKEN_KEY, t);
export const clearToken = (): void => sessionStorage.removeItem(TOKEN_KEY);

export class AuthError extends Error {}

async function req(path: string, init: RequestInit = {}): Promise<Response> {
  const headers = new Headers(init.headers);
  const token = getToken();
  if (token) headers.set("authorization", `Bearer ${token}`);
  const res = await fetch(`${BACKEND_URL}${path}`, { ...init, headers });
  if (res.status === 401) {
    clearToken();
    throw new AuthError("session expired");
  }
  if (!res.ok) {
    const body = (await res.json().catch(() => ({}))) as { error?: string };
    throw new Error(body.error ?? `HTTP ${res.status}`);
  }
  return res;
}

const json = <T>(r: Response): Promise<T> => r.json() as Promise<T>;
const JSON_HEADERS = { "content-type": "application/json" };

export async function login(username: string, password: string): Promise<void> {
  const res = await fetch(`${BACKEND_URL}/admin/login`, {
    method: "POST",
    headers: JSON_HEADERS,
    body: JSON.stringify({ username, password }),
  });
  if (!res.ok) throw new Error("Invalid username or password");
  const { token } = (await json<{ token: string }>(res));
  setToken(token);
}

export const api = {
  listDocuments: () =>
    req("/admin/docs").then(json<{ documents: DocumentSummary[] }>).then((d) => d.documents),
  listUsers: () => req("/admin/users").then(json<{ users: UserSummary[] }>).then((d) => d.users),
  activity: (docId: string) => req(`/admin/docs/${docId}/activity`).then(json<DocumentActivity>),
  listWebhooks: () =>
    req("/admin/webhooks").then(json<{ webhooks: WebhookRecord[] }>).then((d) => d.webhooks),
  createWebhook: (body: { url: string; secret?: string; events?: string[] }) =>
    req("/admin/webhooks", { method: "POST", headers: JSON_HEADERS, body: JSON.stringify(body) }).then(
      json<WebhookRecord>,
    ),
  deleteWebhook: (id: string) => req(`/admin/webhooks/${id}`, { method: "DELETE" }).then(() => undefined),
  setWebhookActive: (id: string, active: boolean) =>
    req(`/admin/webhooks/${id}`, { method: "PATCH", headers: JSON_HEADERS, body: JSON.stringify({ active }) }).then(
      () => undefined,
    ),
  deliveries: (id: string) =>
    req(`/admin/webhooks/${id}/deliveries`).then(json<{ deliveries: DeliveryRecord[] }>).then((d) => d.deliveries),
  listTokens: () =>
    req("/admin/tokens").then(json<{ tokens: ApiTokenRecord[] }>).then((d) => d.tokens),
  createToken: (name: string) =>
    req("/admin/tokens", { method: "POST", headers: JSON_HEADERS, body: JSON.stringify({ name }) }).then(
      json<ApiTokenCreated>,
    ),
  revokeToken: (id: string) => req(`/admin/tokens/${id}`, { method: "DELETE" }).then(() => undefined),
  upload: (file: File, user?: { id: string; firstName: string; lastName: string }) => {
    const headers = new Headers();
    headers.set("x-filename", file.name);
    if (user) headers.set("x-user", JSON.stringify(user));
    return req("/upload", { method: "POST", headers, body: file }).then(json<UploadResult>);
  },
};

export const downloadUrl = (docId: string, fmt: "docx" | "pdf"): string =>
  `${BACKEND_URL}/docs/${docId}/export.${fmt}`;

/** Open the editor against this backend on a specific document. */
export const editorUrl = (docId: string): string =>
  `${EDITOR_URL}/?backend=${encodeURIComponent(BACKEND_URL)}&doc=${encodeURIComponent(docId)}`;
