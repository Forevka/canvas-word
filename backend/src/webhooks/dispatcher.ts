// Webhook dispatch. Fired when the last editor of a document disconnects (see
// the Broadcaster's room-empty debounce). Notifies every active, subscribed
// endpoint that the document was edited and can be downloaded, with an
// HMAC-signed body (X-CW-Signature) and bounded retries. Every attempt cluster
// is logged so the dashboard can show delivery history.

import { createHmac } from "node:crypto";
import { applyReviewOp, emptyReview, WEBHOOK_EVENT, type Comment, type ReviewOp } from "@cw/shared";
import type { ChangeStore, DocumentActivity } from "../store/ChangeStore";

export const SESSION_ENDED_EVENT = WEBHOOK_EVENT.SessionEnded;
/** Fired when a comment (or reply) @-mentions one or more users, so a third-party
 *  system can notify them. Embedders subscribe a webhook with this event. */
export const MENTION_EVENT = WEBHOOK_EVENT.CommentMention;
const MAX_ATTEMPTS = Number(process.env.WEBHOOK_MAX_ATTEMPTS ?? 3);
// Retry backoff: base doubles each attempt, capped. Env-overridable like the others.
const WEBHOOK_BACKOFF_BASE_MS = Number(process.env.WEBHOOK_BACKOFF_BASE_MS ?? 1000);
const WEBHOOK_BACKOFF_MAX_MS = Number(process.env.WEBHOOK_BACKOFF_MAX_MS ?? 8000);

export interface EditorRef {
  id?: string;
  firstName?: string;
  lastName?: string;
}

export interface SessionEndContext {
  /** The participant whose disconnect emptied the room, if known. */
  lastEditor?: EditorRef | undefined;
}

const delay = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));
const backoffMs = (attempt: number): number =>
  Math.min(WEBHOOK_BACKOFF_BASE_MS * 2 ** (attempt - 1), WEBHOOK_BACKOFF_MAX_MS);

export class WebhookDispatcher {
  constructor(
    private readonly store: ChangeStore,
    private readonly baseUrl: string = process.env.PUBLIC_BASE_URL ??
      `http://localhost:${process.env.BACKEND_PORT ?? 8787}`,
  ) {}

  /** Build + send the session-ended notification. Safe to call fire-and-forget;
   *  it swallows its own errors (delivery failures are logged to the store). */
  async onSessionEnded(docId: string, ctx: SessionEndContext = {}): Promise<void> {
    try {
      const activity = await this.store.getActivity(docId);
      // Nothing to notify about if the doc is missing or saw no edits this life.
      if (!activity || activity.entries.length === 0) return;

      const subscribed = (await this.store.listWebhooks()).filter(
        (w) => w.active && w.events.includes(SESSION_ENDED_EVENT),
      );
      if (subscribed.length === 0) return;

      const [title, head, reviewOps] = await Promise.all([
        this.store.getDocumentTitle(docId),
        this.store.getHead(docId),
        this.store.getReviewOps(docId),
      ]);
      const review = reviewOps.reduce((layer, env) => applyReviewOp(layer, env.op).layer, emptyReview(docId));
      const editors = dedupeEditors(activity);
      const payload = {
        event: SESSION_ENDED_EVENT,
        docId,
        title,
        version: head ?? 0,
        editors,
        lastEditor: ctx.lastEditor ?? editors[0] ?? null,
        // Track-changes + comments state at session end (overlay, not the doc).
        review: {
          suggestionCount: review.suggestions.length,
          openThreadCount: review.threads.filter((t) => t.status === "open").length,
          suggestions: review.suggestions,
          threads: review.threads,
        },
        downloadUrls: {
          docx: `${this.baseUrl}/docs/${docId}/export.docx`,
          pdf: `${this.baseUrl}/docs/${docId}/export.pdf`,
        },
        ts: Date.now(),
      };
      const body = JSON.stringify(payload);
      await Promise.all(subscribed.map((w) => this.deliver(w.id, w.url, w.secret, docId, body)));
    } catch {
      // Never let webhook dispatch take down the connection handler.
    }
  }

  /** Notify subscribers when a comment/reply @-mentions users, so a third-party
   *  system can ping them. Fire-and-forget; swallows its own errors. */
  async onReviewOp(docId: string, op: ReviewOp): Promise<void> {
    try {
      const m = mentionFromOp(op);
      if (!m || !m.comment.mentions || m.comment.mentions.length === 0) return;
      const subscribed = (await this.store.listWebhooks()).filter(
        (w) => w.active && w.events.includes(MENTION_EVENT),
      );
      if (subscribed.length === 0) return;
      const c = m.comment;
      const payload = {
        event: MENTION_EVENT,
        docId,
        threadId: m.threadId,
        commentId: c.id,
        author: c.author,
        body: c.body.map((r) => r.text).join(""),
        mentions: c.mentions,
        ts: Date.now(),
      };
      const body = JSON.stringify(payload);
      await Promise.all(subscribed.map((w) => this.deliver(w.id, w.url, w.secret, docId, body, MENTION_EVENT)));
    } catch {
      // Never let webhook dispatch take down the connection handler.
    }
  }

  private async deliver(
    webhookId: string,
    url: string,
    secret: string,
    docId: string,
    body: string,
    event: string = SESSION_ENDED_EVENT,
  ): Promise<void> {
    const sig = "sha256=" + createHmac("sha256", secret).update(body).digest("hex");
    let attempt = 0;
    let responseCode: number | null = null;
    let lastError: string | null = null;
    while (attempt < MAX_ATTEMPTS) {
      attempt++;
      try {
        const res = await fetch(url, {
          method: "POST",
          headers: {
            "content-type": "application/json",
            "x-cw-signature": sig,
            "x-cw-event": event,
          },
          body,
        });
        responseCode = res.status;
        if (res.ok) {
          await this.store.recordDelivery({
            webhookId,
            docId,
            event,
            status: "success",
            responseCode,
            attempts: attempt,
            lastError: null,
          });
          return;
        }
        lastError = `HTTP ${res.status}`;
      } catch (e) {
        lastError = e instanceof Error ? e.message : String(e);
      }
      if (attempt < MAX_ATTEMPTS) await delay(backoffMs(attempt));
    }
    await this.store.recordDelivery({
      webhookId,
      docId,
      event,
      status: "failed",
      responseCode,
      attempts: attempt,
      lastError,
    });
  }
}

/** The comment carrying mentions from a review op (a new thread's root comment or
 *  an added reply), or null for ops without a comment. */
function mentionFromOp(op: ReviewOp): { threadId: string; comment: Comment } | null {
  if (op.type === "addThread") {
    const root = op.t.comments[0];
    return root ? { threadId: op.t.id, comment: root } : null;
  }
  if (op.type === "addComment") return { threadId: op.threadId, comment: op.c };
  return null;
}

/** Distinct authors (by user id) in most-recent-first order, with display names. */
function dedupeEditors(activity: DocumentActivity): EditorRef[] {
  const seen = new Map<string, EditorRef>();
  for (const e of activity.entries) {
    if (e.userId && !seen.has(e.userId)) {
      seen.set(e.userId, { id: e.userId, firstName: e.firstName, lastName: e.lastName });
    }
  }
  return [...seen.values()];
}
