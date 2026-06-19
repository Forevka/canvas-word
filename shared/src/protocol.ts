// Collaboration wire protocol constants, shared by the editor's SyncClient and
// the server's Broadcaster + /ws route. Both ends compare a raw `msg.type`
// string, so centralizing the discriminators here is what stops the two from
// drifting. Payload shapes live with each end (see SyncClient / Broadcaster).

/** WebSocket message `type` discriminators. Direction noted per entry. */
export const WS_MSG = {
  /** client → server: identify (siteId + user), request roster. */
  Hello: "hello",
  /** client → server: submit a change for the canonical order. */
  Submit: "submit",
  /** both ways: ephemeral caret/selection presence. */
  Presence: "presence",
  /** both ways: a review op (suggestion/comment). */
  Review: "review",
  /** server → client: full roster snapshot on join. */
  Roster: "roster",
  /** server → client: an accepted change (also the sender's ack). */
  Change: "change",
  /** server → client: a peer left the room. */
  Leave: "leave",
} as const;

export type WsMsgType = (typeof WS_MSG)[keyof typeof WS_MSG];

/** Webhook event names dispatched to third-party integrations. */
export const WEBHOOK_EVENT = {
  /** The last collaborator left a document (debounced). */
  SessionEnded: "document.session_ended",
  /** A review op @-mentioned someone. */
  CommentMention: "comment.mention",
} as const;

export type WebhookEvent = (typeof WEBHOOK_EVENT)[keyof typeof WEBHOOK_EVENT];
