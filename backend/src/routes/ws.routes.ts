// WebSocket /ws?doc=<id> — live change broadcast + presence. The socket-level
// fan-out lives in the Broadcaster; this only decodes inbound messages and
// dispatches them to the store / broadcaster / webhook dispatcher.
import type { FastifyInstance, FastifyRequest } from "fastify";
import type { WebSocket } from "ws";
import type { Change, ReviewOpEnvelope, UserInfo } from "@cw/shared";
import type { AppContext } from "../context";

export function registerWsRoutes(app: FastifyInstance, { store, bcast, dispatcher }: AppContext): void {
  app.get("/ws", { websocket: true, schema: { hide: true } }, (socket: WebSocket, req: FastifyRequest) => {
    const docId = (req.query as { doc?: string }).doc;
    if (!docId) return socket.close(1008, "missing ?doc");
    bcast.join(docId, socket);
    socket.on("message", (data) => {
      void (async () => {
        let msg: { type?: string; change?: Change; review?: ReviewOpEnvelope; siteId?: string; user?: UserInfo; selection?: unknown };
        try {
          msg = JSON.parse(String(data));
        } catch {
          return;
        }
        if (msg.type === "submit" && msg.change) {
          // Append (with OT rebase) then broadcast to everyone in the room,
          // including the sender — whose client treats the echo as its ack.
          const accepted = await store.appendChange(docId, msg.change);
          bcast.publish(docId, accepted);
        } else if (msg.type === "review" && msg.review) {
          // Persist the review op (append log) and relay to the rest of the room.
          await store.appendReviewOp(docId, msg.review);
          bcast.publishReview(docId, msg.review, socket);
          // Notify third-party systems about any @-mentions in the op.
          void dispatcher.onReviewOp(docId, msg.review.op);
        } else if (msg.type === "hello" && msg.siteId) {
          if (msg.user) await store.upsertUser(msg.user);
          bcast.hello(socket, msg.siteId, msg.user);
        } else if (msg.type === "presence") {
          bcast.presence(socket, msg.selection ?? null);
        }
      })();
    });
  });
}
