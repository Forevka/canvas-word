-- User identity + attribution. The embedder owns auth; we just persist whatever
-- { id, firstName, lastName } it passes so we can attribute documents and changes
-- (who created, who edited, when). All idempotent so Flyway can repair + re-run.

-- Caller-supplied users, upserted whenever a client identifies (WS hello) or
-- publishes a document. id is the embedder's user id (opaque to us).
CREATE TABLE IF NOT EXISTS users (
    id          text PRIMARY KEY,
    first_name  text NOT NULL DEFAULT '',
    last_name   text NOT NULL DEFAULT '',
    updated_at  bigint NOT NULL DEFAULT 0
);

-- Who created the document (the publisher's user id).
ALTER TABLE documents ADD COLUMN IF NOT EXISTS created_by text;

-- Who authored each change.
ALTER TABLE changes ADD COLUMN IF NOT EXISTS user_id text;

CREATE INDEX IF NOT EXISTS idx_changes_user ON changes (user_id);
