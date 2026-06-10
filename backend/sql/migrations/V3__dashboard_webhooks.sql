-- Dashboard + session-end webhooks. Idempotent (IF NOT EXISTS) so Flyway can
-- repair + re-run safely, matching V1/V2.

-- A human-facing title (the uploaded filename, or null → shown as "Untitled").
-- Lets the dashboard list documents by name instead of bare UUIDs.
ALTER TABLE documents ADD COLUMN IF NOT EXISTS title text;

-- Registered 3rd-party webhook endpoints. `secret` signs each delivery
-- (HMAC-SHA256 over the raw JSON body, sent as the X-CW-Signature header).
-- `events` is the set this endpoint subscribes to (today only
-- 'document.session_ended', fired when the last editor disconnects).
CREATE TABLE IF NOT EXISTS webhooks (
    id          uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    url         text   NOT NULL,
    secret      text   NOT NULL,
    events      text[] NOT NULL DEFAULT ARRAY['document.session_ended'],
    active      boolean NOT NULL DEFAULT true,
    created_at  timestamptz NOT NULL DEFAULT now()
);

-- Per-attempt delivery log — powers the dashboard's webhook history and makes
-- retries observable. One row per delivery attempt cluster (attempts counts the
-- tries; status is the terminal outcome).
CREATE TABLE IF NOT EXISTS webhook_deliveries (
    id            uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    webhook_id    uuid NOT NULL REFERENCES webhooks(id) ON DELETE CASCADE,
    doc_id        uuid,
    event         text NOT NULL,
    status        text NOT NULL,            -- 'success' | 'failed'
    response_code int,
    attempts      int  NOT NULL DEFAULT 1,
    last_error    text,
    created_at    timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_webhook_deliveries_webhook
    ON webhook_deliveries (webhook_id, created_at DESC);
