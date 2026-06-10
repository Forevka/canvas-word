-- 3rd-party integration tokens (API keys) for machine-to-machine /upload.
-- Only the SHA-256 hash of the token is stored; the plaintext `cwk_…` value is
-- shown once at creation and never again. Idempotent for Flyway repair/re-run.
CREATE TABLE IF NOT EXISTS api_tokens (
    id           uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    name         text NOT NULL,
    token_hash   text NOT NULL UNIQUE,   -- sha256(token) hex
    prefix       text NOT NULL,          -- first chars, for display ("cwk_1a2b…")
    active       boolean NOT NULL DEFAULT true,
    created_at   timestamptz NOT NULL DEFAULT now(),
    last_used_at timestamptz
);
