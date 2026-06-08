-- Canvas-word collaboration store.
--
-- The history of a document is one base snapshot plus an append-only, totally
-- ordered change log; any version is reconstructed by replaying the log over the
-- snapshot (see shared/replay). Image bytes are content-addressed so they are
-- stored once and referenced by hash. Idempotent (IF NOT EXISTS) so Flyway can
-- repair + re-run safely.

-- One row per collaboratively-edited document. head_version is the seq of the
-- latest accepted change (0 = only the base snapshot exists).
CREATE TABLE IF NOT EXISTS documents (
    id            uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    head_version  bigint NOT NULL DEFAULT 0,
    created_at    timestamptz NOT NULL DEFAULT now()
);

-- Base snapshot (version 0) and periodic checkpoints. A checkpoint is the doc
-- materialized at a given version so clients/servers don't replay from 0 forever.
CREATE TABLE IF NOT EXISTS snapshots (
    doc_id      uuid   NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
    version     bigint NOT NULL,
    doc         jsonb  NOT NULL,
    created_at  timestamptz NOT NULL DEFAULT now(),
    PRIMARY KEY (doc_id, version)
);

-- The append-only ordered change log. seq is the canonical historical order
-- (server-assigned); change_id is the client's idempotency key (a re-sent change
-- must not double-apply). ops/selection are the Change payload as JSON.
CREATE TABLE IF NOT EXISTS changes (
    doc_id        uuid   NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
    seq           bigint NOT NULL,
    change_id     text   NOT NULL,
    base_version  bigint NOT NULL,
    site_id       text   NOT NULL,
    origin        text   NOT NULL,
    ts            bigint NOT NULL,  -- author wall-clock, epoch ms (matches Change.ts)
    ops           jsonb  NOT NULL,
    selection     jsonb,
    PRIMARY KEY (doc_id, seq),
    UNIQUE (doc_id, change_id)
);

-- Content-addressed media: hash = sha256(bytes). One row per distinct image,
-- referenced by ImageBlock.mediaId. bytea now; can move to disk/object storage
-- behind the same accessor later.
CREATE TABLE IF NOT EXISTS media (
    hash        text PRIMARY KEY,
    mime        text NOT NULL,
    bytes       bytea NOT NULL,
    created_at  timestamptz NOT NULL DEFAULT now()
);
