-- Review layer (track changes + comments) persistence. An append-only log of
-- review ops, SEPARATE from the document snapshot and change log — the docx
-- pipeline and serializeDocument stay review-blind. Folding the log through
-- applyReviewOp reconstructs the overlay (see shared/src/review).

CREATE TABLE review_ops (
    id             bigserial PRIMARY KEY,
    doc_id         uuid   NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
    op             jsonb  NOT NULL,            -- the ReviewOp (addSuggestion, removeSuggestion, addThread, …)
    depends_on_seq bigint NOT NULL,            -- core version the op's anchor depends on (causal delivery)
    author_id      text,                       -- attribution (NULL for anonymous)
    created_at     timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX idx_review_ops_doc ON review_ops (doc_id, id);
