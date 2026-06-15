// @cw/shared — the pure, DOM-free, Node-free document core shared by the editor
// frontend and the collaboration backend. Everything here is plain data + logic:
// the document model, positions, the operation log (applyOp), and the text
// helpers. No canvas, no DOM, no Node APIs — so both the browser editor and the
// server can import and run it (replay, transform, serialize).

export * from "./model/document";
export * from "./model/position";
export * from "./model/text";
export * from "./model/stylesheet";
export * from "./model/lists";
export * from "./model/ops";
export * from "./model/tableGrid";

// Collaboration foundations: unique ids, content-addressed media, document <->
// snapshot serialization, the change log, and replay.
export * from "./ids";
export * from "./persist/media";
export * from "./persist/serialize";
export * from "./change";
export * from "./toc";
export * from "./fields";
export * from "./replay";
export * from "./transform";

// Review layer (track changes + comments) — an isomorphic OVERLAY extension. A
// sibling of ./model, never imported by it, so the OOXML-faithful core stays
// unaware of review. See shared/src/review and REVIEW.md.
export * from "./review";
