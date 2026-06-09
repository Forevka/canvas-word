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

// Collaboration foundations: unique ids, content-addressed media, document <->
// snapshot serialization, the change log, and replay.
export * from "./ids";
export * from "./persist/media";
export * from "./persist/serialize";
export * from "./change";
export * from "./replay";
export * from "./transform";
