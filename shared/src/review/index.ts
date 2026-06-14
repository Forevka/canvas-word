// @cw/shared/review — the isomorphic core of the track-changes + comments
// extension: data model, record ops, anchor rebasing, serialization. Pure data +
// logic, reused by both the editor frontend and the collaboration backend. It is
// a SIBLING of shared/src/model (never imported by it) so the OOXML-faithful core
// stays entirely unaware of review. See REVIEW.md.

export * from "./model";
export * from "./ops";
export * from "./rebase";
export * from "./serialize";
export * from "./resolve";
export * from "./bake";
export * from "./rehydrate";
