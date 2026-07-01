// Compile-time parity guard for the hand-written published declarations in
// `types/query.d.ts`. That `.d.ts` mirrors runtime contracts by hand
// (frontend/src/query.ts, frontend/src/layout/pages.ts, and — via re-export —
// shared/src/model/documentEditor.ts + query.ts). This file asserts the runtime
// surface stays assignable to the published surface, so `tsc --noEmit`
// (`npm run typecheck`, run in CI) fails if either side drifts.
//
// Type-only: no runtime output, no test framework. Not a vite entry, not a
// vitest spec — it exists purely so the compiler checks it. The single exported
// tuple keeps every assertion "used" (no noUnusedLocals noise) while forcing each
// element to evaluate to `true`.

import type * as Rt from "./query";
import type * as Pub from "../types/query";

/** Fails to compile unless `T` is exactly `true`. */
type Assert<T extends true> = T;
/** True when type `A` is assignable to type `B`. Tuple-wrapped so a union `A`/`B`
 *  is compared as a whole instead of distributing to `boolean`. */
type AssignableTo<A, B> = [A] extends [B] ? true : false;

type RtEditor = InstanceType<typeof Rt.DocumentEditor>;
type PubEditor = InstanceType<typeof Pub.DocumentEditor>;

/** Each entry must be `true`; drift on either side makes an element `false` and
 *  breaks compilation. Functions: the runtime signature must satisfy the
 *  published one. Value shapes are checked in the direction that matters. */
export type QueryPublicSurfaceParity = [
  Assert<AssignableTo<typeof Rt.walk, typeof Pub.walk>>,
  Assert<AssignableTo<typeof Rt.textOf, typeof Pub.textOf>>,
  Assert<AssignableTo<typeof Rt.getParagraphs, typeof Pub.getParagraphs>>,
  Assert<AssignableTo<typeof Rt.getTables, typeof Pub.getTables>>,
  Assert<AssignableTo<typeof Rt.getImages, typeof Pub.getImages>>,
  Assert<AssignableTo<typeof Rt.findParagraphs, typeof Pub.findParagraphs>>,
  Assert<AssignableTo<typeof Rt.getBlockById, typeof Pub.getBlockById>>,
  Assert<AssignableTo<typeof Rt.getParagraphById, typeof Pub.getParagraphById>>,
  Assert<AssignableTo<typeof Rt.getTableById, typeof Pub.getTableById>>,
  Assert<AssignableTo<typeof Rt.getImageById, typeof Pub.getImageById>>,
  Assert<AssignableTo<typeof Rt.getSections, typeof Pub.getSections>>,
  Assert<AssignableTo<typeof Rt.resolveSections, typeof Pub.resolveSections>>,
  // Content controls (SDTs).
  Assert<AssignableTo<typeof Rt.getSdt, typeof Pub.getSdt>>,
  Assert<AssignableTo<typeof Rt.getSdts, typeof Pub.getSdts>>,
  Assert<AssignableTo<typeof Rt.getSdtsByTag, typeof Pub.getSdtsByTag>>,
  Assert<AssignableTo<typeof Rt.getSdtsByAlias, typeof Pub.getSdtsByAlias>>,
  Assert<AssignableTo<typeof Rt.getSdtNodes, typeof Pub.getSdtNodes>>,
  Assert<AssignableTo<typeof Rt.getSdtRoots, typeof Pub.getSdtRoots>>,
  Assert<AssignableTo<typeof Rt.getSdtChildren, typeof Pub.getSdtChildren>>,
  Assert<AssignableTo<typeof Rt.getSdtAncestors, typeof Pub.getSdtAncestors>>,
  Assert<AssignableTo<typeof Rt.getSdtDescendants, typeof Pub.getSdtDescendants>>,
  Assert<AssignableTo<typeof Rt.getSdtBlocks, typeof Pub.getSdtBlocks>>,
  Assert<AssignableTo<typeof Rt.sdtText, typeof Pub.sdtText>>,
  Assert<AssignableTo<Rt.SdtMatch, Pub.SdtMatch>>,
  Assert<AssignableTo<Rt.SdtNode, Pub.SdtNode>>,
  Assert<AssignableTo<typeof Rt.getPages, typeof Pub.getPages>>,
  Assert<AssignableTo<typeof Rt.pageOfBlock, typeof Pub.pageOfBlock>>,
  // A runtime DocumentEditor instance must satisfy the published instance surface.
  Assert<AssignableTo<RtEditor, PubEditor>>,
  // Shared value-type shapes.
  Assert<AssignableTo<Rt.PageInfo, Pub.PageInfo>>,
  Assert<AssignableTo<Pub.WalkOptions, Rt.WalkOptions>>,
  Assert<AssignableTo<Rt.ParagraphMatch, Pub.ParagraphMatch>>,
  Assert<AssignableTo<Pub.InsertParagraphOptions, Rt.InsertParagraphOptions>>,
];
