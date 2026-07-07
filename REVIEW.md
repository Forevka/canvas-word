# Track changes + comments — design spec

A **review layer** (tracked insertions/deletions/format changes plus threaded
comments) layered on WordCanvas as an extension. The core document model
(`shared/src/model/`) and the docx import/export stay untouched: no `w:ins` /
`w:del` / `w:rPrChange` / `w:comment*` in the core run stream. The overlay is a
parallel, anchored projection the core knows nothing about. This spec covers the
**data model** and the exact seams it hooks into.

---

## 1. The one design decision everything follows from

Two ways to represent a tracked change:

- **In-stream markers (Word/OOXML).** Wrap runs in `w:ins`/`w:del`; deleted text
  stays in the run list with a "deleted" flag. `CharStyle` would grow
  `trackInsert` / `trackDelete` / `rPrChange`, poisoning `styleEq` /
  `normalizeRuns` and forcing every layout / measure / export pass to be
  review-aware. It is part of the OOXML model. **Rejected by the constraint.**
- **Anchored overlay (this spec).** The core holds *exactly one* concrete text
  state. The review layer is a set of typed, attributed, offset-precise
  **anchors** into it, stored outside `Document` and rebased on every edit by the
  machinery bookmarks already use (`rebaseBookmarks`, `index.ts:637`).

The overlay needs one invariant to make accept/reject trivial:

> **The live core document = "all insertions kept, all deletions still present".**
> The text is the *original* plus every pending insertion, with nothing removed.

Given that, the four operations are bookkeeping:

| Record | Core text today | Accept | Reject |
|---|---|---|---|
| **insertion** | present | drop record | `deleteRange` over anchor + drop record |
| **deletion** | **present** (struck-through, not yet removed) | `deleteRange` over anchor + drop record | drop record |
| **format** | applied | drop record | re-apply inverse patch + drop record |

A tracked deletion removes nothing until accepted: the text stays live and is
only *painted* struck-through. That is what keeps the core review-unaware.

**Cost of this choice:** "deleted" text is still real text, so (a) the
caret/selection can land in it — suggestion-mode input must guard it;
(b) overlapping suggestions need precedence rules; (c) a plain
`serializeDocument` or default docx export contains the *rejected-deletions /
accepted-insertions* state, so callers wanting a clean file must bake first (§7).

---

## 2. Where it lives

```
frontend/src/review/            ← the entire extension; nothing here is imported by core
  model.ts        ReviewLayer types (below)
  intercept.ts    suggestion-mode transaction transformer (pure, headless-testable)
  rebase.ts       rebaseReviewLayer(mapPosition) + GC pass
  resolve.ts      accept/reject → core Transaction + ReviewOp
  decorate.ts     ReviewLayer + LayoutTree → paint decorations
  sync.ts         ReviewSyncClient (separate ws channel / message type)
  persist.ts      serializeReview / parseReview
  ooxmlBridge.ts  OPTIONAL export-time fold into w:ins/w:del/w:comment (§7)
```

The review layer is a **sibling of `doc`** in the editor runtime (`index.ts`),
never a field of `Document`:

```ts
// frontend/src/index.ts — runtime state, alongside `let doc: Document`
let review: ReviewLayer = emptyReview(docId);
```

`serializeDocument` is unchanged. The overlay serializes through its own
`serializeReview` on its own endpoint (§6).

---

## 3. Data model

All types live in `frontend/src/review/model.ts` — the `@forevka/wordcanvas`
namespace, **not** `@cw/shared/model`.

```ts
import type { DocPosition, CharStyle, UserInfo } from "@cw/shared";

/** A range into the LIVE document. start/end are (blockId, offset) positions —
 *  identical shape to BookmarkRange — and ride mapPosition on every edit, so the
 *  range tracks its text through inserts/deletes/splits. start===end is a point
 *  (a caret-anchored comment or an empty insertion stub). May span blocks. */
export interface ReviewAnchor {
  start: DocPosition;
  end: DocPosition;
}

export type SuggestionKind = "insert" | "delete" | "format";

export interface Suggestion {
  id: string;                 // client-minted, globally unique (freshId)
  kind: SuggestionKind;
  anchor: ReviewAnchor;       // covers text PRESENT in the doc (see the invariant)
  author: UserInfo;           // attribution; embedder owns identity (same as Change.userId)
  createdAt: number;          // wall-clock ms; display/sort only

  /** format only: the patch this suggestion applied, plus its exact inverse so
   *  reject can undo it. Mirrors applyStylePatchToRuns semantics. */
  patch?: Partial<CharStyle>;
  inverse?: Partial<CharStyle>;

  /** Optional: groups several anchors authored by one continuous action (typing
   *  a word = one suggestion that grows) so the reviewer accepts/rejects as a
   *  unit and the change-list reads naturally. */
  groupId?: string;
}

export interface Comment {
  id: string;
  author: UserInfo;
  body: string;               // plain text in V1 (rich = a fragment in V2)
  createdAt: number;
  editedAt?: number;
}

export interface CommentThread {
  id: string;
  anchor: ReviewAnchor;       // highlighted span; start===end ⇒ point comment
  status: "open" | "resolved";
  comments: Comment[];        // ordered; comments[0] is the root
}

export interface ReviewLayer {
  docId: string;
  /** The core document version (recorder/SyncClient seq) these anchors were last
   *  rebased to. A late joiner whose doc is newer fast-forwards anchors by
   *  replaying core ops baseVersion..head through mapPosition (§5, §6). */
  baseVersion: number;
  suggestions: Suggestion[];
  threads: CommentThread[];
}
```

**Why anchors and not run flags:** anchors are the only representation the
pipeline already maintains for free. Bookmarks prove a `Record` of `{start,end}`
positions survives arbitrary edits via `mapPosition`; suggestions and threads are
that same shape plus a payload. No new invariant, no core change.

---

## 4. ReviewOp — mutating the overlay

Overlay edits are their own op set (for undo + sync). They target records **by
id** and never carry text offsets overlapping another review op, so any two
ReviewOps commute — OT between them is the identity. Their *anchors* move only
through **core** ops (§5).

```ts
export type ReviewOp =
  | { type: "addSuggestion"; s: Suggestion }
  | { type: "removeSuggestion"; id: string }
  | { type: "growSuggestion"; id: string; end: DocPosition }   // typing extends an insert run
  | { type: "addThread"; t: CommentThread }
  | { type: "addComment"; threadId: string; c: Comment }
  | { type: "editComment"; threadId: string; commentId: string; body: string }
  | { type: "resolveThread"; threadId: string; resolved: boolean }
  | { type: "removeThread"; threadId: string };
```

Each applies to `ReviewLayer` and returns an inverse (for the review undo
channel) — same contract as core `applyOp`, but trivial: record-level
upserts/deletes.

On the wire each ReviewOp is wrapped in `{ op, dependsOnSeq, author }`, where
`dependsOnSeq` is the core document version it was authored against; the receiver
uses it for causal delivery and anchor fast-forward (§5.4). Transport-only, not
persisted in the `ReviewLayer` snapshot.

Concurrency stays simple: core ops move text and rebase anchors; review ops mutate
records and never move text. Neither transforms against the other — that is why
this is an extension, not a fork of the OT core.

---

## 5. Integration seams (named, concrete)

### 5.1 Anchor rebasing — mirror `rebaseBookmarks`
`index.ts` already rebases bookmarks through every applied op. Add the twin:

```ts
// frontend/src/index.ts, next to rebaseBookmarks (index.ts:637)
const rebaseReviewLayer = (mapPosition: (p: DocPosition) => DocPosition): void => {
  review = mapReviewAnchors(review, mapPosition);  // pure, in review/rebase.ts
};
```

Call it in the two places `rebaseBookmarks` is already called:
- `runOps` (`index.ts:658`) — local edits.
- `applyRemoteOps` (`index.ts:739`) — collaborator edits.

That makes every anchor track its text through local typing, undo/redo, and
remote OT edits. After rebasing, run the **GC pass** (`review/rebase.ts`): drop
zero-length deletion suggestions (their text got removed by a real edit), and
relocate comment anchors whose block was removed to the neighbor position
`mapPosition` returned (bookmarks collapse this way on `removeBlock`, ops.ts:543).

### 5.2 Suggestion mode — a transaction transformer
A pure function wraps `commit`. Mode **off** → identity; mode **on** → rewrites the
transaction so destructive edits become overlay records:

```ts
// frontend/src/review/intercept.ts  (headless-testable, like commands)
export function intercept(
  trn: Transaction,
  review: ReviewLayer,
  doc: Document,
  author: UserInfo,
): { core: Transaction; reviewOps: ReviewOp[] };
```

Per-op rewrite rules:

- **insertText / insertRuns** → applied unchanged; emit `addSuggestion{insert}`
  over the inserted range (or `growSuggestion` if the caret abuts the author's own
  open insert run → contiguous typing is one record).
- **deleteRange** → **not** applied as a destructive delete. Partition the range:
  - sub-ranges inside *this author's own pending insertion* → really deleted
    (`deleteRange` kept in `core`) and that insertion record trimmed (you can
    hard-delete what was never original);
  - sub-ranges of original / others' text → **no core op**; emit
    `addSuggestion{delete}` over them. Caret moves to range start; text stays.
- **setParaStyle / char-style (setRuns over a range)** → applied; emit
  `addSuggestion{format, patch, inverse}`.
- **splitParagraph / mergeParagraphs / insertBlock / removeBlock /
  insert·removeTableRow / insert·removeTableColumn** → applied unchanged (doc
  stays live) and emit `addSuggestion{structural, structural:{op, inverse,
  blockId}}`. The exact inverse is harvested from `applyOp` at intercept time.
  Accept drops the record; reject re-applies the stored inverse. The record hangs
  on `blockId` (the created/merged/removed block, or the host table) and is GC'd
  when that block no longer exists. Paint marks it with a block-level change-bar
  (no text range to highlight).

**Multi-op transactions (paste, cross-paragraph delete)** are decomposed: each op
routes through the SAME per-op logic against a *running* doc + review, and all
records from one transaction share a `groupId` so the whole action
accepts/rejects as a unit (cases in §8). Coordinate drift is handled as the commit
pipeline does: ops apply in sequence, and each already-emitted record's anchor
rebases forward through the later ops' position-mappers. A record is never rebased
through the op that created it — its anchor is already in that op's
post-coordinates. `acceptAll`/`rejectAll` resolve against a running doc too, so a
structural reject that re-merges/re-inserts whole blocks composes with the text
rejects whose anchors live on them.

The runtime applies `core` through the normal `commit` path (undoable, recorded,
synced as today) and applies `reviewOps` to the sidecar + broadcasts them (§5.4).

### 5.3 Paint decorations — mirror the search-rect channel
`afterMutation` already recomputes search highlights into a paint-only rect
channel (`runSearch` / `paintSearch`, `index.ts:670`). Add the same shape:

```ts
// computed in afterMutation, same spot as paintSearch
paint.setReviewDecorations(decorate(review, tree, scope()));
```

`decorate()` (`review/decorate.ts`) walks the layout tree once and emits:
- **insertion** → colored underline on covered fragments (author color);
- **deletion** → colored strikethrough on covered fragments;
- **comment** → highlight band under the span + a margin gutter pin;
- **change bar** → a vertical rule in the margin for any line a suggestion touches.

These are **paint-only** and metric-neutral (underline/strike/highlight don't
change line breaking), so a suggestion triggers repaint, not relayout — like
search highlights. Measurement, pagination, export are untouched.

### 5.4 Sync — a second, independent channel
Core ops sync via `SyncClient` (Jupiter/ShareDB). ReviewOps ride a parallel
channel — a new WebSocket message `type:"review"` on the same socket, or a small
`ReviewSyncClient`. Because ReviewOps commute (§4), the server totally-orders and
broadcasts them with no transform. Inbound **core** remote ops already rebase local
anchors via `rebaseReviewLayer` in `applyRemoteOps` (§5.1), so a remote insert
shifts everyone's comment pins.

The one non-optional requirement: the two channels are not atomically co-ordered,
so an inbound ReviewOp's *anchor* is causally dependent on the core op it was
authored with. Hence `dependsOnSeq` on each ReviewOp:

- **Causal hold.** Apply an inbound ReviewOp only once the document reaches
  `dependsOnSeq`; otherwise buffer it (a few ms in practice). Prevents an
  `addSuggestion` whose anchor points at not-yet-arrived text.
- **Anchor fast-forward.** When the receiver is *past* `dependsOnSeq` (it applied
  other core ops first), map the inbound anchor forward through core ops
  `dependsOnSeq..head` before inserting the record — the batch form of §5.1, the
  same routine late-join uses (§6). After that the anchor rebases normally.

**Concurrent resolution is convergent.** `removeSuggestion(id)` is id-keyed and
idempotent. Two reviewers accepting the same *deletion* both emit `deleteRange`
over the same range; OT collapses the second to empty (`transformOp` returns `[]`
for a fully-swallowed delete, transform.ts), so the text is removed once. For
**accept vs reject** racing the same record, the server's total order decides:
whichever lands first wins, and the later one finds the record gone and no-ops.
The core text op (if any) is the real authority.

### 5.5 Accept / reject — back to core ops
`review/resolve.ts` turns a reviewer action into a normal core `Transaction`
(undoable, synced, attributed) plus a `ReviewOp`, per the §1 table:

```ts
acceptSuggestion(id):  insert→[removeSuggestion]
                       delete→[core deleteRange(anchor)] + removeSuggestion
                       format→[removeSuggestion]
rejectSuggestion(id):  insert→[core deleteRange(anchor)] + removeSuggestion
                       delete→[removeSuggestion]
                       format→[core applyStylePatch(inverse, anchor)] + removeSuggestion
```

`acceptAll` / `rejectAll` fold **back-to-front by document order** (the discipline
replace-all uses, `ROADMAP.md` find/replace) so earlier anchors stay valid as
later ranges collapse. Anchors auto-rebase after each op, so order affects
determinism, not correctness.

### 5.6 Editing modes — the edit / suggest / view switch
Mode gates three already-existing behaviors. It is **session/UI state, not
document content**: it lives next to `readonly`, is never written to `Document` or
`ReviewLayer`, and is never synced (a local preference, like zoom).

```ts
export type EditMode = "edit" | "suggest" | "view";
// "view" === today's readonly:true. The three are mutually exclusive.
```

| Mode | `commit` (index.ts:685) | `intercept` (§5.2) | Editing chrome |
|---|---|---|---|
| `edit` | applies ops destructively | identity (passthrough) | shown |
| `suggest` | applies the rewritten ops | rewrites to overlay records | shown + a "Suggesting" affordance |
| `view` | **no-op** (drops every mutation) | — | hidden |

Adding `suggest` is a two-line change to the dispatch path: the `readonly` gate
already proves a mode can neutralize `commit`; `suggest` reuses it by routing the
transaction through `intercept` first. Remote ops (`applyRemoteOps`) bypass mode
entirely — a viewer/suggester still receives live peer edits, as readonly does
today.

Switching mid-session needs no migration: `suggest → edit` leaves suggestions
pending (no auto-accept); `edit → suggest` tracks the next edit. Editing over
suggestion-anchored text in `edit` mode is fine — destructive ops rebase anchors
(§5.1) and GC drops any record whose text was consumed. On every switch, clear
`pendingStyle` and any open suggestion-grow run so the next edit starts fresh.

**Per-user, embedder-lockable.** Mode is per *session*, not per document — Alice
can edit while Bob (external reviewer) suggests on the same live doc. Because the
embedder owns auth, expose a lock to *force* a mode and hide the switch:

```ts
new WordCanvas({
  mode?: EditMode,                    // initial mode (default "edit"; "view" === readonly:true)
  allowedModes?: EditMode[],          // omit ⇒ all three; e.g. ["suggest","view"] locks out raw editing
});
```

When `allowedModes` excludes the current `mode`, the picker is hidden and
`setMode` to a disallowed value returns false. This embeds "external collaborators
can only suggest" without trusting the client: the server still authorizes every
change, and the UI won't offer edit mode.

**Presence (optional, cosmetic).** The presence channel may carry each peer's mode
so the roster reads "Alice — editing / Bob — suggesting". Display only.

**UI.** A toolbar mode dropdown (Editing / Suggesting / Viewing), the Google-Docs
pattern. `view` reuses the existing readonly chrome-hiding path.

---

## 6. Persistence & late join

- Stored on its **own** endpoint, e.g. `GET/PUT /docs/:id/review` (snapshot) or an
  append log `POST /docs/:id/review/ops`. **Never** part of the `/docs` document
  snapshot, so `serializeDocument` and the docx pipeline stay clean.
- The backend already reuses the shared model and replays the change log; the
  review store sits beside it. The existing **session-end webhook** can include
  `{ review }` so downstream systems get suggestions + comments.
- **Late join / version skew:** the loader fetches the doc at `head` and the
  review layer at its `baseVersion`. If `review.baseVersion < head`, replay core
  ops `baseVersion..head`, fold each op's `mapPosition` over the anchors (batch
  form of §5.1), then set `baseVersion = head`. Same machinery, run once at load.

---

## 7. The OOXML boundary

- **Untouched:** `shared/src/model/*`, `frontend/src/import/docx/*`,
  `frontend/src/export/docx/writeDocx.ts`. No review concepts enter `CharStyle`,
  `styleEq`, `normalizeRuns`, the importer prop tables, or `writeDocx`.
- **Default export bakes.** Because the live model carries pending deletions as
  real text, `exportDocx()` first produces a clean state — `acceptAll` (or
  `rejectAll`, caller's choice) applied to a *copy* — then runs the existing
  writer. A plain consumer gets a plain document.
- **Optional interop adapter** `review/ooxmlBridge.ts`: when a caller explicitly
  wants Word-native redlines, it folds the overlay into `w:ins` / `w:del` /
  `w:rPrChange` / `word/comments.xml` **on top of** the writer's output (post-
  processing the emitted parts), never inside it. Full Word interop stays opt-in
  while the core writer remains review-blind. The import side (reading a docx that
  *already* has tracked changes) is the mirror adapter and is V2.

---

## 8. Scope & honest hard cases

- **Structural suggestions** — tracked paragraph split/merge, block add/remove,
  table row/column add/remove (§5.2). A `structural` record carries the applied
  `op` + its exact `inverse` (from `applyOp` at intercept time) + the `blockId` it
  hangs on; accept drops the record, reject re-applies the inverse. It is
  block-keyed (its degenerate point anchor still rides `mapPosition`) and GC'd by
  block existence (`gcStructuralReviewLayer`), not by anchor collapse.
  `acceptAll`/`rejectAll` resolve structural records newest-first (reverse
  application order — a paste's split then insert unwinds correctly) before the
  positionally-folded text records, against a running doc so structural and text
  rejects compose.
- **Tracked paste & cross-paragraph delete** — decomposed per-op, bundled under
  one `groupId` (§5.2). A multi-block paste emits insert + structural records for
  all pasted content; a cross-paragraph delete becomes tracked delete + merge
  records. Reject reverses the whole action.

**Deferred / called out:**
- **Overlapping suggestions** (a delete inside an insert; a format over a delete).
  V1 precedence: an insertion fully inside a deletion is hard-removed on
  accept-delete; a format over deleted text is dropped when the deletion is
  accepted. Document the table; complex partial overlaps are V2.
- **Editing inside a deletion range** in suggestion mode — guarded: typing splits
  the deletion record around the caret and starts a fresh insertion (Word
  behavior). Worth a dedicated test matrix.
- **Review undo.** Suggestion creation and accept/reject are real core
  transactions (already undoable) paired with ReviewOps; ReviewOps carry inverses
  and join a **separate review undo channel** so Ctrl+Z over a text edit and over
  a "resolve thread" don't interleave confusingly. Revisit if users want unified
  undo.
- **Anchor GC** after block removal — specified in §5.1; needs the same
  neighbor-collapse semantics bookmarks use, plus dropping emptied deletion
  records.

---

## 9. Public API (`@forevka/wordcanvas`)

```ts
new WordCanvas({
  …,
  mode?: EditMode,                   // "edit" (default) | "suggest" | "view"
  allowedModes?: EditMode[],         // lock the picker (§5.6); omit ⇒ all three
  // readonly?: true                 // back-compat alias for mode:"view"
});

ed.setMode(mode: EditMode): boolean;               // false if disallowed (§5.6)
ed.getMode(): EditMode;
// ed.setSuggestionMode(on) — back-compat alias for setMode("suggest"|"edit")

ed.getReview(): ReviewLayer;                       // read-only snapshot
ed.acceptSuggestion(id) / rejectSuggestion(id): void;
ed.acceptAllSuggestions() / rejectAllSuggestions(): void;
ed.addComment(body: string): string;               // anchors to current selection; returns threadId
ed.replyToComment(threadId, body): void;
ed.resolveThread(threadId, resolved = true): void;
ed.exportDocx({ tracked?: boolean }): Promise<Blob>; // false (default) bakes; true uses ooxmlBridge

// events (added to WordCanvasEventMap)
ed.on("suggestionAdded",  ({ suggestion }) => …);
ed.on("suggestionResolved", ({ id, action }) => …); // "accept" | "reject"
ed.on("commentAdded",     ({ threadId, comment }) => …);
ed.on("threadResolved",   ({ threadId }) => …);
ed.on("modeChanged",      ({ mode }) => …);         // edit | suggest | view
```

All of it is additive: existing constructor options, events, and the core
document model are unchanged.
