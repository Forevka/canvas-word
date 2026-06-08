// Globally-unique id minting for blocks, cells, and other model nodes.
//
// In single-user mode any unique string works. For multi-user collaboration two
// clients must NEVER mint the same id (a collision would make one client's
// paragraph silently alias another's), so ids are namespaced by a per-session
// `siteId`:
//
//   <siteId>-<counter base36>      e.g.  "a3f9c1-0", "a3f9c1-1", …
//
// The counter is monotonic within a session; the siteId makes the whole space
// disjoint from every other session. Replaying an op stream is deterministic
// because generated ids are captured IN the op (splitParagraph.newBlockId,
// insertBlock.block.id), so the consumer never re-mints them.
//
// Randomness lives at the edges (the frontend mints a siteId via crypto), not
// here — this module stays pure and deterministic so it is trivially testable.

export interface IdGenerator {
  readonly siteId: string;
  /** Next unique id in this site's space. */
  next(): string;
  /** Current counter value (for persisting/resuming a session's id space). */
  count(): number;
}

export function createIdGenerator(siteId: string, start = 0): IdGenerator {
  let n = start;
  return {
    siteId,
    next: () => `${siteId}-${(n++).toString(36)}`,
    count: () => n,
  };
}

// A process-wide default generator. The frontend reconfigures it at startup with
// a real siteId (see configureIds); tests and any unconfigured context get the
// deterministic "local" space.
let active: IdGenerator = createIdGenerator("local");

/** Point the default generator at a session's siteId. Call once at startup,
 *  before any ids are minted. `start` resumes a prior session's counter. */
export function configureIds(siteId: string, start = 0): void {
  active = createIdGenerator(siteId, start);
}

/** Mint a fresh unique id from the active generator. */
export function freshId(): string {
  return active.next();
}

/** The active session's siteId (tags Changes so the server can order them). */
export function currentSiteId(): string {
  return active.siteId;
}
