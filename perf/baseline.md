# Performance baseline — perf/incremental-layout-and-resize

Captured against a real client report opened in the dev editor (port 5175),
measured via `window.__cw.engine.layout(doc)` (the exact function `relayout()`
calls on every mutation). CPU throttling 1x. Numbers are ms, median of 21 runs
after 4 warmup runs.

## Test document
`SignedReports_…-7407 Igou Gap Rd…docx`

| metric | value |
|---|---|
| pages | 98 |
| top-level blocks | 849 |
| paragraphs | 1961 (1193 in table cells) |
| images | 60 (20 in table cells) |
| tables | 41 (8 autofit, 33 fixed), 1191 cells, max 33 rows |
| nested tables | 0 |

## Baseline relayout timings (ms)

| scenario | min | median | p90 | notes |
|---|---|---|---|---|
| **warm relayout** (no change) | 33.0 | **38.0** | 46.2 | per-mutation floor — runs once per edit |
| keystroke (1 para dirtied) | 30.5 | 33.9 | 38.2 | ≈ warm; the single re-break is cheap |
| in-cell image resize (fixed table) | 29.9 | 33.1 | 38.9 | per drag frame |
| top-level image resize | 32.3 | 37.5 | 47.0 | per drag frame |
| cold (engine.reset + layout) | 138.4 | 141.5 | 159.0 | full reshape / first load |

## Attribution (sizing item 2 vs 3)

| variant | median ms |
|---|---|
| warm full | 33.1 |
| warm, all 41 tables collapsed to a placeholder paragraph | 24.2 |

→ **table measurement ≈ 9 ms (~27%)** of a warm relayout — the ceiling for
item 2 (table measure cache). The remaining **~24 ms is pure LayoutTree
rebuild / pagination walk** over ~1961 already-cached paragraphs across 98
pages (allocation + placement, not text shaping) — the target for item 3
(incremental relayout that reuses clean-block geometry).

## Headline

A single relayout costs **~33–38 ms** on this 98-page report, and there is
**one relayout per mutation**. That is ~2× the 16.6 ms/60fps frame budget, which
is exactly the resize/keystroke lag. The RAF fix caps resize to one relayout per
frame (~30fps ceiling); the structural fixes below remove the cost itself:

- **Item 1** (overlay-preview resize): per-frame relayout → 0; pay one ~33 ms
  relayout only on mouseup commit.
- **Item 2** (table measure cache): warm relayout ~38 → ~29 ms (reclaim ~9 ms).
- **Item 3** (incremental relayout): keystroke/edit cost target well below the
  ~24 ms walk by reusing clean-block placement.
- **Item 4** (slim transient `afterMutation`): fewer per-frame side effects.
- **Item 5** (per-page paint diffing): repaint only changed live pages.

Re-run the identical `engine.layout` micro-benchmark after each item to compare.
Raw DevTools trace of a 30-frame resize burst: `perf/baseline-resize-burst-trace.json`.

---

## Results after implementing items 1–5

Same document, same `engine.layout` micro-benchmark, CPU 1x, median of 21 runs.

| scenario | baseline | final | speedup |
|---|---|---|---|
| **warm relayout** (per-mutation floor) | 38.0 ms | **4.5 ms** | 8.4× |
| keystroke (1 paragraph dirtied) | 33.9 ms | **3.8 ms** | 8.9× |
| in-cell image resize (per drag frame) | 33.1 ms | **0 relayouts** | n/a (item 1) |
| top-level image resize (per drag frame) | 37.5 ms | **0 relayouts** | n/a (item 1) |
| cold (first load / full reshape) | 141.5 ms | 122.6 ms | 1.15× |

### What each item contributed (measured)

- **Item 1 — overlay-preview resize:** a resize drag does **zero** relayouts (was
  one ~33 ms relayout per frame). Verified live: model unchanged during drag, one
  undoable op on release, aspect preserved.
- **Item 2 — table measure cache:** warm relayout 38 → 23 ms (reclaimed the ~9–15
  ms of per-pass table re-measurement).
- **Item 3 — per-page band cache:** warm 23 → 4.5 ms. Profiling found band
  (header/footer) layout was 20.2 ms of the remaining 24 ms — 196 `layoutBand`
  calls re-shaped every pass; caching their pure output collapsed it to 0.18 ms.
  (The planned dirty-id incremental *pagination* was dropped: the pagination walk
  is only ~2 ms, not worth the regression risk.)
- **Item 4 — slim transient afterMutation:** defers search / peer-caret /
  review-overlay re-measurement during drag & IME frames to the commit that ends
  them.
- **Item 5 — per-page paint diffing:** repaint only the live pages whose content
  changed (signature cost 0.01–0.23 ms/page vs a multi-ms repaint).

Phase breakdown of a warm relayout, before → after item 3:
`{ measure 0.9, walk 2.4, band 20.2 }` → `{ measure 0.9, walk 2.4, band 0.18 }`.
