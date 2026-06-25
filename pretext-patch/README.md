# `/pretext-patch` — local patches to `@chenglou/pretext`

The editor's text layout engine is built on [`@chenglou/pretext`](https://github.com/chenglou/pretext)
(see `ARCHITECTURE.md` §2 — pretext is confined to `frontend/src/layout/`). pretext
already ships a complete-enough Unicode toolbox (CJK line-breaking, kinsoku, and a
UAX#9 bidi implementation), but a couple of those capabilities aren't reachable or
configurable through its **published** API. Rather than re-implement the Unicode
Bidi Algorithm in this repo, we patch pretext itself so the engine can consume it,
and persist the diff here with [`patch-package`](https://www.npmjs.com/package/patch-package).

## How it's applied

- The patch file `@chenglou+pretext+0.0.7.patch` is applied by the **root**
  `postinstall` script: `patch-package --patch-dir pretext-patch`.
- It is version-pinned to `0.0.7`. If pretext is bumped, `patch-package` will warn
  that the patch no longer applies cleanly — regenerate it (see below) against the
  new version and re-verify the bidi tests.
- This is a **temporary divergence**. Both changes are additive and backward-compatible;
  the goal is to upstream them to chenglou/pretext so this directory can be deleted.

## What the patch changes (and why)

### 1. Expose the `analysis` and `bidi` subpaths (`package.json` `exports`)

pretext only exports `.` (the `layout` entry) and `./rich-inline`. The editor needs
`./analysis` (for `isCJK`, `setAnalysisLocale`, the `kinsoku*` sets) and `./bidi`
(for `computeSegmentLevels`). Both modules already exist in `dist/` — the patch just
adds them to the `exports` map (with their `.d.ts` types) so `@chenglou/pretext/bidi`
and `@chenglou/pretext/analysis` resolve in TypeScript and at runtime instead of
forcing fragile deep `node_modules/.../dist/*.js` imports that wouldn't survive
bundling of `@forevka/wordcanvas`.

### 2. Optional explicit base direction for `computeSegmentLevels`

pretext's `bidi.js` picks the paragraph base direction from the **first strong
character** (UBA P2/P3). Word's `w:bidi` paragraph property instead forces the base
direction to RTL regardless of content (an RTL paragraph that happens to start with
a Latin word is still laid out RTL). The patch adds an optional third argument:

```ts
computeSegmentLevels(normalized: string, segStarts: number[], baseLevel?: 0 | 1): Int8Array | null
```

When `baseLevel` is `0` or `1` the first-strong-char scan is skipped and that level
is used as the paragraph embedding base; when it's omitted the original P2/P3
behavior is preserved exactly. No other resolution step (W/N/I rules) changes.

## Regenerating the patch (after editing `node_modules/@chenglou/pretext`)

`package.json` is excluded from `patch-package` diffs by default, so pass an
exclude regex that matches nothing to capture the `exports` change too:

```sh
npx patch-package "@chenglou/pretext" --patch-dir pretext-patch --exclude "^xyzzy-never-match$"
```

Then run the bidi/CJK layout tests (`npm run test --workspace @forevka/wordcanvas`)
to confirm the rebuilt patch still produces correct reordering.
