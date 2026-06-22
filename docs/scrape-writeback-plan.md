# Report scrape / write-back via nested content controls

> Design doc. The diff/scrape engine lives C#-side in AppraiSys; canvas-word's role is **fidelity** — preserve nested content controls through import → edit → export. This file mirrors the working plan and is kept in-repo for the canvas-word fidelity work + test fixtures.

## Problem

AppraiSys renders `.docx` appraisal reports by injecting Tasks-system data into Word content controls (`w:sdt`). A user edits a value in the rendered report and saves; we must scrape the edited docx and push each changed value back to its originating store. Today the document carries **no field-level identity** — section SDTs only tag `{Name, Hash}`, substitutors emit whole sections from many values, and provenance is lost before OOXML emit.

## Approach

Make every editable value **self-describing** by wrapping it in its own inline content control whose `w:tag` carries a `FieldBinding` JSON naming the store + key + original raw value. Keep the existing **section-level** SDT as the outer wrapper; nest **per-field** SDTs inside it. The scrape side is one generic walk + a pluggable write-back strategy registry keyed by a `target` discriminator in the JSON.

### `FieldBinding` tag schema

```jsonc
{
  "k": "fb",                  // discriminator; section SDTs lack "k" → skipped
  "target": "formFieldValue", // "appraiseRequest" | "formFieldValue" | "formSubmission" (the strategy switch)
  "type": "currency",         // value format → drives raw normalization
  "orig": "200",              // ORIGINAL RAW value (not formatted) → oldValue; self-contained diff
  "storyId": 42,              // formFieldValue / formSubmission
  "fieldName": "executive_summary_appraisal_fee",
  "arKey": null,              // appraiseRequest: key into write-back registry
  "itemKey": null             // reserved for deferred collection targets
}
```

`orig` is the **raw** DB value (cell shows `"$200.00"`, column holds `"200"`); scrape normalizes edited text back to raw before diff/upsert.

### V1 write-back strategies (`IScrapeWriteBackStrategy`, keyed DI)

| target | primitive | notes |
|---|---|---|
| `formFieldValue` | `BulkUpsertFormSubmissionFieldValues` | lean atomic upsert on `(AppraiseRequestId, StoryId, FieldName)` |
| `formSubmission` | `UpsertFormSubmissionFieldValues(...)` | full: new `FormSubmission` audit row + `SubmissionData` JSON + mapped fields + `StoryDataUpdatedEvent` |
| `appraiseRequest` | new narrow DAL method + key→column registry | invertible scalars only (FileNumber, ClientOrderNumber, ClientReference, LoanNumber, DueDate, DateOfValue, PropertySize, RestrictedAppraisalReport, PropertyLocation.{AddressLine1,2,City,County,ZipCode}); computed getters like FullAddressLine are NOT wrapped |

Comparables + others are deferred — they slot in as additional strategies, no engine change.

## canvas-word responsibilities (this repo)

1. **Fidelity:** nested binding SDTs survive import → edit → export with tag JSON intact. **(Done — nesting is now first-class; see `docs/nested-sdt-plan.md`.)**
   - Membership is an ordered ancestry path: `CharStyle.sdtPath` (inline) + `Block.sdtPath` (block-level). Import builds it via push/pop stacks (`documentParser.ts`); export reconstructs nested `w:sdt` by LCP grouping (`documentXml.ts`).
   - The earlier flattening risk is resolved: an outer section SDT containing inner field SDTs round-trips as a single control with the inner controls nested, verified by `test-fixtures/scrape/roundtrip-test.mjs`.
2. (Later, with the Syncfusion swap) editor UX: read-only outside bindings, inspector showing `orig → current`, provisional `itemKey` on row insert. Hook points: `frontend/src/editor/commands.ts` (`sdtAtPosition`), `frontend/src/ui/sdtInspector.ts`.

## AppraiSys responsibilities (other repo; monolith today, moving to a reporting microservice)

- `WrapBinding` OOXML helper near `LazyContentTableHelper`; thread provenance `(source, key, storyId, raw)` out to each substitutor's emit point (the binder drops the internal name today).
- `ReportDocxScrapeService` (DAL-only / relocatable): `Descendants<SdtElement>()` → parse `k:"fb"` tags → diff raw values → dispatch by `target`.
- Inbound HMAC webhook receiver (`Controllers/Integration/`, modeled on `ComparablesWebhookController`) + outbound typed `HttpClient` to fetch `GET /docs/:id/export.docx` from canvas-word; correlate via `AppraiseRequestReport.ExternalIdentifier`.

## Test fixture

`test-fixtures/scrape/make-fixture.ps1` hand-authors `scrape-sample.docx`: an outer section SDT containing inner field SDTs (one in a paragraph, three in table value cells covering all three targets) plus one non-nested inline SDT for comparison. `test-fixtures/scrape/roundtrip-test.mjs` imports it via the lib's node build, prints the SDT registry + per-run membership, exports, re-imports, then edits a value and round-trips again.

## Fidelity test RESULTS (measured 2026-06-21)

Run: `node test-fixtures/scrape/roundtrip-test.mjs` (lib `dist-node`). Findings:

1. **Per-field binding controls survive import → edit → export → re-import with full fidelity.** All 5 `k:"fb"` controls round-trip with `target`/`fieldName`/`orig` tags intact; an edit ($200.00 → $300.00) reads back correctly and the `orig` baseline is preserved. **The scrape mechanism works end-to-end through canvas-word.** ✅
2. **Nesting is NOT preserved — canvas-word's model is flat.** A run carries exactly one `CharStyle.sdtId`; there is no SDT-inside-SDT concept. On import, only the "gap" runs (labels) keep the outer section's id; inner-control runs keep their own. On export, `runsXml` groups *consecutive* same-id runs, so the inner controls interrupt the section and **the single outer section SDT fragments into 4 separate controls** (same tag/alias) on re-import. ❌

**Conclusion / design impact:**
- For **scrape**, nesting is irrelevant — the engine only reads the inner `k:"fb"` controls, which are perfect. The edited docx is consumed only for scraping (values pushed to DB; re-render regenerates from template + DB), so section-SDT fragmentation is harmless to the real flow.
- **Recommendation: use FLAT per-field binding controls.** Keep the section SDT in the render *template* for substitutor dispatch (canvas-word never sees the template before render), but do **not** depend on the section wrapper surviving a canvas-word edit. The `FieldBinding` JSON already carries full identity per field, so flat controls lose nothing for write-back.
- If hierarchy must survive (e.g. in-place section re-render from an edited doc), that requires adding nested-SDT support to canvas-word: an sdtId **stack/path** on runs + push/pop in `documentParser.ts` + nested `w:sdt` emit in `documentXml.ts`. Bounded but real; **not needed for V1 scrape.**
- Normalization confirmed necessary: the edit round-trip reported `200 -> $300.00` — the scrape must normalize the formatted display ("$300.00") back to raw ("300") per `type` before diff/upsert, exactly as specified.
