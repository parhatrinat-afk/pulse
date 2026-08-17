# Build 0.3.0 Weekly Intake Office Script Adapter checklist

- [x] Runtime supports exactly New, Duplicate, Correction Review, Rejected and Cache Stale.
- [x] `tblWeeklyIntakeLog` uses the approved 14-column schema on hidden `_Weekly_Cache`.
- [x] Duplicate processing is an idempotent no-op and the only archive-ready Process result.
- [x] Parser source rows do not cross the thin Process boundary.
- [x] Dynamic Active authority and current freshness are revalidated at runtime.
- [x] Accepted New Candidates use deterministic prepare/chunk/finalize payloads.
- [x] Candidate rows are written on a separate hidden engineering surface.
- [x] Candidate finalization validates IDs, grains and WCC before the version manifest.
- [x] The adapter cannot activate a Candidate or mutate the Active analytical rows.
- [x] Repository fixtures cover all five outcomes, idempotency and static compatibility.
- [x] Live canonical workbook contains hidden `tblWeeklyIntakeLog`.
- [x] Live Duplicate/no-op adapter run returns `archiveReady = true`.
- [x] Live Active WCV/WCC, Performance, Reports and QA remain unchanged.
