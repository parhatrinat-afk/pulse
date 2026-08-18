# Build 0.3.0 — Imports Presentation Cleanup

This bounded slice turns the visible `Imports` sheet into an operational weekly
intake view without changing ingestion, cache authority, legacy imports, facts,
Performance, Reports, Mapping, or schemas.

## Visible experience

The normal-user view contains:

- title `Imports` and subtitle `Weekly sales reports processed by Pulse.`;
- summary for latest published week, active-cache coverage and readiness;
- `Weekly Import Activity` with Period, friendly Status, Processed timestamp,
  safe Source report label, Rows, Sales NOK and a concise Message.

The activity view is a dynamic formula facade over genuine rows in
`tblWeeklyIntakeLog`, newest first. It does not synthesize W01–W32 entries from
the cache. Source locators are shown as filenames only when they are already a
safe `.xlsx` filename/path; opaque OneDrive identifiers and manual QA tokens are
displayed as `Source report`.

## Preserved engineering evidence

`tblImports` remains unchanged at `Imports!A4:S6` with its stable IDs,
certificates, lineage and fingerprints. Its rows and the unused technical
columns are hidden from the normal view, not deleted or moved. A hidden blank
buffer row prevents Excel from auto-expanding the table into the operational
summary on reruns. Weekly cache
version/period/intake tables remain the authoritative sources on the hidden
engineering surface.

The script requires one `Active / Active / PASS` cache with 85 periods, one
rollback version, coverage through 2026 W33, four genuine intake events, Weekly
Performance QA 16/16 PASS and Phase 2C Interaction QA 16/16 PASS. It fingerprints
all protected tables plus Performance, Reports and Mapping before and after the
presentation change.

## Live checkpoint

At the accepted W33 checkpoint the visible activity rows are, newest first:

1. 2026 W33 — Duplicate — no data change
2. 2026 W33 — Published
3. 2026 W32 — Duplicate — no data change
4. 2026 W31 — Duplicate — no data change

The summary is `2026 W33`, coverage is `2025 W01–W52` and `2026 W01–W33`, and
status is `Up to date`.
