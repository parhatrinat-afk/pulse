# Build 0.3.0 — Weekly Intake Office Script Adapter and Ledger

## Scope

This slice supplies the thin Excel runtime boundary needed by a later Power
Automate flow. It reuses the accepted weekly parser, dynamic Active-cache
authority and repository intake publisher. It does not implement a flow,
recalculate identity/mapping/cache business logic in Excel, activate a new
cache, redesign Imports, or change Performance.

## Process contract

Power Automate calls `Process_Weekly_Intake.ts` with:

- `operation = Process`; and
- `payloadJson` containing adapter/publisher contract versions, parser/schema
  versions, the parsed source manifest, optional parser error, and processed
  timestamp.

The source rows are deliberately excluded from this call. The script validates
the parser metadata, resolves exactly one fresh `Active / Active` cache, and
looks up `SourcePeriodKey` plus `SourceSemanticFingerprint` in the active period
manifest. It supports exactly:

| Status | Runtime effect | Archive-ready |
|---|---|---|
| New | No ledger row until an accepted complete Candidate is finalized | No |
| Duplicate | Idempotent ledger upsert; no cache change | Yes |
| Correction Review | Idempotent ledger upsert; no cache change | No |
| Rejected | Idempotent ledger upsert; no cache change | No |
| Cache Stale | Idempotent ledger upsert; no cache change | No |

The return object contains `status`, `period`, `message`, `cacheChanged`,
`activeCacheVersion`, `resultingCacheVersion`, `ledgerEventId`, `ledgerAction`,
`archiveReady`, and `operation`. Power Automate branches on `status` and
`archiveReady`; it does not reproduce Pulse calculations.

## Ledger

`tblWeeklyIntakeLog` is stored at `_Weekly_Cache!Y1:AL...`. `_Weekly_Cache`
remains hidden. The columns are:

`IntakeEventID, SourceLocator, SourceFileID, SourcePeriodKey,
SourceSemanticFingerprint, IdentityPreflightFingerprint, IntakeStatus,
StatusMessage, SourceRowCount, SourceSalesNOK, ProcessedAt, PriorCacheVersion,
ResultingCacheVersion, SupersededCacheVersion`.

`IntakeEventID` is deterministic from the source period, semantic fingerprint,
outcome and resulting version. Reprocessing the same semantic event returns
`ledgerAction = Existing`; it cannot append duplicate analytical data or
duplicate the ledger outcome.

## New Candidate materialization

The repository publisher remains the only business-logic implementation. For
an accepted `New` result, `weekly-intake-office-adapter.mjs` produces:

1. `Prepare New` — creates hidden `_Weekly_Cache_Candidate` and fixed headers;
2. `Write New` — writes deterministic bounded chunks with overlap protection;
3. `Finalize New` — validates complete rows, grains and cache fingerprint,
   then writes the Candidate version manifest and New ledger outcome.

The active cache is checked before every operation and again after finalizing.
The Candidate stays `Candidate / Not Active`. No authority transition exists in
the adapter, so a failed or partial write cannot become analytical authority.
The current `_Weekly_Cache` table layout cannot safely expand a second full
version in place; the separate hidden Candidate surface avoids table overlap
while keeping the active cache untouched.

## Duplicate/no-op pilot

The first live Power Automate pilot may use a copied, already-published weekly
report:

1. obtain the OneDrive file content and stable source locator;
2. run `Parse_Weekly_Sales_Report.ts` on that untouched source workbook;
3. construct the compact Process payload from the parser manifest (do not pass
   normalized source rows);
4. run `Process_Weekly_Intake.ts` against canonical `Pulse_Current.xlsx`;
5. require `status = Duplicate`, `cacheChanged = false`, an unchanged active
   WCV/WCC, and `archiveReady = true` before archiving the source copy; and
6. confirm the ledger row exists and Weekly Performance/Phase 2C QA remain
   16/16 PASS.

Any other status must not archive as a successful duplicate. No genuine future
week is approved for this pilot.

## Live workbook evidence — 2026-08-17

Canonical `Pulse_Current.xlsx` received hidden `tblWeeklyIntakeLog` on the
existing `_Weekly_Cache` surface. A safe 2026 W31 duplicate fixture returned:

- `status = Duplicate`;
- `cacheChanged = false`;
- `archiveReady = true`;
- `ledgerAction = Added`, followed by `Existing` on the idempotency check; and
- `IntakeEventID = WINT-49bb211914c81b65`.

The Active authority remained `WCV-1a34ad1f46763d9b` /
`WCC-508dd608166cdb6e`. Performance, Reports, Imports and the retained 306-row
`tblMetricRPGResults` surface were unchanged; Weekly Performance QA and Phase
2C Interaction QA remained 16/16 PASS. No Candidate was created or activated.
