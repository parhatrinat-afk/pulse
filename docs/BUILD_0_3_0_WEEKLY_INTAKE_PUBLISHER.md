# Build 0.3.0 — Weekly Dynamic Authority and Intake Publisher Foundation

## Scope

This slice removes the weekly Performance dependency on one frozen CacheVersion
and defines the deterministic repository-side contract for adding one parsed
weekly report to the accepted compact cache. It does not create a Power
Automate flow, move OneDrive files, activate a new cache, change Imports, or
supersede the retained legacy imports.

## Dynamic cache authority

A weekly consumer resolves authority only when exactly one row in
`tblWeeklyCacheVersions` has both:

- `CacheStatus = Active`; and
- `ActivationState = Active`.

The resolved row must have `ValidationStatus = PASS`, nonblank cache and
identity fingerprints, materialized row counts matching its manifest, current
MappingContent and catalog-content fingerprints, and the current
ReportingEnabled restaurant-scope fingerprint. The active cache's period
manifest must not repeat an ISO year/week.

Zero exact active rows returns `Unavailable — no active cache`. More than one
returns `Unavailable — multiple active caches`. A single authority with stale
content returns `Stale / unavailable`. Consumers never select the first row or
fall back to a frozen WCV/WCC constant.

The accepted component, period, selection, Total, Grand Total, sorting, NOK
Impact, text-facade and Reports formulas are otherwise unchanged.

## Publisher contract

`planWeeklyIntakePublication` accepts one result from the approved weekly
parser, one fully validated active cache snapshot, current Pulse catalogs and
current freshness fingerprints. For a new period it:

1. validates parser/schema/lineage/source reconciliation;
2. validates the single active cache and its full materialized fingerprint;
3. applies the accepted identity preflight and hierarchy-based mapping resolver
   to the incoming week;
4. rekeys the retained historical cache rows into a new version without
   mutating the active rows;
5. adds the new week scope and dense nine-RPG rows;
6. reconciles every period, mapping-state coverage and mapped RPG totals;
7. derives deterministic WCV/WCC fingerprints; and
8. returns `Candidate / Not Active` state ready for later materialization and
   activation.

The incremental Candidate's additive business rows must equal an independent
full rebuild over the same source reports. Identity Pending remains in scope
denominators and outside RPG numerators.

## Classification outcomes

| Outcome | Condition | Cache effect |
|---|---|---|
| New | Fresh valid period not in the active manifest | Complete inactive Candidate returned |
| Duplicate | Same SourcePeriodKey and semantic fingerprint | No-op |
| Correction Review | Same SourcePeriodKey, different semantic fingerprint | No-op; explicit supersession required |
| Rejected | Parser, schema, lineage or reconciliation invalid | No-op |
| Cache Stale | Authority, mapping, catalog/identity or restaurant scope is unavailable/stale | No-op |

Filenames and source locators are audit fields only. Identity Pending does not
cause rejection.

## Intake ledger contract

The hidden `tblWeeklyIntakeLog` on `_Weekly_Cache` uses this minimal operational
schema:

`IntakeEventID, SourceLocator, SourceFileID, SourcePeriodKey,
SourceSemanticFingerprint, IdentityPreflightFingerprint, IntakeStatus,
StatusMessage, SourceRowCount, SourceSalesNOK, ProcessedAt, PriorCacheVersion,
ResultingCacheVersion, SupersededCacheVersion`

The deterministic IntakeEventID supports idempotency. The table is not a fact
store, source archive or visible Imports redesign.

## Full-version retention

The safe target is at most two materialized full analytical versions: the
single Active version and one previous rollback version. All source, period,
version and intake manifests remain historical evidence. This slice only
returns which older inactive analytical versions would later be eligible for
cleanup; it deletes nothing.

## Runtime adapter boundary

`Process_Weekly_Intake.ts` is the thin workbook runtime entry point. `Process`
accepts the parser version/schema plus the parsed manifest, resolves the live
Active authority, classifies the event and upserts the ledger. It returns a
small typed status used by Power Automate for routing.

The script deliberately does not port the parser, identity preflight, mapping
resolver or cache publisher. An accepted repository publisher result for a New
period is converted into deterministic `Prepare New`, bounded `Write New`, and
`Finalize New` payloads. Finalization validates row IDs, grains and WCC before
writing the Candidate version manifest. The new surface remains hidden and
`Candidate / Not Active`; this adapter cannot switch authority.

The first Power Automate pilot is Duplicate/no-op only. A genuine New
publication requires an approved way to execute the repository publisher and
deliver its contract-validated Candidate payloads; that orchestration is not
implemented by this slice.
