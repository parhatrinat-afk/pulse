# Build 0.3.0 — Weekly Performance Cutover

## Scope

This slice moves the accepted Phase 2C Performance experience from two fixed
legacy datasets to independent ISO-week ranges over the single active compact
weekly cache. It does not add a second metric engine or remove the legacy
`tblMetricRPGResults` rollback surface.

The authoritative flow is:

`Active weekly cache → selected Current/Compare weeks → additive RPG/scope components → accepted Phase 2C numeric helpers → Performance → Reports`

## Controls

Performance exposes Year, From week and To week independently for Current and
Compare. The initial defaults are `2026 W01–W32` and `2025 W01–W32`. Visible
summaries contain business labels only; SourcePeriodKey and CacheVersion remain
engineering helpers.

The selectors are range-backed and formula-driven. Restaurant and Reporting
Group selection, detail Reporting Group, five Display modes, Total, Grand
Total, NOK Impact and sorting retain their accepted Phase 2C behavior.

## Calculation contract

For each selected RestaurantID and ReportingGroupID:

- numerator: sum `tblWeeklyRPGCache[MappedSalesNOK]` for selected weeks;
- denominator: sum `tblWeeklyScopeCache[SourceSalesNOK]` for selected weeks;
- share: numerator divided by denominator after aggregation.

The implementation preserves the accepted share, selected-display, Total,
Grand Total, sorting, detail and text-facade formulas. The current installer can
regenerate those same helpers at the active Reporting Group count so capacity is
not fixed at nine. No weekly percentage is averaged.

Existing Reporting Group selections are preserved by stable ID. A newly active
group defaults to `No` until the user includes it. Detail/sort validation lists,
component ranges and visible matrix columns use deterministic active `SortOrder`.

## Validation and availability

Each range must contain every requested ISO week in `tblWeeklyPeriodManifest`.
Incomplete, unavailable or reversed ranges are visibly marked and calculation
for that side is blocked. Same-period comparisons are allowed and return zero
PP Change/NOK Impact. Different complete lengths remain allowed with a warning.

Weekly Performance is installed only after recomputing the live date-neutral
MappingContentFingerprint, CatalogContentFingerprint and ReportingEnabled
restaurant-scope fingerprint. Formula-time availability also requires the
single accepted Active/Active cache manifest and the current enabled restaurant
set. MappingAsOfDate is audit metadata and is not a weekly freshness input.

## Rollback and non-regression

`tblMetricRPGResults`, the Phase 2B/2C helper structures, legacy imports, facts,
mapping tables and selection catalogs remain intact. Reports continues to read
the Performance detail result and now displays the generated weekly Current and
Compare summaries. No legacy import is superseded and no period-specific KPI
engine is introduced.

## Live acceptance fixtures

Add-ons must validate from the weekly cache at the selected all-restaurant
scope:

| Current / Compare | Current share | Compare share | PP Change | Current Sales NOK | NOK Impact |
|---|---:|---:|---:|---:|---:|
| 2026 W31 / 2025 W31 | 1.1883% | 0.6701% | +0.5182 pp | 80,263 | 35,001.58 |
| 2026 W01–W32 / 2025 W01–W32 | 1.3074% | 0.7336% | +0.5737 pp | 2,383,679 | 1,046,091.46 |
| 2026 W20–W30 / 2025 W20–W30 | 1.2494% | 0.6825% | +0.5669 pp | 835,122 | 378,929.36 |

`tblWeeklyPerformanceQA` records the bounded cutover checks on the existing
hidden Metric Results QA surface. Normal selector changes require Excel
recalculation only and never an Office Script rerun.

## Live Excel acceptance evidence

Installed and validated in canonical `Pulse_Current.xlsx` on 2026-08-14, with
postconditions reconfirmed on 2026-08-15:

- the accepted active cache is `WCV-1a34ad1f46763d9b`, with cache freshness
  `Available` and the accepted WCC/MCF/ICC/IDP/RSC fingerprints;
- Phase 2C Interaction QA and Weekly Performance QA are both 16/16 PASS;
- all six Year/From/To controls use List validation;
- the four additive component matrices reference only `tblWeeklyRPGCache` and
  `tblWeeklyScopeCache`, contain no `tblMetricRPGResults` or percentage-average
  path, and have no formula errors;
- the three Add-ons fixtures above match exactly;
- same-period, different-length complete, unavailable, incomplete/reversed,
  restaurant/RPG selection, all five displays, Total/Grand Total, sorting,
  Detail Reporting Group, and Reports linkage behave as specified without an
  Office Script rerun; and
- protected rollback surfaces remain unchanged: `tblMetricRPGResults`
  fingerprint `3b467e3929d8f06a` with 306 rows and `tblImports` fingerprint
  `201d0869568c427b` with two rows.

The accepted restored selector state is Current `2026 W01–W32`, Compare
`2025 W01–W32`, Detail Reporting Group `Wine & Sake`, Display `NOK Impact`,
Sort by `Total`, Order `Highest first`, with all 16 restaurants and all nine
Reporting Groups selected.
