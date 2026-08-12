# Changelog

All meaningful Pulse changes are recorded here.

## Unreleased — 0.3.0 — Mapping + Reporting Groups

### Implemented in Phase 1 foundation

- Added a rerunnable Phase 1 Office Script for the validated 0.2.0-QA checkpoint.
- Added an authoritative Reporting Group registry with nine stable seed IDs.
- Added generic scope/node Mapping Rules targeting Reporting Group IDs.
- Added hierarchy browse tables with descendant counts and sales/quantity context.
- Added deterministic Effective Mapping output with explicit/inherited state.
- Added overlap-conflict, inactive-target, unmapped, and reconciliation QA.
- Preserved legacy Reporting Categories, Remap Rules, facts, and the 0.2.0 metric path.

### Implemented in Phase 2A source

- Added a deterministic Reporting Group metric contract without activating new
  Performance behavior.
- Added a materialized one-row-per-fact bridge from immutable sales facts to
  current Effective Mapping.
- Added current-date stale mapping rejection and a reproducible semantic mapping
  fingerprint stored on every bridge row.
- Added separate Mapped, Unmapped, Conflict, and Inactive Target reconciliation
  for Sales NOK, Quantity, and fact count.
- Added mapped Reporting Group totals, an explanatory legacy CAT/RPG crosswalk,
  and human-authored ID-based equivalence comparisons.
- Added deterministic fixtures and Office Scripts compatibility/regression
  checks.
- Hardened Excel-for-web output range handling with bounded indexed ranges,
  exact range diagnostics, output row-width checks, and batched protected-sheet
  reads for the 18,086-row checkpoint.
- Completed live Excel-for-web Phase 2A QA on 2026-08-11: all source/bridge and
  mapping-state reconciliations passed, legacy surfaces remained protected, and
  the clean no-equivalence state returned the expected WARN.
- Preserved `_Metric_Calc`, Performance, Reports, KPI Registry, KPI-0001, raw
  source data, and `_Sales_Facts` unchanged.

### Implemented and live validated in Phase 2B

- Added a pre-mutation Phase 2B bridge preflight that verifies today's mapping
  fingerprint, one bridge row per fact, complete lineage, Sales NOK, Quantity,
  fact count, and four-state coverage.
- Added deterministic centralized `tblMetricRPGResults` rows at the grain
  KPI × dataset × Reporting Group × company/restaurant scope.
- Migrated KPI-0001 to Reporting Group Sales Share while retaining its stable
  KPI ID.
- Replaced the active `_Metric_Calc` CAT helper with active Reporting Groups and
  stable RPG IDs.
- Migrated Performance formulas from direct legacy CAT `SUMIFS` to centralized
  RPG result lookups, preserving independent current/comparison datasets,
  all-channel behavior, and the restaurant breakdown.
- Preserved Reports as a consumer of the same Performance result.
- Added visible Metric Results QA, deterministic fixtures for all nine RPGs,
  accepted Add-ons targets, and Office Scripts compatibility guards.
- Preserved `_Sales_Facts`, `_Metric_RPG_Facts`, mapping configuration,
  Effective Mapping, Lovable decisions, legacy CAT structures, and CAT/RPG
  equivalence definitions.
- Defined the Phase 2B Company scope as active, ReportingEnabled restaurants;
  excluded restaurants affect neither numerator nor denominator while remaining
  intact in Phase 2A reconciliation.
- Established that enabled set as the shared default Performance scope contract,
  consumed by KPI-0001 now and by future KPIs unless explicitly overridden by
  an approved KPI definition; no future KPI was implemented.
- Added deterministic Company restaurant-scope fingerprints, dynamic
  306/288-row cardinality coverage, and enabled-plus-excluded scope
  reconciliation for facts, Sales NOK, and Quantity.
- Passed live Excel-for-web QA on 2026-08-12. Sixteen enabled restaurants
  produced 306 rows and scope fingerprint `RSC-08df626f217dd94b`; deselecting
  Swift produced 288 rows, excluded 764 facts / NOK 12,561,017.69 / Quantity
  3,008 from Performance only, and re-enabling Swift restored the exact original
  fingerprint, cardinality, and accepted Add-ons results.
- All sixteen Metric Results QA checks passed throughout the restaurant-scope
  round trip.

### Still planned / later phases

- Complete branch review and merge/release acceptance for Phase 2B.
- Continue to Phase 3 only after Phase 2B acceptance.

This section describes active development only. It does not advance the validated release checkpoint.

## [0.2.0-QA] — Current validated checkpoint

### Validated

- Build 0.2.0 workbook and deterministic Category Sales Share vertical slice were subjected to pre-0.3.0 QA.
- QA fixes and supporting documentation were added without changing the core product principles.
- The resulting `Pulse_Build_0_2_0_QA` workbook is the checkpoint used as the starting point for Build 0.3.0.

## [0.2.0] — Category Sales Share

### Added

- Reproducible Office Script build path.
- First deterministic KPI: Category Sales Share.
- Independent current/comparison dataset selection.
- Independent current/comparison channel selection.
- All-channels scope.
- Non-blocking scope mismatch visibility.
- Performance and Reports vertical slice.

### Corrected

- Removed In-house from the KPI definition.
- Removed the assumption that 2025 is permanently the baseline.
- Comparison/baseline reference is user-selected.

### Preserved

- Immutable sales facts.
- Human decision ownership.
- Source neutrality.
- Remap Assistant and Remap Rules naming during the 0.2.0 implementation.

## [0.1.1]

### Changed

- Reorganized the workbook around a user-first experience.
- Primary workflow: Overview, Performance, Reports, Imports, Settings.
- Supporting administration and engine sheets retained rather than removed.
- Refined workbook hierarchy and presentation.

### Preserved

- Existing import architecture.
- Fact store and mapping structures.
- KPI Registry.
- Remap Assistant and Remap Rules names.
- 2025 baseline and Week 31 test data contained in the workbook.

## [0.1.0]

Initial structured Pulse workbook checkpoint containing the core data model, imports, mapping, facts, configuration, and product shell.
