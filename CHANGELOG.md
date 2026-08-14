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

### Implemented and live Excel accepted in Phase 2C

- Added stable-ID Yes/No restaurant and Reporting Group selection over the
  refresh-time eligible Performance scope. Selection state is authoritative;
  All versus Custom is derived, not user-selected.
- Preserved the separate single-RPG detail selector while adding a multi-RPG
  matrix selection table.
- Reused Phase 2B Restaurant result components for arbitrary selected scopes;
  no fact reads, new metric engine, or restaurant-combination results were added.
- Added PP Change, Current Share, Comparison Share, Current Sales NOK, and NOK
  Impact matrix displays with recalculation-only interaction. NOK Impact applies
  the comparison Reporting Group share to current selected-scope total sales,
  then compares that baseline amount with current Reporting Group sales.
- Kept the six canonical component matrices and a bounded selected-display
  helper numeric while rendering the visible matrix through an isolated
  FIXED-based text facade. This avoids Excel-for-web conditional number-format
  parsing without changing calculations or requiring a script for Display
  changes.
- Defined share Grand Total as summed selected numerators divided by summed
  denominators. NOK Impact Grand Total uses the aggregated current numerator
  minus aggregated comparison share × aggregated current denominator; it never
  sums restaurant-level impacts.
- Added a Total column immediately after Restaurant. It aggregates only the
  selected RPG numerators, uses each selected scope denominator once for
  share/PP/NOK Impact, and keeps Current Sales NOK equal to the selected current
  numerator without adding a denominator dependency.
- Added formula-driven Sort by and Order controls. Full-precision numeric
  helpers sort restaurants by Total or a displayed RPG; unavailable keys remain
  last, ties use deterministic RestaurantID ordering, and Grand Total remains
  fixed outside the ranking.
- A retained-but-hidden RPG sort target is visibly reported and falls back to
  Total until the RPG is displayed again. Sorting changes presentation order
  only and leaves canonical component/helper identity untouched.
- Removed the redundant visible Restaurant mode and Reporting Group mode
  controls. Selection summaries now report all-selected or the selected count,
  and concise wrapped scope/sort status text avoids the live-QA clipping found
  in the earlier control layout.
- Kept canonical zero-denominator values at zero while displaying an em dash in
  Performance.
- Preserved Phase 2B Company rows as QA/control totals and linked Reports to the
  same interactive detail scope.
- Added pre-mutation centralized-result validation, protected-surface
  fingerprinting, Phase 2C Interaction QA, deterministic tests, and a live Excel
  checklist.
- Corrected QA-0302C-09 so the legitimate `Current Sales NOK` Display label is
  not mistaken for an ` NOK` text-presentation suffix in the authoritative
  numeric helper. The five Display modes and their calculations are unchanged.
- Completed live Excel validation with all 16 Interaction QA checks PASS,
  accepted Total/Grand Total and full-precision sorting behavior, all five
  display modes, recalculation-only selection, Reports linkage, and the exact
  Add-ons result round trip.
- Added the six-sheet information-architecture slice and restrained Visual
  Slices 2A/2B. The resulting workbook is the accepted clean functional 0.3.0
  foundation, not the final Pulse UI.

### Implemented in weekly source/identity foundations

- Added a deterministic read-only parser for the exact seven-column weekly POS
  `Sales per Item` export, using the internal Period field rather than filenames
  or folders for period identity.
- Added runtime-independent semantic fingerprints, exact source-string
  preservation, source-row lineage, scope/schema validation, and an 84-week
  self-reconciling corpus manifest.
- Added exact-key reuse of current Restaurant, Product, and Source
  Classification catalogs plus deterministic stable-ID candidates for new
  unambiguous source identities.
- Kept identity recognition separate from Mapping. New ProductIDs inherit only
  existing approved hierarchy rules and otherwise remain Unmapped.
- Added explicit Identity Pending handling with reconciled impact for current
  catalog collisions and new multi-hierarchy ProductKeys.
- Preserved current Product hierarchy as mapping authority and surfaced the
  nine accepted alternate-path mapping outcomes as review evidence, including
  PRD-000689 Red Curry (RPG-0001 versus RPG-0009).
- Kept the two exact Test Department spellings separate as RST-0017/RST-0018
  candidates with ReportingEnabled=No.
- Reconciled all 245,632 weekly rows, NOK 484,728,367.25, and Quantity
  2,469,988.09 without workbook publication, cache activation, legacy cutover,
  Power Automate, Performance, or Phase 3 changes.
- Added a deterministic candidate compact weekly cache with one source/state
  denominator row per Restaurant/week and nine dense mapped-RPG numerator rows
  per Restaurant/week.
- Added version/period manifests, current mapping/preflight freshness rejection,
  stable row/version/cache fingerprints, candidate-only activation safety, and
  the shared Performance restaurant-scope fingerprint.
- Reconciled all 84 weeks, each source year, and the complete corpus across
  Mapped, Unmapped, Identity Pending, Conflict, and Inactive Target; kept the
  two Test Department identities outside normal Performance scope without
  removing their facts.
- Added independently aggregated W31, W01-W32, and W20-W30 range fixtures for
  share, PP Change, Current Sales NOK, and NOK Impact. The 14,210-row analytical
  candidate remains repository-only and does not cut over Performance.

### Still planned / later phases

- Complete controlled branch merge and Build 0.3.0 release acceptance.
- Phase 3 remains unstarted and requires a separate approved scope.

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
