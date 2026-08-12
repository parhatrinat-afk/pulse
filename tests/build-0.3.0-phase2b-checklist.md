# Build 0.3.0 Phase 2B validation checklist

## Accepted live result

Live Excel-for-web QA passed on 2026-08-12. The accepted evidence includes:

- sixteen enabled restaurants: 306 rows, zero excluded facts, and Company scope
  fingerprint `RSC-08df626f217dd94b`;
- accepted Add-ons values of NOK 131,487 / 1.47% for W31 2026, 1.09% for 2025,
  and approximately +0.38 percentage points change;
- Swift deselected: 288 rows, fifteen restaurants, and 764 facts / NOK
  12,561,017.69 / Quantity 3,008 excluded from Performance only;
- deselected Company Add-ons of NOK 115,607 / 1.33% and 1.12% comparison;
- Swift absent from the restaurant breakdown;
- Swift re-enabled: exact restoration of 306 rows, zero excluded facts,
  fingerprint `RSC-08df626f217dd94b`, and accepted Add-ons results; and
- all sixteen Metric Results QA checks PASS throughout.

Run against a QA copy after the accepted Lovable migration, Phase 1 refresh,
and Phase 2A refresh have completed on the same day.

## Prerequisite and failure safety

- [ ] Phase 2A has exactly one `_Metric_RPG_Facts` row per source fact.
- [ ] Effective Mapping and bridge MappingAsOfDate are today.
- [ ] Bridge mapping fingerprint matches current Mapping Rules, Reporting
  Groups, Products/Source Classifications, and Effective Mapping.
- [ ] A prior-day bridge fails with `PUL-0302B-001` before `_Metric_Calc` or
  Performance changes.
- [ ] Editing a Mapping Rule without rerunning Phase 1/2A produces the same
  pre-mutation failure.
- [ ] A duplicate/missing SalesFactID or source/bridge measure variance fails
  before presentation mutation.

## Central result layer

- [ ] `_Metric_Calc` contains `tblMetricRPGResults` at O:AC.
- [ ] It contains 306 data rows for the accepted checkpoint.
- [ ] Every MetricResultID is unique and deterministic.
- [ ] Every row uses `MetricID=KPI-0001`.
- [ ] The two active ImportIDs are represented independently.
- [ ] The nine active ReportingGroupIDs are represented.
- [ ] ScopeType has one Company row and sixteen Restaurant rows per
  dataset/RPG.
- [ ] Company rows contain the deterministic fingerprint of the sorted active,
  ReportingEnabled RestaurantIDs; Restaurant rows leave it blank.
- [ ] `ChannelScope=All channels` and no mandatory channel selector exists.
- [ ] MappingAsOfDate and MappingFingerprint are consistent across all rows.

## Metric semantics

- [ ] Numerator contains only `ResolutionStatus=Mapped` facts for the selected
  ReportingGroupID.
- [ ] Denominator includes Mapped, Unmapped, Conflict, Inactive Target, and
  explicit exclusion facts in the identical enabled-restaurant scope.
- [ ] A restaurant with `Status` other than Active or `ReportingEnabled` other
  than Yes contributes to neither Company numerator nor denominator and has no
  Restaurant-scope result.
- [ ] The enabled-restaurant definition is documented as shared Performance
  scope configuration, not KPI-0001-specific business logic.
- [ ] Zero-denominator scopes return zero rather than an Excel error.
- [ ] Company and restaurant denominators reconcile independently.
- [ ] Current and comparison datasets remain independently selectable.
- [ ] No active formula filters `_Sales_Facts[ReportingCategoryID]`.
- [ ] No CAT/RPG equivalence is inferred or consumed by KPI-0001.

## Accepted targets

- [ ] W31 2026 / RPG-0001 Add-ons sales is NOK 131,487.00.
- [ ] W31 2026 / RPG-0001 Add-ons share displays 1.47%.
- [ ] 2025 / RPG-0001 Add-ons sales is NOK 4,543,795.00.
- [ ] 2025 / RPG-0001 Add-ons share displays 1.09%.
- [ ] Change displays approximately +0.38 percentage points.
- [ ] All nine RPG numerators match the approved Phase 2B fixture by dataset.
- [ ] Result-row cardinality is 306.
- [ ] With one of sixteen restaurants deselected and Phase 2B rerun, result-row
  cardinality is 288 and the Company result excludes that restaurant.

## Performance and Reports

- [ ] Performance shows Reporting Group Sales Share and retains KPI-0001.
- [ ] Reporting Group selector contains only the nine active RPG names.
- [ ] `_Metric_Calc!I:J` contains Reporting Group name and stable RPG ID.
- [ ] Company overview shows nine RPG rows, not legacy CAT rows.
- [ ] Restaurant breakdown remains present and uses the selected RPG.
- [ ] Existing column widths, styling, editable-cell cues, and Explain section
  remain readable.
- [ ] Performance has no B11/G11 mandatory channel controls.
- [ ] Reports links to Performance and displays identical current, comparison,
  change, and current-sales results.

## Reconciliation and protection

- [ ] Fact count remains 18,086.
- [ ] Source Sales NOK remains 426,611,113.82.
- [ ] Source Quantity remains 2,069,940.12.
- [ ] Mapped facts/sales/quantity remain
  15,935 / 405,493,085.13 / 1,925,100.06.
- [ ] Unmapped facts/sales/quantity remain
  2,151 / 21,118,028.69 / 144,840.06.
- [ ] Conflict and Inactive Target remain zero in the accepted checkpoint.
- [ ] Performance-enabled plus excluded Performance scope equals the complete
  Phase 2A bridge for fact count, Sales NOK, and Quantity.
- [ ] All sixteen `Metric Results QA` checks are PASS, except a surfaced selector
  validation WARN is acceptable only if Excel rejects nonessential dropdown UI.
- [ ] `_Sales_Facts`, `_Metric_RPG_Facts`, Mapping Rules, Reporting Groups,
  Effective Mapping, Restaurants, Products, Source Classifications, legacy CAT
  tables, and Metric Equivalence remain unchanged during script execution.
- [ ] Phase 2A QA-0302A-08 remains the expected WARN because no human-authored
  CAT/RPG equivalence definitions exist.

## Automated/runtime

- [ ] `node --test tests/*.test.mjs` passes.
- [ ] `git diff --check` passes.
- [ ] Phase 2B compiles and completes in Excel for the web.
- [ ] Excel reports no Map/Set iterator compatibility error.
- [ ] Excel reports no worksheet read-inside-loop performance warning for
  source/table reads.
- [ ] Rerunning Phase 2B replaces the result table without duplicating rows.
- [ ] Changing `tblRestaurants[ReportingEnabled]` affects Performance only after
  Phase 2B is rerun; no interactive Performance multi-select is introduced.
- [ ] A new live QA checkpoint is reviewed before commit/merge.
