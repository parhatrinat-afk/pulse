# Reporting Group Sales Share implementation

Build 0.3.0 Phase 2B activates KPI-0001 on the authoritative Reporting Group
semantic layer.

## Definition

`Reporting Group Sales Share = selected RPG mapped Sales NOK / all Sales NOK in the identical scope`

The numerator includes only `_Metric_RPG_Facts` rows where:

- `PublicationState = Active Finalized`;
- `ResolutionStatus = Mapped`; and
- `EffectiveReportingGroupID` equals the selected stable RPG ID.

The shared Performance restaurant scope is the stable RestaurantID set where
`tblRestaurants[Status] = Active` and `ReportingEnabled = Yes`. A bridge fact
outside that set contributes to neither a Company numerator nor a Company
denominator and receives no Restaurant result. The fact remains unchanged in
the bridge and complete Phase 2A reconciliation.

This restaurant definition is independent of KPI-0001. KPI-0001 consumes the
shared scope in Phase 2B; future Performance metrics must consume it unless an
approved metric definition explicitly requires another organizational scope.

Inside the enabled restaurant set, the denominator includes every
`Active Finalized` bridge row in the same ImportID and company/restaurant
scope. Unmapped, Conflict, Inactive Target, and explicit Product exclusions
therefore remain in the denominator.

The validated presentation is all-channel. Channel remains bridge/fact
metadata and a future optional scope dimension; it is not part of KPI-0001's
identity.

## Central result layer

`tblMetricRPGResults` is materialized on `_Metric_Calc` with one row per:

`MetricID × ImportID × ReportingGroupID × Scope`

The current checkpoint has one KPI, two active datasets, nine active Reporting
Groups, one company scope, and sixteen restaurant scopes: 306 deterministic
rows. With fifteen enabled restaurants the same grain produces 288 rows.
`MetricResultID` is derived only from stable IDs and scope identity.

Company rows store `RestaurantScopeFingerprint`, calculated from the sorted
enabled RestaurantIDs. A configuration change therefore changes the auditable
Company scope identity without changing source facts. Restaurant rows retain
their own RestaurantID and leave the combined-scope fingerprint blank.

Performance reads `MetricValue` and `NumeratorSalesNOK` from this table. It no
longer filters `_Sales_Facts[ReportingCategoryID]`. Reports preserves its
existing linkage to Performance and therefore presents the same result.

## Freshness and lineage

Before any workbook write, Phase 2B:

1. fingerprints current Reporting Groups, Mapping Rules, product hierarchy,
   and materialized Effective Mapping using the Phase 2A semantic-v2 contract;
2. requires today's MappingAsOfDate and a single matching bridge fingerprint;
3. compares the bridge one-for-one with `_Sales_Facts` by SalesFactID and
   source lineage/measures;
4. reconciles Sales NOK, Quantity, fact count, and all four mapping states;
5. proves enabled Performance scope plus excluded Performance scope equals the
   complete Phase 2A bridge; and
6. independently recalculates every proposed central result with the identical
   enabled-restaurant filter.

Failure stops before `_Metric_Calc`, Performance, Reports, KPI Registry, or
Overview is changed.

Build 0.3.0 applies current mapping state to historical facts for analysis.
Fact-date mapping/versioning remains out of scope.

## Legacy compatibility

Legacy Reporting Categories, `_Sales_Facts[ReportingCategoryID]`, and the
human-authored CAT/RPG equivalence table remain intact for compatibility and
migration QA. Active KPI-0001 results do not infer or consume CAT/RPG
equivalence.
