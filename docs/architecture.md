# Architecture

## Purpose

Pulse separates raw/source data, standardized facts, business classification, deterministic metrics, presentation, and human context.

## Conceptual flow

1. **Source data** — uploaded or API-provided operational data.
2. **Adapter / staging** — source-specific translation into Pulse's neutral structure.
3. **Validation** — period, coverage, duplicates, structure, and import safeguards.
4. **Published facts** — trusted standardized observations.
5. **Mapping** — explicit human-authored rules over source hierarchy nodes.
6. **Reporting Groups** — stable business-owned semantic classifications resolved by inheritance.
7. **Metric engine** — deterministic calculations from published facts after mapping resolution.
8. **KPI Registry** — defines which metrics are active and how they may be presented.
9. **Performance** — interactive exploration of trusted KPI results.
10. **Reports** — meeting-ready consumption of the same metric results.
11. **Human context** — acknowledgement/explanation without rewriting the underlying facts.

## Phase 1 mapping implementation

Build 0.3.0 Phase 1 stores mappings as generic scope/node rules targeting stable
Reporting Group IDs. The current source hierarchy resolves Product → Source
Subcategory → Source Main Category, but those are adapter-exposed levels rather
than a permanent three-level platform limit. Inheritance is computed and does
not copy parent rules into descendant rows.

Product-scoped explicit exclusions use the same Mapping Rule lineage and
effective-date contract. They override inherited membership by resolving the
stable ProductID to the existing Unmapped state; there is no synthetic
“Unmapped” Reporting Group and excluded facts remain in all source totals and
metric denominators.

## Phase 2A metric bridge

Phase 2A joins immutable sales facts to the validated current Effective Mapping
state by stable ProductID. It materializes one derived analysis row per fact,
including effective Reporting Group, explicit resolver status, MappingAsOfDate,
and a deterministic semantic mapping fingerprint.

The materialized bridge is an engine-layer classification projection, not a
second source of business facts. SalesAmount, Quantity, fact identity, import,
restaurant, channel, publication state, and legacy ReportingCategoryID remain
traceable to `tblSalesFacts` and reconcile exactly.

Build 0.3.0 applies the current mapping state to historical facts for analysis.
Fact-date mapping/versioning requires a future explicit architecture decision.

## Phase 2B centralized metric result

Phase 2B materializes KPI-0001 Reporting Group Sales Share in
`tblMetricRPGResults` on `_Metric_Calc`. The grain separates MetricID, ImportID,
ReportingGroupID, and company/restaurant scope. Performance reads the
centralized result rather than calculating directly from legacy CAT membership,
and Reports retains its linkage to the same Performance result.

The active calculation path no longer depends on ReportingCategoryID or
CAT/RPG equivalence. Legacy CAT structures and human-authored equivalence remain
available only for compatibility and migration QA. Unmapped, Conflict, Inactive
Target, and explicit-exclusion facts remain in every applicable denominator.

### Shared Performance restaurant scope

The authoritative default organizational scope for Performance is the current
set of stable RestaurantIDs with `Status=Active` and `ReportingEnabled=Yes`.
This is a shared Performance scope contract, not KPI-0001 business logic.
KPI-0001 consumes it in Phase 2B, and future Performance KPIs must consume the
same scope unless their approved definition explicitly documents a different
organizational-scope requirement.

Facts outside the enabled set remain fully reconciled in Phase 2A but enter
neither the applicable Company numerator nor denominator. Company metric rows
store a fingerprint of the sorted enabled RestaurantIDs so the materialized
scope is auditable. This contract does not add or implement future KPIs.

## Layers

### User layer
- Overview
- Performance
- Reports
- Imports
- Settings

### Administration layer
Includes KPI configuration, organization, classification/mapping, exclusions, coverage expectations, and publication controls.

### Engine layer
Includes raw data, staging, facts, audit records, environment/build metadata, and future metric result stores.

## Domain isolation

Pulse may eventually contain Commercial Performance, Labour, Guest Experience, and other domains. A domain is only included in a final product/view when enabled for that context. Missing or parked domains must not implicitly penalize a restaurant.
