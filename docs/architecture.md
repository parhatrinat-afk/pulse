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

## Phase 2C interaction layer

Phase 2C treats the Phase 2B eligible set as a refresh boundary and adds a
separate interactive selection state. Stable-ID Yes/No tables select any
subset of eligible restaurants and active Reporting Groups. All versus Custom
is derived from those rows, not exposed as a separate control. Formula helpers
sum the selected Phase 2B Restaurant-row numerators and denominators; they do
not read facts or rematerialize combinations.

The selected scope applies identically to current and comparison datasets.
Share/PP Grand Total derives from summed selected numerators and denominators.
NOK Impact subtracts aggregated comparison share × aggregated current
denominator from the aggregated current numerator; it is never summed from
restaurant-level impacts.
Phase 2B Company rows remain unchanged as auditable all-eligible controls and
must equal Phase 2C when every eligible restaurant is selected.

The six `_Metric_Calc` component matrices and a bounded selected-display helper
remain numeric. Performance renders the visible matrix through an isolated
text facade so Display changes can recalculate without relying on conditional
number formats. No metric, QA reconciliation, or Reports calculation may parse
or otherwise consume that visible facade.

A separate bounded Total helper aggregates the currently selected RPG
numerators while retaining one current and one comparison denominator per
restaurant scope. Share, PP Change, and NOK Impact are recomputed from those
aggregate components; Current Sales NOK is the aggregate current numerator and
does not depend on a denominator. A numeric sort-key helper orders only the
visible restaurant lookup layer by Total or a displayed RPG. Canonical
RestaurantID/component rows remain unchanged, and Grand Total never enters the
sort range.

This interaction mechanism is KPI-independent as scope state, but the additive
aggregation is valid for KPI-0001 specifically because its approved numerator
and denominator are additive across disjoint restaurants. Future KPIs require
their own approved component contract.

## Future workbook presentation boundary

The development workbook currently exposes many engineering surfaces. They are
valid implementation and troubleshooting assets, but they are not the intended
long-term normal-user navigation model. A future explicitly scoped UX phase
should target approximately these everyday surfaces:

- Overview;
- Performance;
- Reports;
- Imports / Refresh; and
- Mapping.

Restaurants, Reporting Groups, Settings, and KPI Registry are appropriate
administration/configuration surfaces. Raw, staging, adapter, fact, bridge,
helper, metric, audit, environment, build, QA, legacy CAT/remap, and other
engineering sheets should eventually be hidden or protected from normal users
while remaining available for advanced/admin troubleshooting. Build 0.3.0
Phase 2C does not hide or remove any sheet.

The accepted calculation architecture remains underneath that future boundary:

- Performance should receive visual polish, stronger Total/Grand Total
  hierarchy, and more polished selection panels without calculation redesign.
- Overview should become a live, minimal, clickable Pulse home screen and is a
  natural future home for the Pulse ♥ refresh workflow.
- Reports should evolve into a meeting/export surface while consuming the same
  centralized results.
- Mapping should preserve hierarchical inheritance, stable IDs, explicit
  exceptions, and rule lineage while presenting a simpler needs-attention →
  choose Reporting Group → refresh workflow. Legacy Remap Assistant, Remap
  Rules, and Effective Categories belong behind the authoritative RPG mapping
  experience rather than alongside it for normal users.
- Imports should retain the engineering workflow underneath while eventually
  presenting upload/import → Pulse ♥ → validation → refreshed model.
- Detailed QA remains valuable engineering evidence; future user-facing views
  may summarize it as readiness/issues indicators rather than exposing every QA
  table.

This is future-interface guidance only. It does not authorize sheet hiding,
Pulse ♥ implementation, Overview/Reports/Mapping redesign, or Phase 3 work.

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
