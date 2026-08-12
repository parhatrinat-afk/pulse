# Build 0.3.0 Phase 2B migration and QA note

## Scope

Phase 2B migrates the existing KPI-0001 calculation and presentation path from
legacy CAT membership to the authoritative Reporting Group layer. It reuses
the validated Phase 2A bridge and does not resolve mapping independently.

Authoritative flow:

`_Sales_Facts → ProductID → Effective Mapping → _Metric_RPG_Facts → tblMetricRPGResults → Performance → Reports`

Phase 2B does not redesign Performance, add Attach Rate, restore mandatory
channel controls, infer CAT/RPG equivalence, increase mapping coverage, or
modify raw/fact/mapping/legacy CAT structures.

## Reproducible sequence

Starting from the validated 0.2.0-QA checkpoint:

1. run the accepted `Migrate_Lovable_Mapping.ts` when the approved rules have
   not already been populated;
2. run the accepted `Build_0_3_0_Phase1.ts`;
3. run the accepted `Build_0_3_0_Phase2A.ts` on the same day;
4. verify Phase 2A reconciliation passes and mapping fingerprint is current;
5. run `Build_0_3_0_Phase2B.ts`.

Phase 2B rejects a prior-day or inconsistent bridge before any presentation
write. Rerun Phase 1 and Phase 2A after any change to Mapping Rules, Reporting
Groups, Products, Source Classifications, or Effective Mapping.

## Central result table

`_Metric_Calc` retains dataset helpers in A:E and restaurant helpers in L:M.
Its former CAT selector block in I:J becomes the active Reporting Group helper.

`tblMetricRPGResults` is added at O:AC with columns:

- MetricResultID;
- MetricID;
- ImportID;
- ReportingGroupID;
- ScopeType, RestaurantID, and RestaurantScopeFingerprint;
- ChannelScope and PublicationState;
- NumeratorSalesNOK and DenominatorSalesNOK;
- MetricValue;
- MappingAsOfDate and MappingFingerprint; and
- CalculatedAt.

The current result grain is:

`1 KPI × 2 datasets × 9 active RPGs × (1 company + 16 restaurants) = 306 rows`

Cardinality is configuration-driven. Fifteen enabled restaurants produce:

`1 KPI × 2 datasets × 9 active RPGs × (1 company + 15 restaurants) = 288 rows`

Result IDs are deterministic hashes of stable metric, dataset, Reporting Group,
scope, channel-scope, and publication-state identity. Company identity includes
the deterministic fingerprint of sorted enabled RestaurantIDs. Display names
are not part of identity.

## Metric and denominator contract

The active KPI remains `KPI-0001`, renamed to Reporting Group Sales Share.

- Numerator: Mapped Sales NOK for the selected ReportingGroupID.
- Denominator: all Sales NOK in the identical selected scope.
- Zero denominator: metric value is zero.

`tblRestaurants[Status]=Active AND ReportingEnabled=Yes` defines the shared,
authoritative Performance restaurant set. It is organizational-scope
configuration rather than KPI-0001 business logic. Facts outside that set are excluded
from both Company numerator and denominator and do not receive Restaurant-scope
results. They remain unchanged in `_Sales_Facts`, `_Metric_RPG_Facts`, Phase 2A
reconciliation, and mapping structures.

Future Performance KPIs must consume the same contract unless their approved
definition explicitly documents a different organizational scope. Phase 2B
does not implement those future KPIs.

Within the enabled set, Unmapped, Conflict, Inactive Target, and explicit
exclusions are never removed from the denominator. Only Mapped facts can enter
an RPG numerator.

The presentation remains all-channel. The materialized result records
`ChannelScope=All channels`; no B11/G11 channel selector is introduced.

## Performance and Reports

Performance retains its existing flow and visual language:

- selected KPI and Reporting Group;
- independent current and comparison datasets;
- selected result;
- factual comparison-scope check;
- company overview;
- selected-group restaurant breakdown; and
- Explain section.

The company overview now contains active Reporting Groups rather than roughly
130 source categories. The existing layout compacts around nine rows without
adding controls or redesigning the page.

All result formulas read `tblMetricRPGResults`. Reports retains its existing
links to Performance and therefore cannot diverge by reimplementing the metric.

## Bridge and source protection

Before mutation, Phase 2B validates:

- today's Effective Mapping and bridge MappingAsOfDate;
- semantic mapping fingerprint parity;
- unique ProductID and SalesFactID membership;
- exactly one bridge row per source fact;
- source/bridge lineage and measures;
- Sales NOK, Quantity, and fact-count equality;
- complete Mapped/Unmapped/Conflict/Inactive Target coverage; and
- Performance-enabled scope plus excluded Performance scope equals the complete
  bridge for fact count, Sales NOK, and Quantity;
- every proposed central result against an independent in-memory calculation.

After the cutover it re-fingerprints `_Sales_Facts`, `_Metric_RPG_Facts`,
Mapping Rules, Reporting Groups, Effective Mapping, Restaurants, Products, Source
Classifications, legacy Reporting Categories, and the human-authored
equivalence table. Any change is a hard failure.

`Metric Results QA` / `tblMetricResultsQA` records the successful runtime
checks and visibly reports selector-validation warnings without invalidating
otherwise correct centralized results.

Changing `ReportingEnabled` takes effect when Phase 2B is rerun. Phase 2B does
not add an interactive Performance restaurant selector.

## Accepted projected validation targets

For RPG-0001 Add-ons:

| Dataset | Sales NOK | Sales share |
|---|---:|---:|
| Week 31 2026 | 131,487.00 | 1.47% |
| 2025 full year | 4,543,795.00 | 1.09% |

The current-minus-comparison change is approximately +0.38 percentage points.
These values intentionally differ from the old single-CAT Add-ons calculation.
They remain the expected targets while all sixteen accepted restaurants are
enabled. Deselecting a restaurant intentionally changes Company values.

Full accepted source coverage remains:

| State | Facts | Sales NOK | Quantity |
|---|---:|---:|---:|
| Mapped | 15,935 | 405,493,085.13 | 1,925,100.06 |
| Unmapped | 2,151 | 21,118,028.69 | 144,840.06 |
| Conflict | 0 | 0.00 | 0.00 |
| Inactive Target | 0 | 0.00 | 0.00 |
| Source | 18,086 | 426,611,113.82 | 2,069,940.12 |

## Accepted live Excel QA — 2026-08-12

The Phase 2B Office Script compiled and completed in Excel for the web. With
all sixteen active ReportingEnabled restaurants:

- `tblMetricRPGResults` contained 306 rows;
- excluded Performance scope contained zero facts;
- Company `RestaurantScopeFingerprint` was `RSC-08df626f217dd94b`;
- W31 2026 Add-ons was NOK 131,487 / 1.47%;
- 2025 comparison was 1.09%; and
- change was approximately +0.38 percentage points.

Swift was then set to `ReportingEnabled=No` and Phase 2B was rerun:

- cardinality changed to 288 rows / fifteen restaurants;
- 764 facts, NOK 12,561,017.69 Sales, and Quantity 3,008 moved to excluded
  Performance scope without changing Phase 2A;
- Company Add-ons changed to NOK 115,607 / 1.33%;
- 2025 comparison changed to 1.12%; and
- Swift disappeared from the restaurant breakdown.

After Swift was re-enabled and Phase 2B rerun, cardinality returned to 306,
excluded facts returned to zero, the Company scope fingerprint returned exactly
to `RSC-08df626f217dd94b`, and the accepted Add-ons results returned exactly.
All sixteen Metric Results QA checks remained PASS throughout the round trip.

## Excel-for-web runtime note

Selector dropdowns use small literal lists generated from the authoritative
`_Metric_Calc` helper values. This follows the runtime-compatible validation
pattern already established by Phase 1 after cross-sheet Range-backed list
validation proved unreliable in Excel for the web. Business names and IDs are
not hard-coded into metric logic. A validation failure is surfaced in Metric
Results QA and the script return message.

Office Scripts is not transactional. All deterministic bridge/result checks
run before workbook writes, but a later Excel runtime failure during UI/table
formatting can leave partial Phase 2B presentation output. A rerun is
idempotent: it replaces `tblMetricRPGResults`, Performance, Reports, and Metric
Results QA from the still-protected source/bridge state.
