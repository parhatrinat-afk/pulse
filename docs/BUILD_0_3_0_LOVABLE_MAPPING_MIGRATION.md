# Build 0.3.0 Lovable mapping migration

## Scope and authority

This migration applies the frozen, human-approved Lovable business reporting
definitions to Pulse's existing hierarchical Mapping architecture. It is not a
Phase 2B metric or Performance cutover.

The approved translation is:

- Lovable KPI definition → existing Pulse Reporting Group ID;
- Lovable main member → SourceMainCategory Mapping Rule;
- Lovable sub member → SourceSubCategory Mapping Rule;
- Lovable item member → Product Mapping Rule; and
- approved item exclusion → Product Mapping Rule with `RuleAction=Exclude`.

The migration contains exactly 129 stable-ID decisions: 70 main mappings, 2
subcategory mappings, 49 Product mappings, and 8 Product exclusions. Runtime
names do not choose any rule. Names are read only after an approved ID is found
and are stored as audit/display metadata.

## Explicit Product exclusion contract

`RuleAction` is appended to `tblMappingRules`. Existing blank actions normalize
to `Map` and retain their prior meaning. `Exclude` is intentionally narrow:

- scope must be `Product`;
- `NodeID` must be the stable ProductID;
- `TargetReportingGroupID` must be blank;
- an active exclusion wins over inherited main/subcategory mapping;
- the result is `ResolutionStatus=Unmapped`,
  `ResolutionState=Explicit exclusion`, `ResolutionSource=Product`, and the
  exclusion's `MappingRuleID` remains `WinningRuleID`;
- overlapping active Product rules still produce Conflict; and
- deactivating the exclusion restores the nearest applicable ancestor.

No fake Unmapped Reporting Group is created. Excluded facts remain one-for-one
in the Phase 2A bridge and remain in source totals and denominators.

Approved exclusion ProductIDs:

- `PRD-000220`, `PRD-000221`;
- `PRD-000259`, `PRD-000260`;
- `PRD-000546`;
- `PRD-000566`, `PRD-000567`; and
- `PRD-000942`.

## Migration behavior

Run `office-scripts/Migrate_Lovable_Mapping.ts`. Before adding rows it validates:

- the 129-decision scope/action counts;
- uniqueness of every frozen semantic decision;
- presence of every approved stable node/ProductID;
- presence and active state of every target Reporting Group; and
- absence of a different active applicable rule at an approved node.

A semantically identical active applicable rule is reused. This includes
`MAP-000001`, the existing Add-ons main-category rule; the script fails rather
than duplicates it if reuse is not recognized. Inactive/expired audit history
is never rewritten or deleted. New rows receive new sequential MappingRuleIDs,
current effective date, active status, migration note, and explicit RuleAction.

The normal first run against the accepted checkpoint should report 129 logical
decisions, 1 reused rule, and 128 added rules. A direct rerun should report 129
reused and 0 added. Build-log entries are appended on successful runs.

## Accepted reconciliation targets for live QA

After migration, Phase 1 recomputation, and Phase 2A recomputation:

| State | Facts | Sales NOK | Quantity |
|---|---:|---:|---:|
| Source | 18,086 | 426,611,113.82 | 2,069,940.12 |
| Mapped | 15,935 | 405,493,085.13 | 1,925,100.06 |
| Unmapped | 2,151 | 21,118,028.69 | 144,840.06 |
| Conflict | 0 | 0.00 | 0.00 |
| Inactive Target | 0 | 0.00 | 0.00 |

Mapped coverage rounds to 88.11% of facts, 95.05% of Sales NOK, and 93.00% of
Quantity. Mapped plus Unmapped must reconcile exactly to Source.

| Reporting Group | Facts | Sales NOK | Quantity |
|---|---:|---:|---:|
| RPG-0001 Add-ons | 1,632 | 4,675,282.00 | 89,257.00 |
| RPG-0002 Non-Alcohol | 1,174 | 15,191,061.00 | 232,317.50 |
| RPG-0003 Spirits/Cocktails | 2,314 | 13,044,456.00 | 76,001.32 |
| RPG-0004 Coffee & Tea | 841 | 3,468,276.00 | 50,715.51 |
| RPG-0005 Beer & Cider | 461 | 22,134,730.00 | 161,469.47 |
| RPG-0006 Desserts | 389 | 14,476,736.00 | 80,901.73 |
| RPG-0007 Wine & Sake | 1,513 | 28,839,442.00 | 104,471.45 |
| RPG-0008 Starters | 1,266 | 23,863,705.00 | 135,290.75 |
| RPG-0009 Mains | 6,345 | 279,799,397.13 | 994,675.33 |

## Exact live Excel-for-web QA procedure

1. Save a QA copy of the accepted Phase 2A workbook. Keep GitHub as the source
   of truth for all three Office Scripts.
2. Add/run `Migrate_Lovable_Mapping.ts` once. Expect `129 decisions`, at least
   `MAP-000001` reused, and normally `128 added` against the clean checkpoint.
   If any `PUL-030M-*` error appears, stop; do not manually force rule rows.
3. In `Mapping Rules`, confirm `RuleAction` exists, `MAP-000001` remains a
   single active Add-ons Map rule, and the eight approved ProductIDs each have
   one active Exclude rule with a blank target. Preserve any older inactive
   audit rows.
4. Rerun `Migrate_Lovable_Mapping.ts` as the idempotency check. Expect `129
   reused, 0 added`; confirm the Mapping Rule row count did not increase.
5. Run the updated `Build_0_3_0_Phase1.ts`. Confirm dropdown validation reports
   `7/7`, `Mapping QA` has no conflicts or inactive targets, all nine checks are
   PASS, and fact row count remains 18,086.
6. Filter `Effective Mapping` to the eight approved ProductIDs. Each must show
   blank effective/target RPG, `Product`, `Explicit exclusion`, `Unmapped`, and
   its exclusion MappingRuleID as WinningRuleID. Confirm a normal positive
   Product override still reports `Explicit / Mapped`.
7. Run the updated `Build_0_3_0_Phase2A.ts`. It must accept today's Effective
   Mapping, create exactly 18,086 bridge rows, and produce a single mapping
   fingerprint consistently across the bridge and QA summary.
8. In `Metric Migration QA`, verify all reconciliation rows PASS, the four
   state totals equal the table above, and `tblMetricReportingGroupTotals`
   equals all nine RPG totals above. Confirm mapped plus unmapped fact count,
   Sales NOK, and Quantity equal source exactly.
9. In `_Metric_RPG_Facts`, filter the eight excluded ProductIDs. Their rows must
   be Unmapped with `ResolutionState=Explicit exclusion` and remain present in
   the bridge/source denominator.
10. Confirm source-fact protection and legacy-surface protection pass:
    `_Sales_Facts`, raw imports, `_Metric_Calc`, Performance, Reports, KPI
    Registry, and KPI-0001 must remain unchanged. The empty real-equivalence
    configuration may continue to show the previously accepted WARN.
11. Save the resulting workbook as a new QA checkpoint for human review. Do not
    begin Phase 2B and do not treat the workbook as a replacement for repository
    implementation sources.

Automated fixtures cover the resolver semantics, deterministic fingerprint,
denominator behavior, migration count/idempotency/reuse, contract parity, and
accepted reconciliation arithmetic. They do not replace the live Excel runtime
steps above.

## Accepted live Excel QA — 2026-08-12

The approved migration, updated Phase 1 resolver, and updated Phase 2A bridge
completed successfully in Excel for the web.

The accepted checkpoint confirmed:

- migration rerun was idempotent; Mapping Rules remained at `MAP-000133` and
  no additional rule was created;
- all nine Phase 1 Mapping QA checks passed;
- Conflict and Inactive Target were both zero;
- all 18,086 source facts were retained;
- Phase 2A materialized exactly 18,086 bridge rows;
- 15,935 facts were Mapped and 2,151 were Unmapped;
- Mapped Sales NOK was 405,493,085.13 and Unmapped Sales NOK was
  21,118,028.69, reconciling exactly to 426,611,113.82;
- Mapped Quantity was 1,925,100.06 and Unmapped Quantity was 144,840.06,
  reconciling exactly to 2,069,940.12;
- every generated reconciliation scope passed;
- source-fact protection passed;
- mapping fingerprint was `MAP-342029f71a922b47`;
- QA-0302A-08 correctly remained WARN because no human-authored legacy
  equivalence definitions exist; and
- the live totals exactly matched the approved projection and all nine
  Reporting Group targets above.

Scope coverage also reconciled exactly:

| Scope | Mapped facts | Unmapped facts | Total facts |
|---|---:|---:|---:|
| 2025 baseline | 12,485 | 1,869 | 14,354 |
| W31 2026 | 3,450 | 282 | 3,732 |
| Combined | 15,935 | 2,151 | 18,086 |

This live acceptance covers the Lovable mapping migration and the Phase 1 / 2A
recomputation path only. It does not authorize or begin Phase 2B.
