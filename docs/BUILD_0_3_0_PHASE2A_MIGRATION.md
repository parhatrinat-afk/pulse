# Build 0.3.0 Phase 2A migration and QA note

## Scope

Phase 2A establishes the deterministic Reporting Group metric contract and a
materialized analysis bridge. It does not cut `_Metric_Calc`, Performance,
Reports, or KPI-0001 over to Reporting Groups. That presentation migration is
Phase 2B.

Build 0.3.0 uses **current mapping state applied to historical facts for
analysis**. A mapping change may reclassify derived analysis, but it never edits
historical facts. Fact-date mapping/versioning is out of scope.

## Reproducible migration sequence

1. Start from `excel/Pulse_Build_0_2_0_QA.xlsx.xlsx`.
2. Run `office-scripts/Build_0_3_0_Phase1.ts`.
3. Confirm Phase 1 completed and `Effective Mapping` has today's `AsOfDate`.
4. Run `office-scripts/Build_0_3_0_Phase2A.ts`.
5. Inspect `Metric Contract`, `Metric Equivalence`, `_Metric_RPG_Facts`, and
   `Metric Migration QA`.

There is no manually saved Phase 1 workbook dependency. On a later day, or
after editing Mapping Rules, Reporting Groups, Products, or Source
Classifications, rerun Phase 1 before Phase 2A.

Phase 2A recomputes expected hierarchical resolution from the current Phase 1
inputs and compares it to `tblEffectiveMapping` before writing any Phase 2A
output. It fails with `PUL-0302A-001` when the materialized mapping is stale,
incomplete, duplicated, structurally invalid, or resolved for another day.

## Added workbook structures

### `Metric Contract` / `tblMetricContract`

Documents the reusable Phase 2A calculation contract for:

- Sales NOK;
- Quantity;
- Fact Count;
- Reporting Group Sales Share;
- Reporting Group Quantity Share.

Sales and Quantity Share are contract/QA capabilities only in Phase 2A.
Performance and KPI Registry exposure remain unchanged. Attach Rate is not
implemented.

### `_Metric_RPG_Facts` / `tblMetricRPGFacts`

Contains one derived row for every `tblSalesFacts` row. The join uses stable
`ProductID` and the validated current `tblEffectiveMapping` row.

Each derived row preserves:

- `SalesFactID`, `ImportID`, `RestaurantID`, and `ProductID` lineage;
- period, channel, publication state, SalesAmount, and Quantity;
- legacy `ReportingCategoryID` for migration comparison;
- effective Reporting Group and resolver audit fields;
- one explicit state: Mapped, Unmapped, Conflict, or Inactive Target;
- `MappingAsOfDate`;
- deterministic `MappingFingerprint`;
- `MetricRefreshAt`.

The bridge is regenerated on every Phase 2A run. It is derived analysis state,
not a second fact authority.

### `Metric Equivalence` / `tblLegacyRPGEquivalence`

Human-owned legacy CAT-to-RPG comparison definitions. Definitions are
ID-based; Pulse never infers equivalence from display names.

Multiple rows may share a `DefinitionID` when several legacy CAT IDs together
form one comparison definition. Every row must provide:

- one `ReportingGroupID`;
- one `LegacyReportingCategoryID`;
- a consistent `ComparisonStatus`: Equivalent, Partial, or Not Comparable;
- `Active` = Yes or No;
- optional notes and approval metadata.

Invalid definitions are surfaced as QA failures and comparison output is not
calculated until the configuration is corrected. An empty table is valid and
reported as a warning because no equivalence may be inferred automatically.

### `Metric Migration QA`

Contains:

- `tblMetricQA` — central Phase 2A checks;
- `tblMetricReconciliation` — source/bridge/state reconciliation by import and
  optional channel scope;
- `tblMetricReportingGroupTotals` — independently inspectable mapped totals;
- `tblLegacyRPGComparison` — human-configured side-by-side numerator,
  denominator, share, and variance results;
- `tblLegacyRPGCrosswalk` — CAT-to-RPG/state membership explaining differences.

The validated channel behavior is preserved: `All channels` is calculated by
omitting a channel restriction, and actual channel scopes are added for QA.
Phase 2A does not add mandatory channel selectors to Performance.

## Mapping fingerprint

The semantic v1 fingerprint is deterministic and row-order independent. It
includes:

- mapping `AsOfDate`;
- Reporting Group IDs, names, active status, and sort order;
- classification-affecting Mapping Rule fields;
- Product IDs and hierarchy node identity;
- recomputed effective resolution state.

Notes, timestamps, and other non-semantic audit text do not change the
fingerprint. The fingerprint is stored on every bridge row and in the QA
summary.

## Reconciliation contract

For every generated scope:

`Source facts = Metric bridge facts`

and:

`Mapped + Unmapped + Conflict + Inactive Target = Source facts`

The equality is checked independently for:

- fact count;
- Sales NOK;
- Quantity.

All source fact rows are bridged. Normal metric scopes remain restricted to
`PublicationState = Active Finalized`; the global all-facts row proves complete
lineage coverage.

## Preserved legacy behavior

Phase 2A snapshots and rechecks `_Metric_Calc`, Performance, Reports, and KPI
Registry during execution. It does not change:

- the legacy `ReportingCategoryID` fact field;
- Category Sales Share formulas or values;
- current/comparison dataset behavior;
- Performance or Reports layout;
- channel UI;
- KPI-0001 metadata or availability.

## Validation still requiring Excel

Automated tests validate the deterministic contract and Office Script source,
but the following require Excel for the web / Work-mode validation:

- Office Scripts compiler/runtime compatibility;
- runtime and memory behavior when writing and rereading 18,086 bridge rows;
- table creation and rerun behavior;
- stale-mapping failure followed by successful Phase 1/Phase 2A refresh;
- preservation of human equivalence rows across reruns;
- visual inspection of the new contract, equivalence, bridge, and QA sheets;
- verification that existing Performance and Reports values are unchanged in
  the migrated workbook.
