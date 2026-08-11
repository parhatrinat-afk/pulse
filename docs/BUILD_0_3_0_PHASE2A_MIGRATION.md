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

## Excel-for-web runtime hardening

The first live Phase 2A run against the 18,086-row checkpoint surfaced a queued
`Worksheet.getRange` `GeneralException` at the next workbook read. The generated
bridge geometry was within Excel limits (`A4:T18090`), so the failure was not a
row-count overflow. Phase 2A no longer uses string-address
`Worksheet.getRange()` calls for output creation or formatting. Output ranges
are acquired by bounded row/column indexes, output row widths are validated,
and range failures report the sheet, computed A1 equivalent, and operation as
`PUL-0302A-020` or `PUL-0302A-021`.

Workbook/table reads used for the legacy-surface fingerprint are performed
before in-memory hashing rather than from inside the hashing loop. Bridge data
continues to be written in bounded chunks because a single 18,086-by-20 write
is less reliable in Excel for the web.

Office Scripts execution is not transactional. A failed run may therefore
leave `Metric Contract`, `Metric Equivalence`, or `_Metric_RPG_Facts` partially
created or formatted, while the later `Metric Migration QA`, environment, and
build-log updates may be absent. A rerun clears and rebuilds the generated
contract, bridge, and QA surfaces; the human-owned equivalence table is
preserved. Raw data, `_Sales_Facts`, `_Metric_Calc`, Performance, Reports, and
KPI-0001 are outside those output writes.

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

## Live Excel QA — 2026-08-11

`Pulse_Build_0_3_0_Phase2A_QA.xlsx` was rebuilt and reviewed in Excel for the
web using the Phase 2A source at commit
`045b8938159d544bfc75206c9ca735d8da11d609`.

The final clean run confirmed:

- the Office Script completed without the prior iterator or
  `Worksheet.getRange` runtime failures;
- MappingAsOfDate was 2026-08-11 and the reproducible mapping fingerprint was
  `MAP-3416c94758ea1743`;
- `tblMetricRPGFacts` occupied `A4:T18090` and contained exactly 18,086 data
  rows, one for each source fact;
- all-facts Sales NOK reconciled exactly at 426,611,113.82;
- all-facts Quantity reconciled at 2,069,940.12, subject only to the displayed
  floating-point residue already accepted by the deterministic tolerance;
- mapped coverage was 521 facts, Sales NOK 1,013,589.00, and Quantity
  20,171.11;
- unmapped coverage was 17,565 facts, Sales NOK 425,597,524.82, and Quantity
  2,049,769.01;
- Conflict and Inactive Target coverage were both zero for this clean
  checkpoint;
- all generated all-facts, import, and channel reconciliation rows reported
  PASS for fact count, Sales NOK, Quantity, and four-state coverage;
- source-fact protection reported PASS with fingerprint
  `DATA-19428a5949d194ed`;
- `_Metric_Calc`, Performance, Reports, and KPI Registry protection reported
  PASS;
- Reporting Group totals were materialized for `RPG-0001` / Add-ons and
  reconciled to the mapped coverage;
- the temporary `EQ-QA-0001` QA row was removed without replacing the
  human-owned equivalence table, then Phase 2A was rerun;
- `tblLegacyRPGEquivalence` was empty after cleanup,
  `tblLegacyRPGComparison` contained no inferred definitions, and
  QA-0302A-08 returned the expected WARN: no equivalence definitions entered
  yet.

This evidence validates the Phase 2A runtime and clean checkpoint only. It does
not activate the Reporting Group metric path in Performance or begin Phase 2B.

## Remaining Excel validation for later migration work

Automated tests validate the deterministic contract and Office Script source,
and the 2026-08-11 live run validates Phase 2A against the current checkpoint.
Future mapping changes, real human-authored equivalence definitions, and the
Phase 2B metric cutover require their own Excel QA before release.
