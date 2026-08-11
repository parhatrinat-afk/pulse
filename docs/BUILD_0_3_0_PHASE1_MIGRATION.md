# Build 0.3.0 Phase 1 migration note

Starting checkpoint: `excel/Pulse_Build_0_2_0_QA.xlsx.xlsx`.

Run `office-scripts/Build_0_3_0_Phase1.ts` in Excel for the web. The script is
safe to rerun: user edits in `Reporting Groups` and persisted rows in `Mapping
Rules` are retained, while computed browse, effective-mapping, and QA surfaces
are refreshed.

## Legacy handling

`Reporting Categories` remains the 0.2.0 source-default compatibility layer and
is marked as legacy. It is not renamed or converted into Reporting Groups.

The existing `Remap Rules` and `Effective Categories` structures are preserved.
The validated checkpoint contains no persisted legacy remap rows, so Phase 1
does not invent a semantic conversion from legacy Reporting Category targets to
the new business Reporting Groups. New authoritative rules are stored in
`tblMappingRules` and target `ReportingGroupID`.

## Phase boundary

Phase 1 does not migrate `_Sales_Facts[ReportingCategoryID]`, `_Metric_Calc`,
Performance, Reports, or KPI-0001. The 0.2.0 Category Sales Share path remains
unchanged until Phase 2 metric integration.

The script snapshots fact row count, total quantity, total sales, Active
Finalized totals, and boundary fact IDs before and after migration and fails if
they differ.

## Office Scripts runtime compatibility

List validations read the current values from `_Mapping_Lists` and pass them to
`ListDataValidation.source` as comma-separated literal strings. This follows the
documented Office Scripts pattern while keeping active Reporting Group IDs
derived from the human-editable `Reporting Groups` registry. Each dropdown is
wired independently; a rejected nonessential validation is surfaced as
`PUL-0301-013` on `Mapping` and in the script result without blocking mapping
recomputation.

Map and Set collections are traversed with `forEach` callbacks. Iterator-based
`for...of`, `entries`, and `Array.from` patterns are avoided for compatibility
with the Office Scripts TypeScript compiler target.
