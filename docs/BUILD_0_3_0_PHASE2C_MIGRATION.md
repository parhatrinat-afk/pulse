# Build 0.3.0 Phase 2C — Interactive Sales Performance Model

## Scope

Phase 2C adds an Excel-recalculation interaction layer downstream of the
accepted Phase 2B metric results. It does not rematerialize facts, mapping,
Reporting Groups, or metric results.

Authoritative flow:

`_Sales_Facts → _Metric_RPG_Facts → tblMetricRPGResults → Phase 2C additive helpers → Performance → Reports`

After the model is refreshed, restaurant selection, Reporting Group
multi-selection, current dataset, comparison dataset, matrix display mode, sort
target, and sort order all update through Excel formulas. No Office Script is
required for normal exploration.

Phase 2C adds NOK Impact only as a derived display over existing KPI-0001
components. It does not add Attach Rate, a new KPI, channel selection, the
future Pulse refresh control, a dashboard redesign, or Phase 3 work.

## Prerequisite and reproducible sequence

Starting from the validated 0.2.0-QA checkpoint:

1. run the accepted Lovable migration when the approved rules are absent;
2. run Phase 1;
3. run Phase 2A;
4. run Phase 2B;
5. verify all Phase 2B Metric Results QA checks pass; and
6. run `Build_0_3_0_Phase2C.ts` once to install or refresh the interaction
   layer.

Phase 2C validates the complete Phase 2B grain and deterministic result IDs
before workbook mutation. For every dataset and active Reporting Group, the
sum of all eligible Restaurant-scope numerators and denominators must equal
the retained Company control row. A stale or inconsistent model fails with
`PUL-0302C-004` and directs the user to rerun Phase 2B.

## Selection contract

`tblRestaurants[Status]=Active AND ReportingEnabled=Yes` remains the
refresh-time eligibility contract. Phase 2C never exposes an ineligible
RestaurantID in normal Performance selection.

Performance adds two stable-ID configuration tables:

- `tblPerformanceRestaurantSelection`: Include, Restaurant, RestaurantID;
- `tblPerformanceRPGSelection`: Include, Reporting Group, ReportingGroupID.

The Include state is authoritative. Rows marked Yes participate; rows marked No
do not. There is no separate All/Custom user control. When every eligible row is
Yes the derived scope is All; otherwise it is a selected subset.

On the first Phase 2C run, all eligible IDs default to Yes. On later reruns,
choices are preserved by stable ID. A newly eligible ID defaults to No so an
existing selected scope cannot expand unnoticed. Display-name changes do not
change selection identity. A one-time compatibility step preserves the prior
effective All scope by setting existing Include rows to Yes when migrating from
the removed mode controls.

The single detail Reporting Group selector remains separate from matrix RPG
multi-selection. A detail group therefore remains inspectable even when its
matrix Include value is No.

## Additive calculation model

Phase 2C reads only the accepted Restaurant rows of `tblMetricRPGResults` for
interactive calculation. For each selected RestaurantID, ReportingGroupID, and
ImportID, `_Metric_Calc` retrieves:

- NumeratorSalesNOK;
- DenominatorSalesNOK; and
- canonical share = numerator ÷ denominator, or zero when denominator is zero.

The Company rows remain materialized and unchanged as Phase 2B QA/control
totals. They are not used as an arbitrary-selection shortcut.

For a selected restaurant set `S`, the combined result is:

`SUM(NumeratorSalesNOK for S) ÷ SUM(DenominatorSalesNOK for S)`

This is valid for KPI-0001 because its numerator and denominator are additive
across disjoint restaurant scopes. It is not an average of restaurant shares.
It must not be assumed valid for a future KPI unless that KPI's approved
definition exposes compatible additive components.

All mapping states, including Unmapped, Conflict, Inactive Target, and explicit
exclusions, remain represented in the Phase 2B denominator components. Phase
2C cannot remove them from a selected scope.

## Workbook implementation

### `_Metric_Calc`

Phase 2C owns a helper surface beginning at AE. It contains:

- formula-filtered selected Restaurant names and IDs;
- formula-filtered selected Reporting Group names and IDs;
- current/comparison ImportIDs and selection counts;
- six bounded matrices: current numerator, current denominator, current share,
  comparison numerator, comparison denominator, and comparison share; and
- one bounded selected-display matrix that remains numeric and chooses among
  those authoritative components from the Display selector;
- four selected-RPG Total components plus one numeric Total display value; and
- one numeric sort key per canonical RestaurantID plus a separate sorted
  RestaurantID presentation list.

The six matrices cover active Reporting Groups and eligible individual
restaurants only. Their final row is a sum-of-components total. No restaurant
combination is materialized, and no helper formula reads `_Sales_Facts` or
`_Metric_RPG_Facts`.

### `Performance`

The existing detail, current dataset, comparison dataset, comparison check,
and Explain concepts remain. Phase 2C adds:

- matrix display selector;
- Sort by and Order selectors;
- the two adjacent stable-ID Include tables; and
- a formula-driven Restaurant × (Total + Reporting Group) matrix with Grand
  Total.

Supported matrix displays are:

- PP Change;
- Current Share;
- Comparison Share;
- Current Sales NOK; and
- NOK Impact.

For a zero relevant denominator, the canonical helper remains zero and the
user-facing detail/matrix displays an em dash (`—`). PP Change is expressed in
percentage points. Share and PP Grand Totals derive from summed numerators and
denominators; sales displays use summed numerator components.

Total aggregates exactly the Reporting Groups currently selected for the
matrix. Current and comparison denominators are identical across RPGs at a
given dataset/restaurant scope and are used once, never summed once per RPG.
The Total semantics by Display mode are:

- Current Share = selected current numerator ÷ current denominator;
- Comparison Share = selected comparison numerator ÷ comparison denominator;
- PP Change = (aggregate current share − aggregate comparison share) × 100;
- Current Sales NOK = selected current numerator, with no denominator
  dependency; and
- NOK Impact = selected current numerator − aggregate comparison share ×
  current denominator.

The same formulas apply to Grand Total after first aggregating the selected
restaurant components. No percentage, PP value, or restaurant-level impact is
summed. Selecting every active RPG does not imply 100% because Unmapped,
Conflict, Inactive Target, and explicitly excluded facts remain in the
denominator.

NOK Impact is current selected-scope Reporting Group Sales NOK minus comparison
Reporting Group share × current selected-scope total Sales NOK. At Grand Total
grain the exact calculation is aggregated current numerator minus
((aggregated comparison numerator ÷ aggregated comparison denominator) ×
aggregated current denominator). It never sums restaurant-level impacts and
never reads, sums, or parses visible matrix strings. Differing period lengths
remain permitted because the comparison supplies a share baseline and the
current scope supplies the monetary base; the existing factual warning remains.

The visible matrix is a presentation-only text facade over the bounded numeric
selected-display matrix. It uses `FIXED` decimal-count arguments to render PP,
share, and NOK labels without applying custom or conditional number formats to
the visible cells. Conditional formatting is limited to visual styling.

The six canonical component matrices never become text and never read the
visible facade. Grand Total remains numeric in the helper layer. Reports, QA,
and any arithmetic continue to consume numeric helpers or the separate numeric
detail cells. The text facade exists only because Excel for the web interpreted
decimal placeholders incorrectly when number formats were switched through
conditional formatting, including when supplied with the workbook-local detail
format.

`FIXED` delegates decimal and thousands separators to Excel rather than a
repository-authored custom format string. The Norwegian Excel-for-web
checkpoint must therefore confirm comma decimals and local thousands grouping;
static tests intentionally do not manufacture a separator override.

Sorting changes only the presentation lookup order. The canonical
RestaurantID/component/helper rows retain their Phase 2B order. Sort by offers
Total and every active RPG; Order offers Highest first and Lowest first. The
default is Total / Highest first. Sort keys use the full-precision bounded
numeric selected-display helper, never the visible text facade. Unavailable
keys remain last, exact ties use deterministic RestaurantID ordering, and Grand
Total remains fixed outside the sort range.

If an RPG retained in Sort by is removed from the visible matrix through its
Include state, its selector value is preserved. Performance visibly reports that
the target is not displayed and uses Total until the RPG is selected again.
Display or dataset changes immediately recalculate the numeric keys and visible
order without an Office Script.

### `Reports`

Reports continues to link to Performance rather than calculating a second
result. It now also exposes the selected interactive restaurant-scope summary.

### `Metric Results QA`

`tblMetricResultsQA` remains the Phase 2B QA table. Phase 2C appends the
separate `tblPerformanceInteractionQA`, covering the centralized grain,
Company controls, stable-ID catalogs, additive-only formulas, Grand Total,
selected-RPG Total, numeric presentation sorting, zero denominators,
display modes, authoritative Include-state selection, Reports linkage,
validation, and protected surfaces.

QA-0302C-09 checks that all five Display modes exist while numeric helpers stay
separate from the `FIXED` text facade. Its token checks must distinguish the
legitimate mode label `Current Sales NOK` from the exact presentation suffix
literal `" NOK"`; the broader substring ` NOK` causes a false failure.

## UX audit outcome — future guidance only

The accepted workbook audit confirms Performance is the strongest current
user-facing surface and its calculation architecture should remain intact.
Future work may improve its visual hierarchy—especially Total, Grand Total, and
selection panels—without replacing the validated helpers.

The eventual normal-user workbook should be substantially smaller than the
current engineering workbook: approximately Overview, Performance, Reports,
Imports/Refresh, and Mapping. Restaurants, Reporting Groups, Settings, and KPI
Registry may remain administration surfaces. Engineering, QA, raw, helper,
bridge, audit, legacy CAT/remap, environment, metric, adapter, and staging
sheets should later be hidden/protected from normal users but retained for
advanced/admin troubleshooting.

Overview should eventually be a live, minimal, clickable home screen and the
natural location for a future Pulse ♥ refresh workflow. Reports should evolve
into a meeting/export surface. Mapping should retain its hierarchical rules and
lineage underneath a simpler needs-attention → choose Reporting Group → refresh
experience, with legacy remap surfaces hidden from normal navigation. Imports
should eventually present upload/import → Pulse ♥ → validation → refreshed
model while preserving the current engineering workflow underneath. Detailed
QA should eventually roll up to simple readiness/issues indicators.

No sheet is hidden, removed, protected, or redesigned by Phase 2C. Pulse ♥ and
the broader workbook UX remain separately scoped future work.

## Recalculation size

At the accepted sixteen-restaurant/nine-RPG checkpoint, the six component
blocks contain 918 bounded formula cells and the per-RPG numeric selected-
display helper contains 153 cells. The selected-RPG Total layer adds 85 cells
(four components and one numeric display across sixteen restaurants plus Grand
Total); sorting adds 32 cells (one key and one presentation-order ID across
sixteen restaurants). The visible text facade contains at most 170 cells: Total
plus nine RPG columns across sixteen restaurants plus Grand Total. Selection
helpers add only one row per eligible ID. The model uses FILTER, SORTBY,
INDEX/MATCH, SUMIFS, SUM, COUNTIF, IF, and FIXED; it has no volatile OFFSET or
INDIRECT formulas and does not scan the 18,086 fact rows during interaction.

Formula size grows linearly with eligible restaurants and active Reporting
Groups. It does not grow with the number of possible restaurant subsets.

## Edge-case behavior

- Zero selected restaurants: detail values show `—`; the matrix is blank.
- One selected restaurant: Grand Total equals that Restaurant-scope result.
- All restaurants: Grand Total equals the retained Phase 2B Company control.
- Zero selected RPGs: matrix is blank; the independent detail selector remains
  usable.
- One or all RPGs: the matrix shows exactly the selected active columns.
- One selected RPG: Total equals that RPG for every display mode.
- Multiple selected RPGs: Total aggregates exactly those RPG numerators while
  using each current/comparison denominator once.
- A hidden selected sort RPG falls back visibly to Total without discarding the
  requested sort target.
- Highest/Lowest sort only selected restaurants; unavailable values remain last
  and Grand Total remains fixed.
- No selected-dataset rows for a restaurant: denominator is zero and the cell
  displays `—`.
- Same current/comparison dataset: shares are equal, PP Change is zero where
  denominators exist, and NOK Impact is zero; the comparison check says the
  dataset is the same.
- Current and comparison selections remain independent; no dataset is
  inherently current or baseline.

## Protection and rerun behavior

Phase 2C fingerprints facts, Phase 2A bridge, Phase 2B results, imports,
Reporting Groups, restaurants, Mapping Rules, Effective Mapping, KPI Registry,
and Phase 2B QA before and after presentation writes. Any mutation is a hard
failure.

The script is rerunnable. It preserves selection choices by stable ID, replaces
only the Phase 2C selection tables/helper/presentation/QA surfaces, and records
the successful installation in Environment and Build Log.

Office Scripts is not transactional. A later Excel runtime failure can leave a
partial presentation/helper surface, but it cannot validly pass the protected
fingerprint check. Correct the surfaced error and rerun Phase 2C; do not edit
authoritative results to work around it.

## Live Excel QA procedure and accepted evidence

1. Complete the accepted Phase 1 → Phase 2A → Phase 2B refresh sequence and
   confirm all Phase 2B QA checks pass.
2. Run Phase 2C in Excel for the web. Confirm it reports 8/8 validation ready
   and all Phase 2C QA rows are PASS.
3. Confirm all eligible Restaurant and active RPG Include rows initially show
   Yes. Performance should report all selected, show the Total column, all nine
   active RPGs, and one Grand Total row. Sort by should default to Total /
   Highest first.
4. Confirm the all-selected Add-ons detail still equals the accepted Phase 2B
   values: W31 2026 NOK 131,487 / 1.47%; 2025 1.09%; approximately +0.38 pp.
5. Mark all but an arbitrary three-restaurant set No. Record their Phase 2B
   components and verify Grand Total equals summed numerator ÷ summed
   denominator, not the average of their displayed shares.
6. Select one restaurant and verify its Grand Total exactly equals its
   Restaurant row.
7. Deselect and reselect Swift (when eligible). Verify values update immediately
   without rerunning an Office Script and return exactly after reselection.
8. Mark all but one RPG No and confirm Total equals that RPG in all five
   displays. Add a second RPG and independently verify Total from the two
   selected numerators and one denominator. Restore all to Yes and confirm Total
   share does not become 100% merely because all mapped RPGs are selected.
9. Cycle all five display modes and verify the visible text facade. Use a zero-
   denominator scope if available and confirm `—` is displayed. In the
   Norwegian workbook, confirm PP Change and both Share modes use comma
   decimals with exactly two places (for example `+0,38 pp` and `1,47%`), and
   confirm Current Sales NOK uses the expected local thousands grouping and
   ` NOK` suffix. Confirm NOK Impact uses the same local grouping, an explicit
   positive or negative sign when nonzero, and `0 NOK` at zero. With W31 2026
   current and 2025 comparison, all-selected Add-ons should be approximately
   `+34 019 NOK` using the workbook's local grouping separator. No display
   should show three-digit zero padding.
10. Swap current/comparison datasets and confirm Add-ons becomes approximately
    `-1 585 928 NOK` because the monetary base changes to 2025 total sales. Then
    select the same dataset on both sides and confirm zero PP and `0 NOK` Impact.
    Confirm a period-length mismatch remains a warning and does not block NOK
    Impact.
11. With Sort by Total, verify Highest and Lowest reverse the available
    restaurant order while unavailable values remain last and Grand Total does
    not move. Use a tied fixture if available and confirm deterministic
    RestaurantID ordering.
12. Sort by Add-ons in PP Change and NOK Impact, then change Display. Confirm the
    order follows full-precision numeric values rather than rounded visible
    strings. Change either dataset and confirm the order recalculates without a
    script.
13. Mark the RPG retained in Sort by No. Confirm Sort status
    visibly says Total is being used, the selector remains unchanged, and the
    RPG sort resumes when the RPG is restored.
14. Set all restaurant Include values to No, then all RPG Include values
    to No. Confirm the documented blank/dash guards and restore selections.
15. Confirm Reports matches the Performance detail values and restaurant scope.
16. Confirm ineligible `ReportingEnabled=No` restaurants are absent from the
    selection table after a proper Phase 2B/2C refresh.
17. Confirm Phase 1, Phase 2A, Phase 2B, source facts, mapping fingerprint,
    `tblMetricRPGResults`, and reconciliation totals remain unchanged.

### Accepted live result — 2026-08-13

The canonical `OneDrive/Pulse/Development/Pulse_Current.xlsx` passed all 16
Phase 2C Interaction QA checks. Restaurant and Reporting Group selection,
selected-RPG Total, weighted Grand Total, full-precision numeric sorting, the
five Display modes, NOK Impact, dataset changes, and Reports linkage all
recalculated without rerunning an Office Script.

The accepted all-selected Add-ons result remained `1.47%` current share,
`1.09%` comparison share, `+0.38 pp`, and `131,487 NOK`. The Swift restaurant
selection round trip restored those values exactly. Subsequent IA and bounded
visual slices retained 16/16 PASS and established the clean functional 0.3.0
workbook foundation without changing Phase 2C calculations.
