# Build 0.3.0 — Overview Live Management Facade

## Scope

Overview is the Pulse management landing page. It follows the live Performance
state and projects existing operational status from Imports and Mapping. It has
no independent selectors, metric engine, ranking engine, period logic, or
freshness contract.

The accepted flow remains:

`Weekly cache -> Performance authority -> Overview`

with Imports and Mapping contributing their existing operational facades.
Overview never reads raw facts or weekly cache rows.

The final 0.3.0 presentation occupies the compact management footprint
`Overview!A1:J36`. Navigation is horizontal, Analysis Context presents Current,
Compare, and Latest available side by side, and the newer-week state appears
only in that context. Attention and Data Status share a compact lower block.

## Authority map

| Overview field | Authority |
| --- | --- |
| Current | Performance authority: `_Metric_Calc!AL33` |
| Compare | Performance authority: `_Metric_Calc!AL34` |
| Latest available | Active weekly-manifest output in Performance authority: `_Metric_Calc!AL36` |
| Period warning | Performance period authority: `_Metric_Calc!AL35` |
| Newer week available | Performance selection-recency output: `_Metric_Calc!AL37` |
| Total Sales | Selected-restaurant Current denominator: `_Metric_Calc!AL41` |
| Selected Category Sales | Selected-RPG Current numerator: `_Metric_Calc!AL42` |
| Sales Share | Performance selected-scope metric: `_Metric_Calc!AL43` |
| PP Change | Performance selected-scope metric: `_Metric_Calc!AL44` |
| NOK Impact | Performance selected-scope metric: `_Metric_Calc!AL45` |
| Selected restaurant/RPG scope | Performance authority: `_Metric_Calc!AL39:AL40` |
| Reporting Group Top/Bottom 3 | Performance ranking authority: `_Metric_Calc!AP33:AQ38` |
| Restaurant Top/Bottom 3 | Performance ranking authority: `_Metric_Calc!AU33:AV38` |
| Mapping attention | Existing Mapping health facade: `Mapping!D8`, `F6:F8`, `H6:H8`, `J6:J8` |
| Latest published / coverage / Imports | Existing Imports operational facade: `Imports!A9`, `D9`, `G9` |
| Classifications | Existing Mapping freshness facade: `Mapping!A9` |
| Performance availability | Performance authority: `_Metric_Calc!AL32` |

The new `_Metric_Calc` outputs are an upstream extension of Performance. They
reuse the accepted additive components and full-precision selected Display
values; they are not an Overview result table or a second metric engine.

## Ranking contract

Top and Bottom 3 follow the active Performance Display metric and the current
restaurant/Reporting Group selections. They use authoritative numeric helpers,
place unavailable values last, exclude Grand Total from restaurant ranking,
and use stable IDs for deterministic ties. Overview only formats the returned
name/value pairs.

## Unavailable behavior

`_Metric_Calc!AL32` is Available only when the active weekly cache is fresh,
both periods are valid, and the Performance scope contains at least one
restaurant and one Reporting Group. Otherwise selected Performance metrics and
rankings are blank upstream. Overview does not fall back to facts, cache rows,
or legacy results. Imports and Mapping status remain visible.

## Live Excel QA

1. Run `Build_0_3_0_Weekly_Performance` against canonical `Pulse_Current.xlsx`.
2. Run `Build_0_3_0_Overview`.
3. At 100% zoom, confirm Current, Compare, Latest available, five selected-
   performance values, scope, Top/Bottom 3, Attention, Data Status, and the five
   navigation links are readable without horizontal scrolling.
4. On Performance, switch Display between NOK Impact and PP Change. Confirm
   both Overview ranking sections update without rerunning a script.
5. Change a complete Current/Compare range and a restaurant/RPG selection.
   Confirm Overview follows recalculation and Performance remains authoritative.
6. Confirm a Current range ending before the latest published week displays
   `Newer week available` without changing the selection.
7. In a controlled reversible stale-state check, confirm selected-performance
   values and rankings suppress while Data Status reports Performance
   unavailable; restore the accepted Available state afterward.
8. Reconfirm Mapping QA 9/9, Phase 2C QA 16/16, Weekly Performance QA 16/16,
   Reports linkage, active/rollback cache authority, and six visible sheets.

No chart, Overview selector, full analytical table, future-domain placeholder,
or overall Total Sales growth metric is introduced in this slice.
