# Build 0.3.0 UX — Visual Slice 2B

Visual Slice 2B is the rerunnable pre-checkpoint presentation cleanup after
accepted Phase 2C, UX IA Slice 1, and UX Visual Slice 2A.

## Owned ranges

- `Performance!A2:T2`: keep the accepted subtitle on one visible line across
  its existing blank band; set row 2 to 28 points.
- `Performance!A4:T47`: compact existing rows, widths, and spacing without
  moving cells; preserve the matrix and compact Explain block.
- `Performance!N4:P21` and `Performance!R4:T14`: narrow ID/input columns,
  center readable body rows, and simplify output fills.
- `Mapping!A2:N2` and `Mapping!A4:N15`: compact the subtitle/action area and
  show workflow guidance across the existing blank status band without overlap.
- `Mapping!B10`: retain the Excel date value and display it as `dd.mm.yyyy`.

Excel row height applies to the entire worksheet row. Performance spacing is
therefore coordinated across its left controls and right selection tables. No
cells, tables, formulas, or functional surfaces move.

The restrained color roles are navy for major headers, blue for interactive
section headers, pale yellow only for editable/input cells, white or light
neutral for calculated outputs, and pale green only for the ready Mapping
status. Existing negative-value styling remains unchanged.

## Frozen behavior

The repair does not write values or formulas, resize tables, change validation,
change selections, alter sheet visibility/order, or touch calculations,
mapping, imports, Reports linkage, Total, Grand Total, sorting, or NOK Impact.

Before and after repair, the script checks Performance, Mapping, and Reports
values/formulas, the Slice 2A protected table content/schema fingerprints, both
selection validation rules, the six-sheet IA state, and all 16 Phase 2C QA
results. Postconditions require the compact row/column geometry, subtitle and
workflow wrapping behavior, selection alignment, and real Mapping date display.

## Live QA

Run `office-scripts/Build_0_3_0_UX_Visual_Slice2B.ts` only in the connected
canonical `Pulse_Current.xlsx`. At 100% zoom confirm the Performance subtitle is
fully visible, the top workspace and result card are compact, the matrix and
Explain section remain readable, Mapping workflow text does not overlap, and
Reports has no `#######`. Confirm the accepted Add-ons values and
`tblPerformanceInteractionQA` 16/16 PASS. A second run must succeed without
changing content or behavior.

## Accepted live verification — 2026-08-13

The connected canonical `Pulse_Current.xlsx` was reviewed at 100% Excel zoom
after the bounded repair. The Performance subtitle rendered in full on one
line; controls, selected result, scope check, matrix, selection tables, and the
four-row Explain block were compact and readable. Mapping workflow guidance
and ready status rendered without overlap, and `EffectiveFrom` displayed as
`12.08.2026`. Reports rendered without `#######`.

All 16 `QA-0302C` checks remained PASS. The accepted Add-ons result remained
`1.47%` current share, `1.09%` comparison share, `+0.38 pp`, and `131,487 NOK`.
The workbook retained 47 sheets and only Overview, Performance, Reports,
Imports, Mapping, and Settings visible. The complete automated suite passed
112/112 and `git diff --check` passed.
