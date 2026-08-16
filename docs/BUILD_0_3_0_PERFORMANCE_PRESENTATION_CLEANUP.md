# Build 0.3.0 — Performance Presentation Cleanup

## Scope

This bounded slice follows the accepted flexible weekly Performance cutover.
It changes presentation only: the two existing native selection tables move
below the compact Explain section, the existing matrix receives balanced
column widths and alignment, and one obsolete dataset reference is replaced by
weekly-period wording.

No table is rebuilt. Table names, stable IDs, Include values, Yes/No validation,
structured references, formulas, weekly-cache authority, Current/Compare logic,
NOK Impact, Total/Grand Total, sorting, Reports linkage and rollback structures
remain unchanged.

## Layout

- `tblPerformanceRestaurantSelection`: `Performance!B51:D67`
- `tblPerformanceRPGSelection`: `Performance!F51:H60`
- Restaurant selection and Reporting Group selection titles: row 50
- Matrix: unchanged at `Performance!A23:K40`
- Explain: unchanged at rows 43–47 except for the approved Definition wording

The tables are relocated together through one native range move from
`N4:T21` to `B50:H67`. This retains the existing Excel Table objects rather
than recreating the selection mechanism.

## Presentation contract

- Restaurant matrix labels remain left-aligned.
- Matrix category headers, numeric facade values and Grand Total values are
  centered.
- Category columns use a bounded 95–110 width range; Beer & Cider is reduced
  from 180 to 105 and its header remains wrapped/readable.
- The top Interactive matrix status text wraps within the narrower Beer & Cider
  column without changing its value.
- Current and Compare Year/From week/To week labels use the same neutral label
  treatment.
- Selection headers wrap within a readable 32-point header row, and selection
  names wrap within readable 28-point body rows below Explain.

The Explain definition is:

> Selected Reporting Group sales as a share of total sales for the selected restaurants and period.

## Guards

The script requires the accepted 48-sheet, six-visible-sheet workbook and both
16/16 QA tables before mutation. It fingerprints the matrix, top controls,
Performance formulas, Current/Compare selections, Reports, rollback metric
results, Imports, both selection-table contents and both validation rules.
Postconditions require the new native table addresses and an empty former
far-right selection area.

Normal exploration remains recalculation-only. No Office Script rerun is needed
for period, restaurant, Reporting Group, display or sort changes.
