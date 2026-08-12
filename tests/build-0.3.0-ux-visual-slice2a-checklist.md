# Build 0.3.0 UX Visual Slice 2A — Live Excel QA Checklist

Baseline: accepted Phase 2C + UX IA Slice 1 workbook.

## Before running

- [ ] Confirm visible tabs are exactly:
      `Overview | Performance | Reports | Imports | Mapping | Settings`.
- [ ] Confirm `tblPerformanceInteractionQA` contains 16 rows and all are
      `PASS`.
- [ ] Record current Performance Detail Reporting Group, datasets, restaurant
      and RPG selections, Display, Sort by, Order, and representative results.

## Run

- [ ] Add or replace the Office Script with
      `office-scripts/Build_0_3_0_UX_Visual_Slice2A.ts`.
- [ ] Run it once and confirm the completion message reports presentation-only
      changes.
- [ ] Confirm Overview is active at A1.

## Shared primary-sheet visual system

- [ ] The six visible tabs and their order are unchanged.
- [ ] Gridlines are off on all six primary sheets.
- [ ] Titles/subtitles are consistent, readable, and not clipped.
- [ ] Overview navigation still opens all five correct target sheets.
- [ ] Each primary sheet opens at A1.

## Performance

- [ ] Controls/dropdowns remain in the same cells and accept changes.
- [ ] Detail Reporting Group is visibly distinguished from matrix RPG
      selection.
- [ ] Restaurant and RPG Include cells remain editable and visually clear.
- [ ] Restaurant/RPG selection summaries and comparison/scope status remain
      readable.
- [ ] Cycle all five Display modes.
- [ ] Verify Highest/Lowest sorting and hidden-RPG sort fallback.
- [ ] Verify Total and Grand Total remain correct.
- [ ] Confirm negative matrix values retain red presentation.
- [ ] Confirm the matrix structure remains
      `Restaurant | Total | selected RPG columns`.
- [ ] Confirm the Explain section contains only Metric, Definition, Total, and
      Comparison and all text is fully visible.

## Reports

- [ ] Report context labels and dataset/status values are fully visible.
- [ ] No cell displays `#######`.
- [ ] No label/value overlaps or spills into unrelated cells.
- [ ] Current share, Comparison share, Change, and Current Reporting Group
      sales are readable and aligned.
- [ ] Change retains two-decimal pp formatting and Current Sales retains whole
      NOK formatting.
- [ ] Changing Performance Detail Reporting Group updates Reports.
- [ ] Changing current/comparison datasets and restaurant scope updates Reports.
- [ ] Reports values agree exactly with the Performance detail result.

## Imports, Mapping, Settings

- [ ] Imports table rows, columns, filters, values, and freeze pane are intact.
- [ ] Mapping action dropdowns, dates, NodeID/RuleID workflow, status, and all
      three hierarchy tables remain intact.
- [ ] Mapping inputs remain editable and clearly distinguished.
- [ ] Settings IDs and values are unchanged; Value cells remain editable.

## Regression and rerun

- [ ] `tblPerformanceInteractionQA` remains exactly 16/16 PASS.
- [ ] No worksheet or table was added, deleted, renamed, or resized.
- [ ] Hidden supporting sheets remain normally hidden and accessible to admins.
- [ ] Run Slice 2A a second time.
- [ ] Confirm the second run succeeds and produces no duplicate objects,
      conditional formats, merges, or content.
- [ ] Confirm all interactions and Reports linkage still work after the rerun.

## Acceptance evidence

- Workbook/checkpoint name:
- Excel-for-web run date:
- First-run result:
- Idempotent rerun result:
- Phase 2C QA result:
- Reviewer:
- Notes:
