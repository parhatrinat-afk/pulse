# Build 0.3.0 UX IA Slice 1 — Live Excel QA Checklist

Baseline: `Pulse_0.3.0_Phase2C_16of16_QA.xlsx.xlsx`

## Before running the script

- [ ] Open the exact accepted Phase 2C workbook in Excel for the web.
- [ ] Unhide `Metric Results QA` if needed and confirm
      `tblPerformanceInteractionQA` contains 16 checks and all are `PASS`.
- [ ] Record a representative Performance result and selection state for the
      post-run interaction check.

## Run

- [ ] Add/replace the Office Script with
      `office-scripts/Build_0_3_0_UX_IA_Slice1.ts`.
- [ ] Run it once and confirm it reports six visible primary tabs.
- [ ] Run it a second time and confirm it completes without adding or deleting
      sheets/tables or changing the result.

## Information architecture

- [ ] The only visible tabs are, in order:
      `Overview | Performance | Reports | Imports | Mapping | Settings`.
- [ ] Overview is active and opens at A1 after the script.
- [ ] Overview links open Performance, Reports, Imports, Mapping, and Settings.
- [ ] Each navigation target opens at A1.
- [ ] Performance opens at the top, not around row 16.
- [ ] A representative hidden sheet can be unhidden and still contains its
      prior data/formulas/table; hide it again after inspection.
- [ ] Hidden sheets are ordinary hidden sheets and can be accessed by an admin.

## Phase 2C regression

- [ ] Restaurant selection still recalculates Performance without a script.
- [ ] Reporting Group selection still recalculates without a script.
- [ ] Current/comparison dataset selection still recalculates without a script.
- [ ] All five matrix display modes, Total, Grand Total, and sorting still work.
- [ ] The representative Performance result and selection state are unchanged.
- [ ] Reports still agrees with the Performance detail result.
- [ ] `tblPerformanceInteractionQA` remains exactly 16/16 PASS.
- [ ] Metric Results QA and all other accepted calculation outputs are unchanged.

## Acceptance evidence

- Workbook/checkpoint name:
- Excel-for-web run date:
- First run result:
- Idempotent rerun result:
- Visible tab order:
- Phase 2C QA result:
- Reviewer:
- Notes:
