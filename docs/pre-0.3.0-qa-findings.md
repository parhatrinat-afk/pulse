# Pre-0.3.0 QA findings

Reviewed workbook: `Pulse_Build_0_1_1 (2) (1).xlsx`

## Confirmed
- No merged cells on Overview, Performance, Reports, Imports, Settings, Context or KPI Registry.
- Performance current/comparison dataset dropdown exists.
- Existing workbook XML contains no cached formula errors on the reviewed primary sheets.
- Current corrected model does not force an In-house/Takeaway selector in Performance.

## Problems found
1. Performance columns are physically far too narrow in the generated workbook.
   - Example XML widths were around 2–4.6 Excel width units.
2. Reports has the same width problem.
3. The Performance category dropdown is stored as one enormous literal list.
   - This is fragile and difficult to maintain.
   - It should reference `_Metric_Calc` instead.
4. Performance currently extends to roughly row 182 because it exposes raw source categories and restaurant detail.
   - This is not being removed in the QA patch.
   - Build 0.3.0 Mapping + Reporting Groups is intended to solve the underlying product problem.
5. All engine/development sheets are still visible.
   - They are intentionally retained for now rather than hidden/removed before real usage tells us what is needed.
6. `Effective Categories` and `_Metric_Calc` were appended late in the tab order.
   - Not harmful, but worth reorganizing during the Mapping/Reporting Groups build rather than doing another cosmetic-only reorder now.

## QA patch
`Pulse_Pre_0_3_0_QA_FIX.ts` fixes:
- Performance column widths and row heights
- Reports column widths and row heights
- range-backed Category dropdown
- range-backed Current/Compare dataset dropdowns
- selector visual cues
- user-facing page wrapping/freeze-pane consistency

It deliberately does not change facts, mappings, calculations, category membership, or Build 0.3.0 functionality.
