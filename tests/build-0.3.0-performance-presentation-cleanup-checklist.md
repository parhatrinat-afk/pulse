# Performance presentation cleanup — live Excel checklist

- [ ] Run `Build_0_3_0_Performance_Presentation_Cleanup.ts` once in canonical
      `Pulse_Current.xlsx`.
- [ ] At 100% zoom, verify normal Performance use does not require horizontal
      scrolling to reach Restaurant or Reporting Group selection.
- [ ] Verify `tblPerformanceRestaurantSelection` is at `B51:D67` and
      `tblPerformanceRPGSelection` is at `F51:H60`, side-by-side below Explain.
- [ ] Change one Restaurant Include value Yes → No → Yes and confirm the matrix,
      Total and Grand Total recalculate immediately without rerunning a script.
- [ ] Change one Reporting Group Include value Yes → No → Yes and confirm the
      matrix and Total recalculate immediately without rerunning a script.
- [ ] Verify Restaurant names are left-aligned; matrix headers, numeric values
      and Grand Total values are centered.
- [ ] Verify Beer & Cider is balanced with the other category columns and its
      heading is fully readable.
- [ ] Verify all five Display modes, Highest/Lowest sorting, Total, Grand Total,
      negative-red styling and zero/unavailable display remain unchanged.
- [ ] Verify Current and Compare Year/From week/To week labels are visually
      consistent and all six validation lists still work.
- [ ] Verify Explain says “selected restaurants and period”.
- [ ] Verify Reports follows Detail Reporting Group and Current/Compare labels.
- [ ] Verify `tblWeeklyPerformanceQA` remains 16/16 PASS.
- [ ] Verify `tblPerformanceInteractionQA` remains 16/16 PASS.
- [ ] Verify no clipping or overlap at 100% zoom.
