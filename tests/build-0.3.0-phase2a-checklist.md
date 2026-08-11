# Build 0.3.0 Phase 2A validation checklist

Run after applying Phase 1 and then Phase 2A to a copy of the validated
`excel/Pulse_Build_0_2_0_QA.xlsx.xlsx` checkpoint.

## Prerequisite and stale-state handling

- [ ] Phase 1 runs successfully immediately before Phase 2A.
- [ ] Phase 2A rejects an Effective Mapping whose `AsOfDate` is not today.
- [ ] Editing a Mapping Rule without refreshing Phase 1 causes Phase 2A to fail
  with `PUL-0302A-001` before Phase 2A outputs are replaced.
- [ ] Refreshing Phase 1 and rerunning Phase 2A succeeds.
- [ ] Reordering semantically identical Mapping Rules, Reporting Groups,
  Products, or Effective Mapping rows does not change the mapping fingerprint.
- [ ] A semantic mapping change produces a different fingerprint.

## Structural

- [ ] `Metric Contract` / `tblMetricContract` exists.
- [ ] `Metric Equivalence` / `tblLegacyRPGEquivalence` exists and preserves
  human-entered rows on rerun.
- [ ] `_Metric_RPG_Facts` / `tblMetricRPGFacts` contains one row per source fact.
- [ ] `Metric Migration QA` contains all five Phase 2A QA tables.
- [ ] Every bridge row carries `MappingAsOfDate`, `MappingFingerprint`, and
  `MetricRefreshAt`.
- [ ] Every bridge row is Mapped, Unmapped, Conflict, or Inactive Target.

## Reconciliation

- [ ] Bridge fact count equals `tblSalesFacts` fact count exactly.
- [ ] Bridge Sales NOK equals source Sales NOK.
- [ ] Bridge Quantity equals source Quantity.
- [ ] The four mapping-state fact counts reconcile to source fact count.
- [ ] The four mapping-state Sales NOK values reconcile to source Sales NOK.
- [ ] The four mapping-state Quantity values reconcile to source Quantity.
- [ ] Reconciliation passes for All facts and each Active Finalized import at
  All channels.
- [ ] Reconciliation passes for each actual channel without adding mandatory
  Performance channel UI.
- [ ] Reporting Group totals independently match selected bridge rows.

## Legacy comparison

- [ ] An empty equivalence table produces a visible warning and no inferred
  definitions.
- [ ] An explicit Equivalent definition with matching membership reports PASS.
- [ ] A material difference reports VARIANCE with Sales NOK, Quantity, fact
  count, and share variance visible.
- [ ] Partial and Not Comparable definitions report INFO rather than false PASS.
- [ ] Unknown RPG/CAT IDs, mixed targets, inconsistent statuses, and duplicate
  CAT members are surfaced as configuration errors.
- [ ] `tblLegacyRPGCrosswalk` explains membership assigned to Unmapped,
  Conflict, and Inactive Target states.

## Regression and non-mutation

- [ ] Raw source sheets are unchanged.
- [ ] `tblSalesFacts` values, row count, Sales NOK, Quantity, and fingerprint are
  unchanged.
- [ ] `_Metric_Calc` is unchanged.
- [ ] Performance formulas, values, validations, and layout are unchanged.
- [ ] Reports formulas, values, and layout are unchanged.
- [ ] KPI Registry and KPI-0001 remain unchanged.
- [ ] Phase 1 Mapping Rules, Effective Mapping, Reporting Groups, Mapping, and
  Mapping QA remain intact.
- [ ] Attach Rate is absent.

## Automated and runtime validation

- [ ] `node --test tests/*.test.mjs` passes.
- [ ] `git diff --check` passes.
- [ ] Phase 2A Office Script compiles and executes in Excel for the web.
- [ ] The 18,086-row bridge completes without a `Worksheet.getRange`
  `GeneralException`; if indexed range acquisition fails, the message identifies
  the sheet, exact A1 equivalent, and output operation.
- [ ] Excel does not report protected-surface workbook reads from inside the
  fingerprint iteration.
- [ ] After an intentionally interrupted/failed run, rerunning Phase 2A safely
  replaces generated contract/bridge/QA output while preserving human-authored
  equivalence rows.
- [ ] Runtime is acceptable against the 18,086-row checkpoint.
- [ ] New sheets are readable and tables have usable widths/filters.
- [ ] A saved Phase 2A QA checkpoint is reviewed before Phase 2B begins.

## Recorded live checkpoint — 2026-08-11

Workbook: `Pulse_Build_0_3_0_Phase2A_QA.xlsx`

- [x] Phase 2A compiled and completed in Excel for the web after indexed-range
  runtime hardening.
- [x] `tblMetricRPGFacts` contained exactly 18,086 data rows at
  `_Metric_RPG_Facts!A4:T18090`.
- [x] Fact count, Sales NOK 426,611,113.82, and Quantity 2,069,940.12
  reconciled between source facts and the bridge.
- [x] Mapped, Unmapped, Conflict, and Inactive Target state coverage reconciled
  for all generated scopes.
- [x] All reconciliation rows reported PASS.
- [x] Source facts remained protected with fingerprint
  `DATA-19428a5949d194ed`.
- [x] `_Metric_Calc`, Performance, Reports, and KPI Registry protection
  reported PASS.
- [x] The temporary `EQ-QA-0001` row was removed and the equivalence table was
  verified empty after the cleanup rerun.
- [x] QA-0302A-08 returned WARN with zero inferred equivalences, as expected
  until real ID-based definitions are authored.
- [x] No Phase 2B metric or presentation cutover was performed.
