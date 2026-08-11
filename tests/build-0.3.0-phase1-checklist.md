# Build 0.3.0 Phase 1 validation checklist

Run after applying `office-scripts/Build_0_3_0_Phase1.ts` to the validated
`excel/Pulse_Build_0_2_0_QA.xlsx.xlsx` checkpoint.

## Structural

- [ ] `Reporting Groups` / `tblReportingGroups` contains nine unique stable seed IDs.
- [ ] Group display names can be edited without changing rule targets.
- [ ] Only active Reporting Group IDs appear in the Mapping action dropdown.
- [ ] `Reporting Categories` remains present and is clearly marked legacy/source-default.
- [ ] `Remap Rules` and `Effective Categories` remain present for 0.2.0 compatibility.
- [ ] `Mapping Rules`, `Mapping`, `Effective Mapping`, `Mapping QA`, and `_Mapping_Lists` exist.
- [ ] Performance, Reports, and `_Metric_Calc` retain the 0.2.0 category metric path.

## Hierarchy workflow

- [ ] Mapping exposes filterable Main Category, Subcategory, and Product browse tables.
- [ ] Parent rows show affected subcategory/product counts and sales/quantity breadth.
- [ ] Product rows show effective group, resolution source/state/status, and winning rule.
- [ ] A Mapping action appends an explicit rule without editing source hierarchy or facts.
- [ ] A Deactivate action changes only rule status and restores nearest-ancestor resolution.

## Deterministic resolver fixtures

- [ ] Main-category mapping is inherited by an otherwise-unmapped product.
- [ ] Subcategory mapping overrides a main-category mapping.
- [ ] Product mapping overrides both ancestors.
- [ ] Product override survives a later parent remap.
- [ ] Inactivating the override restores the nearest valid ancestor.
- [ ] Unmapped remains visible.
- [ ] Same-node/same-scope overlapping active rules produce Conflict.
- [ ] An inactive Reporting Group target produces Inactive Target.

## Reconciliation and regression

- [ ] `Mapping QA` reports mapped + unmapped Active Finalized sales equals fact sales.
- [ ] `Mapping QA` reports mapped + unmapped Active Finalized quantity equals fact quantity.
- [ ] Fact row count, total sales, total quantity, first ID, and last ID are unchanged.
- [ ] The 0.2.0-QA range-backed Performance selectors and practical column widths remain.
- [ ] `node --test tests/build-0.3.0-phase1.test.mjs` passes.

## Human review

- [ ] Run the Office Script in Excel for the web and save a separate Phase 1 checkpoint.
- [ ] Visually inspect Mapping, Reporting Groups, Effective Mapping, and Mapping QA.
- [ ] Exercise an Apply action on a copy, then remap its parent and confirm the exception remains.

## Accepted live QA record — 2026-08-11

- [x] Office Script executed successfully in Excel for the web.
- [x] Dropdown validation reported ready for all six selectors.
- [x] SourceMainCategory and SourceSubCategory inheritance/override behavior passed.
- [x] Product-level explicit overrides passed.
- [x] Lower-level explicit rules survived parent remapping.
- [x] Overlapping explicit rules surfaced as conflicts rather than resolving silently.
- [x] Deactivation surfaced inactive targets; reactivation restored mappings.
- [x] All nine `Mapping QA` checks passed.
- [x] Sales and quantity reconciliation passed.
- [x] Fact row count remained 18,086.
- [x] Conflict table was empty after cleanup.
- [x] Performance remained on the intentionally deferred legacy category calculation layer.
