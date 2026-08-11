# Build 0.2.0 validation checklist

Run these after applying `office-scripts/Build_0_2_0.ts`.

## Structural

- [ ] Build script completes without an Office Script error.
- [ ] `Performance` opens with Category Sales Share.
- [ ] `Reports` consumes the same selected result.
- [ ] `_Metric_Calc` exists.
- [ ] No rows in `_Sales_Facts` were edited or deleted.
- [ ] KPI-0001 is Active and visible.
- [ ] Unimplemented KPIs remain Draft.

## Selector behavior

- [ ] Current dataset can be changed independently.
- [ ] Compare-with dataset can be changed independently.
- [ ] Current channel can be changed independently.
- [ ] Compare-with channel can be changed independently.
- [ ] `All channels` is selectable.
- [ ] Available published datasets are generated from `tblImports`.
- [ ] Available channels are generated from active finalized facts.

## Comparison guardrail

- [ ] Same channel + same period length shows comparable scope.
- [ ] Different channel shows a scope difference.
- [ ] Different period length shows a scope difference.
- [ ] A scope difference does NOT prevent a result from calculating.

## Regression checks for current Build 0.1.1 data

Select category `Add-ons`.

Expected approximately:

- Week 31, 2026 / In-house: 0.1686%
- 2025 full year / In-house: 0.1169%
- Week 31, 2026 / Takeaway: 0.0877%
- 2025 full year / Takeaway: 0.1542%

Small display rounding is expected; calculation logic should reconcile to the underlying facts.

## Product principle

- [ ] No recommendation, action instruction, or personality is introduced by the build.
