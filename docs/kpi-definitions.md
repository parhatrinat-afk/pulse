# KPI Definitions

## KPI-0001 — Reporting Group Sales Share

**Domain:** Commercial Performance  
**Status:** Build 0.3.0 Phase 2B implementation

### Definition

Reporting Group Sales Share measures Mapped sales for the selected stable
ReportingGroupID as a proportion of all sales **inside the selected scope**.

`Reporting Group Sales Share = selected Mapped RPG Sales NOK / all Sales NOK in the identical scope`

Unmapped, Conflict, Inactive Target, and explicit Product-exclusion facts remain
in the denominator when they belong to the selected organizational scope. For
Phase 2B Company results, that scope is the stable RestaurantID set where
`Status=Active` and `ReportingEnabled=Yes`. Facts from deselected restaurants
enter neither numerator nor denominator, remain unchanged in Phase 2A, and are
reconciled separately. Only `ResolutionStatus=Mapped` facts for the selected RPG
enter the numerator.

The enabled-restaurant definition is the shared default Performance eligibility
scope. Phase 2C may interactively select a subset of those stable RestaurantIDs;
unselected restaurants enter neither selected numerator nor denominator. The
same subset applies to current and comparison. KPI-0001 consumes this scope
contract; it does not own or redefine it. Future KPIs must use the same scope
unless their approved definition explicitly documents a different
organizational-scope requirement.

### Important architecture rule

**Channel is a scope dimension, not part of the KPI definition.**

The same KPI can therefore be calculated for:

- In-house
- Takeaway
- All channels
- future valid channels introduced by an adapter

### Current and comparison

Current and comparison are independently selected.

A user may intentionally compare, for example:

- Week 32, 2026 Takeaway vs Week 32, 2025 Takeaway
- Week 32, 2026 In-house vs Week 31, 2026 In-house
- one period/channel against another different period/channel

Pulse may surface that the scopes differ, but it must not prohibit an intentional comparison.

### Dataset vs baseline

A dataset is simply trusted data available to Pulse.

It is **not permanently a baseline**.

A dataset or period becomes the baseline only when the user selects it as the comparison reference.

### Required dimensions

- current dataset/period
- comparison dataset/period
- Reporting Group
- channel
- restaurant/organizational scope where applicable

### Presentation

Build 0.3.0 Phase 2B implements the centralized materialized results. Phase 2C
adds a formula-driven presentation over their additive Restaurant components:

- selected-Reporting-Group result
- independent current/comparison datasets
- direct Yes/No eligible restaurant selection
- direct Yes/No active Reporting Group matrix selection
- Restaurant × Reporting Group matrix
- selected-RPG Total column immediately after Restaurant
- numeric sorting by Total or a displayed Reporting Group, Highest/Lowest
- weighted selected-scope Grand Total
- five display modes: PP Change, Current Share, Comparison Share, Current Sales NOK, NOK Impact
- comparison-scope check
- Explain section
- Reports consuming the same selected result

The single detail Reporting Group selector is independent of matrix
multi-selection. A zero denominator retains canonical value zero but displays
as `—` in Performance.

NOK Impact is a derived Performance display, not a new KPI identity. It equals
current selected-scope Reporting Group numerator Sales NOK minus comparison
Reporting Group share × current selected-scope denominator Sales NOK. Its Grand
Total uses aggregated numerators and denominators rather than summing
restaurant-level impacts. Different period lengths remain allowed because the
comparison supplies a share baseline and the current scope supplies the
monetary base.

The matrix's visible strings are presentation only. KPI-0001 components,
selected-display values, Grand Total, QA, and downstream arithmetic remain
numeric; they must not consume or parse the visible text facade.

Total means the aggregate of the currently selected Reporting Groups. It sums
selected RPG numerators but uses the scope denominator only once. Current Sales
NOK Total is the current selected numerator even when its scope denominator is
zero. Current/Comparison Share, PP Change, and NOK Impact preserve their
denominator requirements and are recomputed from the aggregate components.
Sorting consumes the full-precision numeric selected-display value and changes
only the visible RestaurantID lookup order; it never sorts component rows or
includes Grand Total.

### Decision boundary

Pulse presents the calculation and factual scope differences.

Pulse does not decide whether a comparison is commercially appropriate and does not recommend an action.

### Build 0.3.0 Phase 2B calculation path

Phase 2B materializes the central KPI result from `_Metric_RPG_Facts` into
`tblMetricRPGResults`. Performance retrieves those results by MetricID,
ImportID, ReportingGroupID, company/restaurant scope, all-channel scope, and
publication state. Reports links to the same Performance result.

Legacy ReportingCategoryID and CAT/RPG equivalence do not participate in the
active KPI-0001 calculation. KPI-0001 retains its stable ID while its
classification dimension evolves from the Build 0.2.0 CAT path to the approved
Reporting Group semantic layer.

Phase 2C does not create another KPI result table. For any selected restaurant
set, it sums Phase 2B Restaurant-scope NumeratorSalesNOK and
DenominatorSalesNOK, then divides the sums. This is mathematically equivalent to
calculating KPI-0001 over the union of those disjoint restaurant scopes and is
not an average of restaurant percentages.

## Planned KPIs

- Reporting Group Quantity Share
- Add-ons Attach Rate
- Percentage-point Change

They remain Draft until their calculation paths are implemented and validated.
