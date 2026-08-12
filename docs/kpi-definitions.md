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

The enabled-restaurant definition is the shared default Performance scope.
KPI-0001 consumes that contract; it does not own or redefine it. Future KPIs
must use the same scope unless their approved definition explicitly documents a
different organizational-scope requirement.

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

Build 0.3.0 Phase 2B implements:

- selected-Reporting-Group result
- company Reporting Group overview
- selected-Reporting-Group restaurant breakdown
- comparison-scope check
- Explain section
- Reports consuming the same selected result

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

## Planned KPIs

- Reporting Group Quantity Share
- Add-ons Attach Rate
- Percentage-point Change
- Estimated NOK Impact

They remain Draft until their calculation paths are implemented and validated.
