# KPI Definitions

## KPI-0001 — Category Sales Share

**Domain:** Commercial Performance  
**Status:** Build 0.2.0 implementation

### Definition

Category Sales Share measures the selected reporting category's sales amount as a proportion of total sales amount **inside the selected scope**.

`Category Sales Share = Category Sales Amount / Total Sales Amount`

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
- reporting category
- channel
- restaurant/organizational scope where applicable

### Presentation

Build 0.2.0 implements:

- selected-category result
- company category overview
- selected-category restaurant breakdown
- comparison-scope check
- Explain section
- Reports consuming the same selected result

### Decision boundary

Pulse presents the calculation and factual scope differences.

Pulse does not decide whether a comparison is commercially appropriate and does not recommend an action.

## Planned KPIs

- Category Quantity Share
- Add-ons Attach Rate
- Percentage-point Change
- Estimated NOK Impact

They remain Draft until their calculation paths are implemented and validated.
