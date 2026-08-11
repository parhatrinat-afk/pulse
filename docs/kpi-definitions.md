# KPI Definitions

## KPI-0001 — Category Sales Share

**Domain:** Commercial Performance  
**Status:** First implementation target  
**Default scope:** In-house sales

### Definition

Category Sales Share measures the selected reporting category's sales amount as a proportion of total sales amount within the same valid scope and period.

`Category Sales Share = Category Sales Amount / Total Sales Amount`

### Required dimensions

- Period.
- Restaurant or aggregate organizational scope.
- Reporting category.
- Channel/scope filter, initially in-house.

### Initial comparison

- Current: Week 31, 2026.
- Baseline: 2025 reference data available in the current workbook.

The exact baseline interpretation used in a production comparison must be explicit in the metric result/report rather than inferred by the user.

### Presentation

Initial views:
- Company/category overview.
- Restaurant view for a selected category.

Future views may include trend and leaderboard once the first vertical slice is validated.

## Planned KPIs

- Category Quantity Share.
- Add-ons Attach Rate.
- Percentage-point Change.
- Estimated NOK Impact.

These remain planned until the first metric engine path is proven.
