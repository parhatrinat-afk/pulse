# Category Sales Share implementation — legacy Build 0.2.0

Historical implementation target: `KPI-0001`

This document preserves the validated Build 0.2.0 CAT-based calculation for
compatibility and regression analysis. Build 0.3.0 Phase 2B retains the stable
KPI-0001 ID but activates Reporting Group Sales Share through
`tblMetricRPGResults`; this legacy formula is no longer authoritative for
Performance.

The authoritative facts remain `tblSalesFacts`.

The build script creates UI formulas rather than a second permanent source of sales truth.

## Formula

For a selected scope:

`Category Sales Share = SUM(Category SalesAmount) / SUM(Total SalesAmount)`

Both numerator and denominator use exactly the same:

- ImportID
- ReportingChannel when a specific channel is selected
- RestaurantID when restaurant scope is applied
- PublicationState = Active Finalized

Only the numerator additionally filters `ReportingCategoryID`.

`All channels` intentionally removes the ReportingChannel filter from both numerator and denominator.

## Traceability

The UI is derived from finalized facts and stable category/restaurant identities.

The script does not alter the fact store.
