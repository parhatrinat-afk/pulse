# reporting

Build 0.3.0 Phase 2B preserves Reports as a presentation consumer of
Performance and centralizes the fixed-import KPI-0001 regression layer in
`tblMetricRPGResults`.

Phase 2C adds deterministic interaction primitives mirrored by workbook
formulas. Restaurant and Reporting Group selection use stable IDs and direct
Yes/No Include state; All versus Custom is derived, not user-selected.
Arbitrary restaurant scopes sum Phase 2B Restaurant-row numerator/denominator components;
share Grand Total divides those sums and never averages restaurant shares. NOK
Impact Grand Total subtracts aggregated comparison share × aggregated current
denominator from the aggregated current numerator; it never sums restaurant
impacts. Reports continues to link to the same Performance detail result and
interactive scope.
The workbook keeps numeric component and selected-display helpers authoritative;
the visible matrix is an isolated FIXED-based text facade that is never an
input to metric arithmetic or Reports.

`matrixTotalDisplayValue` aggregates only selected RPG cells. Denominators must
agree across RPGs and are used once; Current Sales NOK returns the selected
current numerator without a denominator dependency. `sortInteractivePerformanceRows`
sorts Restaurant rows by the full-precision Total or RPG numeric value, retains
unavailable values last, applies deterministic RestaurantID ties, and leaves
Grand Total outside the ranking. These pure helpers mirror the workbook formula
contract without changing canonical row identity.

`weekly-performance.mjs` adds the bounded weekly Current/Compare contract. It
requires the single fresh Active compact cache, rejects incomplete or invalid
ISO-week ranges, and aggregates weekly RPG numerators and scope denominators
before the unchanged Phase 2C metric math. Same ranges are allowed; different
complete lengths are allowed with a factual warning. The workbook cutover uses
this cache for the four additive component matrices while preserving the
existing Phase 2C helpers and Reports linkage; `tblMetricRPGResults` remains
rollback-only.
