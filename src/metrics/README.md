# metrics

Deterministic metric implementation assets.

- `category-sales-share.md` documents the validated Build 0.2.0 legacy CAT
  calculation path.
- `reporting-group-metrics.mjs` implements the Build 0.3.0 Phase 2A derived
  Reporting Group metric contract, mapping fingerprint, fact bridge,
  reconciliation, group aggregation, and human-configured legacy comparison.
- `reporting-group-sales-share.mjs` implements the Phase 2B bridge preflight,
  deterministic centralized KPI-0001 result grain, enabled-restaurant Company
  scope fingerprint, all-state denominator, and company/restaurant calculation
  validation.
- `reporting-group-sales-share.md` documents the active Phase 2B metric and
  lineage contract.

Build 0.3.0 applies the current mapping state to historical facts for analysis.
Phase 2B consumes the Phase 2A bridge without mutating facts or independently
resolving mapping.
