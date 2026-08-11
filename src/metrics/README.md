# metrics

Deterministic metric implementation assets.

- `category-sales-share.md` documents the validated Build 0.2.0 legacy CAT
  calculation path.
- `reporting-group-metrics.mjs` implements the Build 0.3.0 Phase 2A derived
  Reporting Group metric contract, mapping fingerprint, fact bridge,
  reconciliation, group aggregation, and human-configured legacy comparison.

Phase 2A applies the current mapping state to historical facts for analysis. It
does not mutate facts and does not cut `_Metric_Calc`, Performance, Reports, or
KPI-0001 over to Reporting Groups. That cutover belongs to Phase 2B.
