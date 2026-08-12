# office-scripts

This directory contains Office Scripts used to build, repair, validate, or administer the Pulse Excel workbook.

Current scripts include:

- `Build_0_2_0.ts` — reproducible Build 0.2.0 implementation path.
- `Pre_0_3_0_QA_FIX.ts` — pre-0.3.0 QA corrections applied to the validated checkpoint.
- `Build_0_3_0_Phase1.ts` — rerunnable Reporting Groups, hierarchy mapping,
  Effective Mapping, browse/action, and QA foundation. It intentionally leaves
  Performance and Reports on the 0.2.0 metric path.
- `Build_0_3_0_Phase2A.ts` — validates current Effective Mapping, materializes
  one Reporting Group analysis row per immutable sales fact, fingerprints the
  mapping state, adds reconciliation/coverage QA, and supports human-authored
  legacy CAT/RPG comparisons. It leaves `_Metric_Calc`, Performance, Reports,
  KPI Registry, and KPI-0001 unchanged.
- `Build_0_3_0_Phase2B.ts` — validates Phase 2A bridge freshness and one-to-one
  source reconciliation before mutation, materializes centralized KPI-0001
  Reporting Group results in `_Metric_Calc`, and cuts the existing Performance
  and Reports presentation over without restoring mandatory channel UI or
  redesigning the page. Active, ReportingEnabled RestaurantIDs define the
  auditable Company scope; excluded restaurants remain untouched in source and
  Phase 2A.
- `Migrate_Lovable_Mapping.ts` — idempotently applies the approved 129-rule
  Lovable business-definition migration using stable Pulse node/product and
  Reporting Group IDs. It reuses semantically identical active rules and adds
  the eight approved Product exclusions without changing facts or metric
  presentation.

Build 0.3.0 scripts should preserve the existing workbook unless an explicit migration is required. They must not rewrite raw imported source data as part of mapping.

Where practical, scripts should be:

- deterministic and repeatable;
- safe to rerun or explicit about non-idempotent behavior;
- scoped to the smallest necessary workbook changes;
- documented with prerequisites and expected postconditions;
- validated against the current checkpoint before release.
