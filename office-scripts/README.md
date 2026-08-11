# office-scripts

This directory contains Office Scripts used to build, repair, validate, or administer the Pulse Excel workbook.

Current scripts include:

- `Build_0_2_0.ts` — reproducible Build 0.2.0 implementation path.
- `Pre_0_3_0_QA_FIX.ts` — pre-0.3.0 QA corrections applied to the validated checkpoint.
- `Build_0_3_0_Phase1.ts` — rerunnable Reporting Groups, hierarchy mapping,
  Effective Mapping, browse/action, and QA foundation. It intentionally leaves
  Performance and Reports on the 0.2.0 metric path.

Build 0.3.0 scripts should preserve the existing workbook unless an explicit migration is required. They must not rewrite raw imported source data as part of mapping.

Where practical, scripts should be:

- deterministic and repeatable;
- safe to rerun or explicit about non-idempotent behavior;
- scoped to the smallest necessary workbook changes;
- documented with prerequisites and expected postconditions;
- validated against the current checkpoint before release.
