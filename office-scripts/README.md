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
- `Build_0_3_0_Phase2C.ts` — validates the accepted Phase 2B result grain and
  Company controls, then installs stable-ID Yes/No restaurant and Reporting
  Group selection, additive component helpers, a bounded numeric display helper,
  selected-RPG Total helpers, numeric presentation sorting, and a formula-driven
  Performance text facade. Normal selection, sorting, display, and dataset
  exploration requires recalculation only; the script does not create another
  metric engine or restaurant combinations.
- `Build_0_3_0_UX_IA_Slice1.ts` — runs after an accepted 16/16 Phase 2C
  checkpoint, exposes and orders the six primary workflow sheets, normally hides
  all 41 supporting sheets, installs five ordinary Overview hyperlinks, and
  resets the six primary saved views to A1. It does not change calculations,
  tables, formats, mappings, imports, or facts.
- `Build_0_3_0_UX_Visual_Slice2A.ts` — runs after accepted Phase 2C + IA Slice
  1, applies the restrained Pulse visual system to the six primary sheets,
  repairs Reports clipping, and replaces the technical Performance Explain
  block with four operational rows. Formula and protected table fingerprints
  must remain unchanged before/after the presentation-only pass.
- `Build_0_3_0_UX_Visual_Slice2B.ts` — runs after accepted Visual Slice 2A and
  performs the final pre-checkpoint Performance/Mapping cleanup for comfortable
  100% zoom. Values, formulas, tables, validation, selections, QA, and sheet
  state are fingerprinted or revalidated before/after the bounded formatting.
- `Migrate_Lovable_Mapping.ts` — idempotently applies the approved 129-rule
  Lovable business-definition migration using stable Pulse node/product and
  Reporting Group IDs. It reuses semantically identical active rules and adds
  the eight approved Product exclusions without changing facts or metric
  presentation.
- `Parse_Weekly_Sales_Report.ts` — read-only parser adapter for one untouched
  weekly `Sales per Item` export. It derives period identity from `A1`, validates
  the exact seven-column schema, returns normalized source rows and a
  filename-independent manifest, and performs no staging, publication, mapping,
  cache, or workbook mutation.
- `Materialize_Weekly_Compact_Cache.ts` — validates the accepted date-neutral
  weekly mapping/catalog state, writes the deterministic 84-week compact cache
  through hidden staging, reconciles/fingerprints the complete candidate, and
  leaves it `Candidate` / `Not Active`. It does not cut Performance over.
- `Activate_Weekly_Compact_Cache.ts` — revalidates the exact materialized cache,
  current mapping/catalog/ReportingEnabled content, reconciliation and Phase 2C
  QA before changing only the accepted version's two authority fields to
  `Active`. It does not add period selectors or cut Performance over.
- `Build_0_3_0_Weekly_Performance.ts` — validates the exact fresh Active cache,
  installs independent Year/From week/To week controls, and sizes the accepted
  Phase 2C component/presentation layers from the active Reporting Group count.
  Stable-ID selections are preserved and newly active groups default `No`.
  Invalid or incomplete ranges remain blocked.
- `Build_0_3_0_Performance_Presentation_Cleanup.ts` — runs after the accepted
  weekly Performance cutover, moves the two existing native selection tables
  below Explain, balances matrix widths/alignment, and updates the obsolete
  dataset wording. It fingerprints formulas, selections, validations, Reports,
  rollback results, Imports, and both 16/16 QA surfaces before/after.

Build 0.3.0 scripts should preserve the existing workbook unless an explicit migration is required. They must not rewrite raw imported source data as part of mapping.

After Phase 2C is live-validated, apply UX IA Slice 1 as a separate final step.
Its exact 47-sheet and 16/16 QA preflight prevents navigation state from being
applied to an incompatible checkpoint.

After IA Slice 1 is accepted, apply Visual Slice 2A. It validates the IA state
and 16/16 Phase 2C QA before formatting; no Office Script rerun is needed for
normal Performance exploration after the visual pass.

Where practical, scripts should be:

- deterministic and repeatable;
- safe to rerun or explicit about non-idempotent behavior;
- scoped to the smallest necessary workbook changes;
- documented with prerequisites and expected postconditions;
- validated against the current checkpoint before release.
