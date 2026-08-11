# mapping

Build 0.3.0 Phase 1 introduces a generic scope/node mapping contract. The
current Katria adapter exposes `SourceMainCategory`, `SourceSubCategory`, and
`Product` hierarchy levels, but resolver semantics are expressed as “most
specific explicit mapping, otherwise nearest mapped ancestor, otherwise
Unmapped.”

`hierarchical-resolver.mjs` is the pure deterministic fixture implementation.
The workbook implementation is in `office-scripts/Build_0_3_0_Phase1.ts`.

The narrow Lovable migration extension adds Product-scoped explicit exclusion:
an active `Exclude` rule wins at Product specificity and resolves to Unmapped
with rule lineage. `lovable-migration.mjs` contains the frozen 129-decision
stable-ID contract and idempotent migration planner mirrored by
`office-scripts/Migrate_Lovable_Mapping.ts`.
