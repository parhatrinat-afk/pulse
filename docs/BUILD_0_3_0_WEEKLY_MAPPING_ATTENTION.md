# Build 0.3.0 — Weekly Mapping Attention Foundation

This bounded slice adds derived administrative evidence for the future Mapping
workspace. It does not add mapping authority, change mapping semantics, or
redesign the visible Mapping workflow.

## Contract

The accepted weekly identity preflight already produces stable Product
identities, applies the current Pulse Product hierarchy, and calls the existing
hierarchical mapping resolver. The attention projection consumes those accepted
outputs and aggregates the active weekly history to one row per stable
`ProductID`.

The hidden `tblWeeklyMappingAttention` projection contains:

- stable Product identity and exact source label fields;
- current authoritative Main Category / Subcategory lineage;
- effective Reporting Group, winning rule, resolution source, and mapping state;
- 85-week Fact count, Sales NOK, and Quantity;
- Identity Pending and alternate-hierarchy attention evidence; and
- no user-facing cache or fingerprint identifiers.

Resolution labels are `Explicit Product`, `Inherited Subcategory`, `Inherited
Main`, `Explicit exclusion`, `Unmapped`, `Identity Pending`, `Conflict`, and
`Inactive Target`. The projection does not resolve or recommend a mapping.

The hidden `tblWeeklyMappingAttentionControl` row ties the projection to the
single validated Active weekly cache and records the technical fingerprints
needed for reproducibility. Those fields remain engineering evidence.

## Accepted 85-week evidence

The complete `2025 W01–W52` plus `2026 W01–W33` corpus produces:

| State | Products | Facts | Sales NOK | Quantity |
| --- | ---: | ---: | ---: | ---: |
| Mapped | 929 | 229,190 | 479,649,885.10 | 2,367,590.46 |
| Unmapped | 302 | 19,256 | 10,595,207.65 | 131,052.98 |
| Identity Pending | 6 | 126 | 120,048.00 | 999.00 |
| Conflict | 0 | 0 | 0.00 | 0.00 |
| Inactive Target | 0 | 0 | 0.00 | 0.00 |
| **Total** | **1,237** | **248,572** | **490,365,140.75** | **2,499,642.44** |

The Product universe is 1,041 existing catalog identities plus 196 identities
discovered through weekly ingestion. The projection fingerprint is
`WMA-637d7a94536ac4ed`. There are no duplicate ProductIDs or ProductKeys.

The six non-blocking Identity Pending products are:

- `PRD-001102` Fuelbox Teams - English;
- `PRD-001125` Hot Food Combo;
- `PRD-001138` Korean Fried Chicken;
- `PRD-001187` Staff Coca Cola;
- `PRD-001214` Sumo Reusable Bag; and
- `PRD-001223` Tempura Poke Bowl.

The nine previously accepted alternate-hierarchy review cases remain visible;
the current Pulse Product hierarchy remains authoritative. No hierarchy or
mapping is changed by this projection.

## Workbook surface

`Materialize_Weekly_Mapping_Attention.ts` writes the validated evidence through
bounded `Prepare`, `Write`, and `Finalize` operations to hidden sheet
`_Weekly_Mapping_Attention`:

- `tblWeeklyMappingAttentionControl` — one row; and
- `tblWeeklyMappingAttention` — 1,237 Product rows.

The visible Mapping change is limited to a compact block at `E9:N14` showing
Mapped, Unmapped, Identity Pending, Conflict, and Inactive Target Product/Facts/
Sales coverage plus the plain-language freshness state:

`Performance classifications are up to date through 2026 W33`

Identity Pending and Unmapped produce `Attention required`, not a system
failure. A later Phase 1 mapping mutation marks `Performance refresh required`;
the weekly historical cache must then be rebuilt through its existing freshness
contract before classification is described as current again.

## Non-regression boundary

This slice does not change Mapping Rules, Effective Mapping, the active or
rollback weekly cache, source facts, Performance, Reports, Imports, weekly
ingestion, or the six visible-sheet architecture. It does not implement the
cascading Mapping browser or a new Map/Change/Remove form.

The local 85-week corpus remains a read-only development fixture and is never
added to Git. The canonical workbook remains
`OneDrive/Pulse/Development/Pulse_Current.xlsx` and is not committed.
