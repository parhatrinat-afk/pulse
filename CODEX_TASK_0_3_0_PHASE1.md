# Codex Task — Pulse 0.3.0 Phase 1: Mapping Foundation

Work in the `parhatrinat-afk/pulse` repository.

Read `AGENTS.md` and `docs/BUILD_0_3_0_SPEC.md` first, then inspect the current 0.2.0 Office Scripts, QA fix, docs, tests, and validated workbook checkpoint under `excel/`.

## Objective

Implement **Phase 1 only** of Build 0.3.0: the business-owned Reporting Groups model, hierarchy-aware Lovable-style mapping foundation, and deterministic inheritance/resolution engine.

Do not redesign Performance yet except where a tiny compatibility change is unavoidable. Do not implement future KPIs.

## Required implementation

1. Add a repository-owned Build 0.3.0 Phase 1 Office Script or migration script that safely runs against the validated `0.2.0-QA` workbook checkpoint.
2. Create an authoritative `Reporting Groups` table with stable IDs, Active status, SortOrder, description/notes, and these seed groups:
   - Add-ons
   - Non-Alcohol
   - Spirits/Cocktails
   - Coffee & Tea
   - Beer & Cider
   - Desserts
   - Wine & Sake
   - Starters
   - Mains
3. Preserve `Reporting Categories` for compatibility; mark it clearly as legacy/source-default structure. Do not delete it.
4. Implement the Mapping foundation according to the previously agreed Lovable-style workflow: users browse the real source hierarchy, inspect descendants/products and useful sales context, map at the highest safe level, inherit downward, and add more-specific exceptions only where needed. Do not reduce the design to a flat manual rule table.
5. Evolve mapping rules to support the currently observed explicit human-authored scopes:
   - SourceMainCategory
   - SourceSubCategory
   - Product
   Treat these as the current source hierarchy implementation, not a permanent architectural assumption that Pulse can only support three levels. Prefer a generic scope/node representation where practical.
6. Mapping targets must reference `ReportingGroupID`.
7. Implement the general deterministic resolver rule: **most-specific applicable explicit mapping wins; otherwise inherit the nearest mapped ancestor; otherwise Unmapped**. For the current hierarchy that is `Product > SourceSubCategory > SourceMainCategory > Unmapped`.
8. Parent/high-level mapping must never overwrite or delete lower-level explicit mappings. A product/subcategory exception must survive later parent remapping.
9. Detect same-node/same-scope overlapping active rule conflicts instead of resolving them by row order.
10. Create/evolve a read-only Effective Mapping output with enough fields to trace every product from source hierarchy through applicable explicit/inherited rules to the final Reporting Group and resolution status. Make explicit versus inherited state distinguishable.
11. Create/evolve an Excel-first hierarchy-aware Mapping administration surface. It does not need a literal web tree widget, but it must let a user efficiently browse/filter parent nodes and descendants, inspect affected products, see sales NOK/quantity context where available, and map a parent or lower-level exception without hunting through ~130 unrelated POS rows.
12. Preserve raw imports and `_Sales_Facts`; do not rewrite fact amounts/quantities for mapping.
13. Add deterministic QA fixtures/checks for inheritance, override precedence, persistence of exceptions after parent remapping, unmapped behavior, conflicts, inactive targets, and reconciliation invariants.
14. Update architecture/data-model/changelog/roadmap documentation necessary to describe the Phase 1 implementation without falsely claiming that Performance has already migrated.

## Mapping UX acceptance condition

The Phase 1 implementation is not acceptable if the only practical workflow is “open Remap Rules and manually enter scope/key rows.” The persisted rules may remain table-based underneath, but the workbook must establish the Lovable-inspired hierarchy workflow:

`browse source branch -> inspect descendants/impact -> map broad node when safe -> inherit -> override exceptions`

A broad mapping should visibly affect inherited descendants while preserving explicit lower-level exceptions.

## Important migration rule

The existing ~130 POS-derived `Reporting Categories` must **not** simply be renamed into Reporting Groups. They are source-default classifications from the 0.2.0 bridge architecture. The new Reporting Groups are a separate business semantic layer.

## Reconciliation invariants

Mapping changes may redistribute facts across Reporting Groups but must never change source totals.

For any validated scope:

- mapped sales + unmapped sales = total sales;
- mapped quantity + unmapped quantity = total quantity.

## Definition of done for this task

Return a reviewable diff/PR with:

- implementation/migration script(s);
- tests/QA checklist or deterministic fixtures;
- documentation updates;
- a concise migration note explaining what happens to legacy Reporting Categories and existing Remap Rules;
- evidence that the script/checks do not mutate raw/fact sales values.

Do not advance the repository status to a completed Build 0.3.0 release. This task is Phase 1 foundation only.
