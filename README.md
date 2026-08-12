# Pulse

Pulse is an operational performance and reporting platform designed to turn trusted business data into clear, neutral performance views and meeting-ready reporting.

## Current status

**Current validated checkpoint:** Build 0.2.0-QA  
**Active development:** Build 0.3.0 — Mapping + Reporting Groups

Phase 1, Phase 2A, and the approved Lovable mapping migration have passed live
Excel-for-web QA. Phase 2B has also passed live Excel-for-web calculation,
restaurant-scope, reconciliation, and rerun QA: it validates the Phase 2A
bridge, materializes centralized Reporting Group Sales Share results in
`_Metric_Calc`, and cuts the existing Performance/Reports path over without a
visual redesign or mandatory channel selector. Build 0.2.0-QA remains the
validated release checkpoint until Build 0.3.0 release acceptance.

The Excel workbook is currently the user-facing product. The repository is the source of truth for product rules, metric logic, tests, documentation, Office Scripts, and reproducible build assets.

Build 0.3.0 introduces the semantic layer between source-system classifications and reporting:

`Raw POS data → Source hierarchy → Hierarchical Mapping → Reporting Groups → Metrics / KPIs → Performance`

Reporting Groups are business-defined groupings such as Add-ons, Non-Alcohol, Spirits/Cocktails, Coffee & Tea, Beer & Cider, Desserts, Wine & Sake, Starters, and Mains. They are not KPIs. Metrics operate on Reporting Groups.

## Mapping principle

Pulse should make it easy to browse the source hierarchy and map at the highest safe level.

- A broader source node may provide an inherited Reporting Group to its descendants.
- A more-specific explicit mapping overrides an inherited mapping.
- The current POS hierarchy may expose levels such as Main Category → Subcategory → Product, but the architecture should not assume those are the only hierarchy levels Pulse will ever support.
- Raw source data is never rewritten by mapping.
- Human control over mapping and Reporting Group configuration is preserved.
- Active/inactive Reporting Groups control what becomes available to Performance.

## Product principles

- Trust before insight.
- Pulse observes; humans decide.
- Business context and conclusions remain human-owned.
- Do not hide or remove capabilities without real-world evidence.
- Complexity may live underneath the product, but normal use should remain simple.
- KPI calculations must be deterministic, explainable, and reproducible.
- The base product must remain source-system and restaurant neutral.
- Uploaded datasets are not inherently baselines; current and comparison references are user-selected.
- Dataset channel/scope such as In-house or Takeaway is metadata/context, not a mandatory KPI definition.
- Do not destroy or overwrite raw source data when mapping.

## Repository structure

- `excel/` — validated workbook checkpoints and workbook artifacts.
- `src/` — implementation logic, organized by responsibility.
- `office-scripts/` — Excel automation and build/QA scripts.
- `tests/` — deterministic validation, expected outputs, and regression checks.
- `docs/` — architecture, product principles, data model, KPI definitions, QA findings, and active build specifications.
- `releases/` — release notes and validated release/checkpoint material.
- `AGENTS.md` — repository-wide engineering rules for coding agents, including Codex.

## Current user journey

`Overview → Performance → Reports → Imports → Settings`

Administration and engine sheets remain available behind the primary user-facing workflow during development.

## Development rule

A build number is only advanced when there is a concrete, validated artifact or reproducible implementation change. Discussion alone is not a build.

Build 0.3.0 is **in development**, not released. Build 0.2.0-QA remains the validated checkpoint until 0.3.0 has a validated artifact and QA evidence.
