# Pulse

Pulse is an operational performance and reporting platform designed to turn trusted business data into clear, neutral performance views and meeting-ready reporting.

## Current status

**Current checkpoint:** Build 0.1.1  
**Next milestone:** Build 0.2.0 — first deterministic KPI vertical slice (Category Sales Share)

The Excel workbook is currently the user-facing product. The repository is the source of truth for product rules, metric logic, tests, documentation, and reproducible build assets.

## Product principles

- Trust before insight.
- Pulse observes; humans decide.
- Business context and conclusions remain human-owned.
- Do not hide or remove capabilities without real-world evidence.
- Complexity may live underneath the product, but normal use should remain simple.
- KPI calculations must be deterministic, explainable, and reproducible.
- The base product must remain source-system and restaurant neutral.
- A domain only affects the final performance impression when it is enabled/presented for that product or reporting period.

## Repository structure

- `excel/` — current workbook checkpoints.
- `src/` — implementation logic, organized by responsibility.
- `office-scripts/` — Excel automation scripts when needed.
- `tests/` — deterministic validation and regression tests.
- `docs/` — architecture, product principles, data model, and KPI definitions.
- `releases/` — release notes and packaged release artifacts when appropriate.

## Current user journey

`Overview → Performance → Reports → Imports → Settings`

Administration and engine sheets remain available behind the primary user-facing workflow during development.

## Development rule

A build number is only advanced when there is a concrete, validated artifact or reproducible implementation change. Discussion alone is not a build.
