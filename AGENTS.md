# Pulse — Agent Engineering Instructions

This file defines repository-level rules for AI coding agents working on Pulse.

## Product purpose

Pulse is an operational performance and reporting platform. It turns trusted uploaded or API-provided business data into deterministic, explainable performance views and meeting-ready reporting.

Pulse presents facts. Humans own business interpretation and decisions.

## Source of truth order

When requirements conflict, use this precedence:

1. Explicit task / build specification for the active build.
2. Validated release/checkpoint workbook for the active starting point.
3. `docs/architecture.md`, `docs/pulse-principles.md`, `docs/data-model.md`, `docs/comparison-model.md`, and KPI/metric definitions.
4. Build-specific QA findings and release notes.
5. Existing implementation code.
6. README/status text when it conflicts with newer validated artifacts.

Do not silently resolve a real architectural conflict. Describe it briefly in the PR/task result and choose the least destructive path consistent with the higher-priority source.

## Non-negotiable architecture

The target commercial-performance flow is:

`Raw source data -> Adapter/Staging -> Validation -> Published Facts -> Mapping -> Reporting Groups -> Metrics/KPIs -> Performance/Reports`

### Raw data and lineage

- Never rewrite, delete, or semantically mutate raw imported source data to achieve a mapping result.
- Published facts remain traceable to the source import and source entity.
- Mapping is a semantic layer over trusted facts/products/source classifications.
- A correction to mapping changes classification, not historical source facts.
- Import publication/finalization must remain recoverable; do not introduce irreversible data traps.

### Dataset and scope

- Pulse digests the reports/datasets supplied to it. It is not restricted to In-house.
- Channel (In-house, Takeaway, etc.) is a scope/property dimension, not a KPI definition.
- A dataset is not inherently a baseline.
- Current and comparison scopes are independently selected.
- Unusual comparisons may generate factual warnings but must not be blocked simply because scopes differ.

### Reporting Groups

- A Reporting Group is a Pulse/business-owned semantic classification used for analysis.
- Reporting Groups are NOT KPIs and are NOT source/POS categories.
- Source main category, source subcategory, sales account, product name, channel, and similar source fields are source structure or metadata.
- The default Build 0.3.0 business groups are:
  - Add-ons
  - Non-Alcohol
  - Spirits/Cocktails
  - Coffee & Tea
  - Beer & Cider
  - Desserts
  - Wine & Sake
  - Starters
  - Mains
- These defaults are editable business configuration, not hard-coded eternal taxonomy.
- Reporting Groups have stable IDs. Display names may be edited without breaking mappings.
- Active/inactive status controls whether a group is offered to downstream Performance/reporting selectors.
- Inactive groups may retain historical mappings and lineage; deactivation must not delete mapping history.

### Mapping hierarchy, inheritance and precedence

The Mapping experience is intentionally inspired by the previously validated Lovable concept: the user browses the actual source/POS hierarchy, understands what sits underneath a node, maps at the highest safe level, and creates more-specific exceptions only where needed.

The architecture is hierarchy-based, not permanently hard-coded to exactly three levels. The currently observed Katria implementation is:

`Source Main Category -> Source Subcategory -> Product`

For Build 0.3.0, rules may therefore target Main Category, Subcategory, or Product. However, resolver code and documentation should express the general rule as:

**The most-specific applicable explicit mapping wins; otherwise inherit the nearest mapped ancestor; otherwise remain Unmapped.**

For the current hierarchy this resolves as:

1. Active Product mapping/override
2. Active Source Subcategory mapping
3. Active Source Main Category mapping
4. Unmapped / needs review

This precedence is an implementation of the hierarchy rule, not the permanent definition of the mapping architecture. Future source adapters may expose additional hierarchy levels without forcing a redesign of Reporting Groups or metrics.

A higher-level mapping is allowed only as an explicit human-authored rule. Descendants inherit it automatically unless they have a more-specific explicit rule. A product override may intentionally differ from its inherited parent mapping and must survive later parent remapping.

The Mapping UI/administration surface must make inherited versus explicit mappings visually distinguishable and must let the user inspect the products/descendants affected before or after assigning a broad parent rule.

Do not infer a business mapping from text similarity during deterministic resolution unless a future explicitly approved assistant/recommendation layer proposes it for human review.

Conflicting active rules at the same source node/scope and overlapping effective dates are invalid and must surface in QA; do not silently choose one.

### Metrics and KPIs

- Metrics/KPIs operate on Reporting Groups after mapping resolution.
- Definitions must be deterministic, explainable, and reproducible.
- Keep calculation logic centralized so Performance and Reports consume the same result rather than duplicate business logic.
- Build 0.2.0's Category Sales Share evolves semantically into Reporting Group Sales Share; preserve calculation behavior while changing the classification dimension.
- Do not hard-code 2025, prior year, prior week, Week 31, or a named channel as the universal comparison.

## Human control

- Do not auto-finalize mappings.
- Do not overwrite a user's explicit product-level override when a parent mapping changes.
- Preserve notes/audit fields where present.
- Surface unmapped products/classifications and conflicts clearly.
- Recommendations may be added later, but deterministic mapping must remain understandable without AI.

## Workbook product rules

- The workbook remains a user-facing product and a release artifact.
- The single authoritative development workbook is `OneDrive/Pulse/Development/Pulse_Current.xlsx`.
- Prefer the connected live Excel session for that canonical workbook. If it is unavailable, stop and ask the user; do not search Desktop, Downloads, or other folders for substitutes.
- Continue normal development in the same canonical workbook. Do not create QA, test, copy, or downloaded workbook variants unless the user explicitly requests a checkpoint or release copy.
- Preserve existing sheets/functions unless the active build specification explicitly migrates them.
- Do not remove development/engine sheets merely for visual simplification.
- Normal user-facing navigation remains goal-oriented: Overview -> Performance -> Reports -> Imports -> Settings.
- Mapping administration may live one level deeper than routine use.
- Performance should not expose the raw POS taxonomy once Reporting Groups are available.
- Avoid giant literal validation lists; use authoritative range/table-backed selectors.
- User-editable cells should be visually distinguishable but understated.
- Prefer tables/stable IDs/references over fragile cell-position coupling.

## Repository and build discipline

- A build number advances only with a concrete validated implementation or artifact.
- The repository owns implementation logic, tests, documentation, and reproducible build assets.
- Generated workbooks are release/checkpoint artifacts; they are not the only implementation source.
- Build scripts should be idempotent where practical, or explicitly fail safely when run against an incompatible checkpoint.
- Never require manual reconstruction of already encoded decisions when they can be migrated deterministically.
- Avoid broad refactors unrelated to the active build.

## Testing expectations

Every implementation task must include relevant deterministic checks.

For Mapping + Reporting Groups, verify at minimum:

- raw source sheets/facts unchanged by mapping operations;
- Reporting Group IDs are unique and stable;
- only active Reporting Groups are exposed downstream;
- most-specific explicit mapping wins, with Product > Subcategory > Main Category for the current Katria hierarchy;
- parent mapping inheritance works when no more-specific explicit mapping exists;
- product override survives parent remapping and is clearly distinguishable from inherited state;
- same-level conflicting active mappings are detected;
- unmapped items are visible and not silently classified;
- mapping changes alter metric classification without altering sales amount/quantity facts;
- total sales/quantity across all mapped + unmapped facts reconciles to source scope totals;
- Performance/Reports use the same calculation source;
- current/comparison dataset independence remains intact;
- no existing 0.2.0 QA fixes regress (readability, range-backed dropdowns, sensible widths).

## Coding-agent behavior

Before changing files:

1. Read the active build specification and relevant docs.
2. Inspect the current implementation and tests.
3. Identify migration impact on workbook tables, formulas, Office Scripts, and QA.

While changing files:

- Make small, reviewable changes.
- Prefer explicit names and stable identifiers.
- Preserve compatibility unless migration is intentional and documented.
- Do not invent business rules not present in the specification.

After changing files:

- Run available tests/checks.
- Summarize files changed, migration behavior, assumptions, and unresolved risks.
- Report validation failures rather than masking them.
