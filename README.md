# Pulse

Pulse is an operational performance and reporting workbook. It turns trusted
weekly POS exports into deterministic, explainable management views while
preserving source lineage and human control over business classification.

## Release status

**Current candidate:** Pulse 0.3.0, validated through 2026 W33
**Remaining release gate:** one controlled Power Automate `New -> Published`
pilot using the production flow

The repository is the source of truth for scripts, tests, documentation and
reproducible build logic. The single development workbook is
`OneDrive/Pulse/Development/Pulse_Current.xlsx`; workbook binaries and weekly
source reports are not stored in Git.

## Normal workbook experience

Pulse exposes six normal-use sheets:

1. **Overview** — current management summary, attention and data status.
2. **Performance** — independent Current/Compare weekly ranges, restaurant and
   Reporting Group selection, shares, PP change, Sales NOK and NOK Impact.
3. **Reports** — meeting-ready detail linked to the same Performance result.
4. **Imports** — latest published week, historical coverage and genuine intake
   activity.
5. **Mapping** — hierarchy browsing, bulk assignment, inherited/custom state
   and weekly historical attention.
6. **Settings** — application configuration and Reporting Group administration.

Engineering, lineage, reconciliation and rollback surfaces remain hidden but
available for controlled maintenance and QA.

## Current product foundation

- Weekly coverage: **2025 W01–W52 and 2026 W01–W33**.
- A versioned compact weekly cache keeps one Active authority and one previous
  full rollback version.
- The Power Automate intake contract supports New, Duplicate, Correction
  Review, Rejected and Cache Stale outcomes.
- Duplicate reports are safe no-ops. A New report is first built as a complete
  inactive Candidate, reconciled, then activated through a final authority
  switch.
- Mapping uses stable IDs and the source hierarchy. The most-specific active
  explicit rule wins; otherwise a product inherits the nearest mapped ancestor
  or remains Unmapped.
- Reporting Groups are editable business classifications, not KPIs or POS
  categories. Current runtime and weekly cache logic support a dynamic positive
  number of active groups.
- Performance and Reports consume the same centralized weekly additive
  components. Shares are calculated only after period/scope aggregation.
- Unmapped, Identity Pending, Conflict and Inactive Target facts remain visible
  and in source denominators; they never silently disappear.
- Raw source reports, legacy facts and historical lineage are not rewritten by
  mapping.

## Architecture

`Untouched weekly POS report -> Parser -> Identity preflight -> Hierarchical Mapping -> Versioned weekly cache -> Metrics -> Performance / Reports / Overview`

Source and business semantics remain separate:

- POS main category, subcategory and product are source hierarchy.
- Reporting Groups are business-owned classification.
- Metrics and KPIs operate on Reporting Groups.
- Current and comparison periods are independent selections.
- Mapping and identity changes require a truthful cache refresh rather than
  rewriting history.

See [architecture.md](docs/architecture.md),
[pulse-principles.md](docs/pulse-principles.md), and the
[operations runbook](docs/BUILD_0_3_0_OPERATIONS_RUNBOOK.md).

## Development and release discipline

- Do not search for or substitute workbook copies. Use the connected canonical
  workbook or stop and ask for access.
- Do not commit source-report or development-workbook binaries.
- Run the complete deterministic suite with:

  `node --test tests/*.test.mjs`

- A build becomes a release only after repository validation, live Excel QA,
  the production Power Automate New-path pilot, controlled merge and explicit
  release acceptance.
