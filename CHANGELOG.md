# Changelog

All meaningful Pulse changes are recorded here.

## Unreleased — 0.3.0 — Mapping + Reporting Groups

### Planned / in development

- Introduce business-defined Reporting Groups as a semantic layer between source classifications and metrics.
- Preserve source hierarchy and raw source data rather than replacing source classifications.
- Support hierarchical mapping: map at the highest safe source level and inherit downward.
- Allow more-specific explicit mappings to override inherited mappings.
- Keep the mapping resolver conceptually hierarchy-aware rather than permanently coupling it to one POS structure.
- Allow Reporting Groups to be active or inactive; active groups become available to Performance.
- Keep Reporting Groups separate from KPI definitions.
- Prepare Performance to consume Reporting Groups instead of exposing the full raw POS classification set.
- Preserve human control, deterministic resolution, traceability, and recoverability.

This section describes active development only. It does not advance the validated release checkpoint.

## [0.2.0-QA] — Current validated checkpoint

### Validated

- Build 0.2.0 workbook and deterministic Category Sales Share vertical slice were subjected to pre-0.3.0 QA.
- QA fixes and supporting documentation were added without changing the core product principles.
- The resulting `Pulse_Build_0_2_0_QA` workbook is the checkpoint used as the starting point for Build 0.3.0.

## [0.2.0] — Category Sales Share

### Added

- Reproducible Office Script build path.
- First deterministic KPI: Category Sales Share.
- Independent current/comparison dataset selection.
- Independent current/comparison channel selection.
- All-channels scope.
- Non-blocking scope mismatch visibility.
- Performance and Reports vertical slice.

### Corrected

- Removed In-house from the KPI definition.
- Removed the assumption that 2025 is permanently the baseline.
- Comparison/baseline reference is user-selected.

### Preserved

- Immutable sales facts.
- Human decision ownership.
- Source neutrality.
- Remap Assistant and Remap Rules naming during the 0.2.0 implementation.

## [0.1.1]

### Changed

- Reorganized the workbook around a user-first experience.
- Primary workflow: Overview, Performance, Reports, Imports, Settings.
- Supporting administration and engine sheets retained rather than removed.
- Refined workbook hierarchy and presentation.

### Preserved

- Existing import architecture.
- Fact store and mapping structures.
- KPI Registry.
- Remap Assistant and Remap Rules names.
- 2025 baseline and Week 31 test data contained in the workbook.

## [0.1.0]

Initial structured Pulse workbook checkpoint containing the core data model, imports, mapping, facts, configuration, and product shell.
