# Build 0.3.0 UX — Information Architecture Slice 1

## Purpose

This slice gives the accepted Phase 2C workbook a clear normal workflow without
changing its analytical architecture or hiding audit surfaces permanently.

Normal reporting flow:

`Overview -> Performance -> Reports`

Primary administration:

`Imports -> Mapping -> Settings`

The implementation is owned by
`office-scripts/Build_0_3_0_UX_IA_Slice1.ts`. Run it after the accepted Phase
2C build. It is safe to rerun against the same compatible checkpoint.

## Mandatory preflight

Before changing any sheet state, the script requires:

- the exact accepted 47-sheet Phase 2C workbook contract;
- `tblPerformanceInteractionQA`;
- exactly `QA-0302C-01` through `QA-0302C-16`;
- `PASS` for every one of those checks.

An incompatible or failed checkpoint stops with a visible `PUL-030UX-*` error
before information-architecture mutation begins.

## Default workbook structure

The following sheets are visible and first, in this exact order:

1. Overview
2. Performance
3. Reports
4. Imports
5. Mapping
6. Settings

The remaining 41 accepted sheets stay present but are normally hidden:

- KPI Registry
- Context
- Views
- Restaurants
- Reporting Categories
- Products
- Source Classifications
- Remap Assistant
- Remap Rules
- Import Exclusions
- Expected Coverage
- Publication Control
- Import Actions
- Import Certificates
- Domains
- Source Systems
- Adapters
- Adapter Contract
- Import Control
- Test Run Control
- _Raw_2025_Baseline
- _Raw_2026_Week31
- _Sales_Facts
- _Standard_Staging
- _Remap_Audit
- _Import_Action_Audit
- _Build_Log
- _Environment
- _Lists
- Effective Categories
- _Metric_Calc
- Reporting Groups
- Mapping Rules
- Effective Mapping
- _Mapping_Lists
- Mapping QA
- Metric Contract
- Metric Equivalence
- _Metric_RPG_Facts
- Metric Migration QA
- Metric Results QA

The script uses ordinary hidden state, not `veryHidden`. This is workflow
organization, not security; users can unhide supporting sheets when needed.
No sheet protection is added.

## Navigation and saved views

Overview cells E8:E12 contain ordinary workbook hyperlinks, in order, to:

1. Performance
2. Reports
3. Imports
4. Mapping
5. Settings

Every target opens at A1. Each of the six primary sheets is activated and A1
selected once during the script, then Overview A1 is selected last. This resets
the saved starting position, including the former Performance row-16 position.

This first slice intentionally does not add a larger navigation system or
back-to-Overview links. Those links are optional and can be assessed with later
visual work without adding content to five more sheets now.

## Frozen analytical surfaces

The script's only workbook mutations are sheet visibility, sheet position, and
the five required Overview hyperlinks. It does not write formulas or analytical
values, add/delete/resize tables, delete sheets, move functional ranges, or
change formats.

The accepted Phase 2C script and interactive reporting module are hash-locked
by regression tests for this slice. Phase 2C calculations, selections, sorting,
Total/Grand Total, PP Change, NOK Impact, the numeric-helper/text-facade model,
Reports linkage, mappings, imports, facts, and QA logic remain unchanged.

## Execution sequence

1. Open the canonical `OneDrive/Pulse/Development/Pulse_Current.xlsx` in Excel
   for the web.
2. Confirm `tblPerformanceInteractionQA` is 16/16 PASS.
3. Add or replace the Office Script with
   `Build_0_3_0_UX_IA_Slice1.ts` from the repository.
4. Run the script once.
5. Continue in the same canonical workbook; do not create an ad-hoc copy.
6. Complete the live checklist in
   `tests/build-0.3.0-ux-ia-slice1-checklist.md`.
