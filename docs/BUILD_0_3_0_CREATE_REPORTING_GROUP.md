# Build 0.3.0 — Create Reporting Group

## Scope

This slice adds one bounded administrator action: create a new active Reporting
Group from `Settings`. It does not add rename, deactivate, delete, mapping, or
historical-cache rebuild behavior.

`tblReportingGroups` remains the single authority. The user supplies a business
name and optional description/notes; Pulse allocates the stable ID and sort
order. IDs are derived from the highest ID ever issued, including inactive
rows, and are never reused.

## Settings workflow

The visible `Settings` sheet contains:

- a full-width Settings header and a readable three-column Application
  Settings block whose existing table contract remains unchanged;
- a business-facing Reporting Groups overview: Reporting Group, Status,
  Products, and Sales NOK;
- a compact New Reporting Group form: Name, optional Description, optional
  Notes, and status;
- the instruction to run the user-facing Office Script **Create Reporting
  Group**.

IDs remain backstage. Products and Sales NOK are derived from the accepted
Mapping membership projection; a newly created group begins at zero.

## Creation contract

`office-scripts/Create_Reporting_Group.ts` validates all inputs and downstream
surfaces before appending one row to `tblReportingGroups`:

- the authority has valid, unique stable IDs and SortOrder values;
- at least one active `DOM-SALES` Reporting Group exists;
- the new business name is nonblank and does not duplicate an active name;
- the next ID and SortOrder are unique;
- Mapping, Phase 2C, and Weekly Performance checkpoint QA is present.

On success the row is Active, uses `DOM-SALES`, preserves the optional text,
and creates no Mapping Rule. Re-running the same name fails clearly without a
duplicate. A failure after append removes the appended row and restores the
derived selectors/catalogs.

## Downstream behavior

The creation action refreshes the count-driven business surfaces only:

- Mapping browse, Assign-to, overview, and reverse-membership selection include
  the new business name; membership is initially zero;
- the Performance Reporting Group selection table gains one row, preserves
  existing choices by ReportingGroupID, and defaults the new row to `No`;
- the Detail Reporting Group and Sort lists include the new group;
- Reports continues to follow the Performance detail result.

The active historical cache is not rebuilt. Because the active group catalog is
part of both MappingContentFingerprint and CatalogContentFingerprint, creation
makes the current weekly cache stale. Mapping shows **Performance refresh
required** and Performance becomes unavailable until the existing deterministic
cache rebuild/activation path creates a complete dense cache for the new active
group set. A zero-member group is still represented by zero additive rows in
that future dense version.

The creation action expands the current selection and selector helper catalogs
immediately. The numerical weekly component grid remains unavailable while the
cache is stale; the existing deterministic cache rebuild/activation followed by
the weekly Performance refresh installs the complete component and visible
matrix geometry for the expanded active-group set.

## Reversible live QA — 2026-08-22

The connected canonical `Pulse_Current.xlsx` passed the controlled temporary
`RPG-0010` proof:

- exactly one Active `RPG-0010` / SortOrder `100` row was created;
- Settings and Mapping showed the QA business name with zero Products, Facts,
  and Sales NOK;
- Mapping Rules and Effective Mapping remained 133 and 1,041 rows;
- Performance selection expanded to ten rows, preserved the nine existing
  stable-ID choices, and defaulted `RPG-0010` to `No`;
- Detail and Sort helper lists contained the tenth business name;
- mapping/catalog fingerprints changed deterministically to
  `MCF-0f4b33742afe9302` / `ICC-2cb8d60253ef779c`;
- Mapping showed **Performance refresh required** and weekly Performance showed
  **Stale / unavailable**; active and rollback cache evidence did not change;
- Mapping, Phase 2C, and Weekly Performance checkpoint tables remained 9/9,
  16/16, and 16/16 PASS as preserved regression evidence.

The QA-only row was then removed. The canonical workbook finished with nine
groups, no `RPG-0010`, the original fingerprints
`MCF-759cc92c4304a913` / `ICC-5644a77c18a97437`, weekly Performance
**Available**, and the original active/rollback cache authority.

## Explicit exclusions

This slice does not implement rename, deactivate, or delete. It does not mutate
facts, identities, Mapping Rules, Effective Mapping, cache rows, legacy result
tables, or KPI definitions.
