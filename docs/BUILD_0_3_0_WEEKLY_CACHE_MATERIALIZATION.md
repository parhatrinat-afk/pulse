# Build 0.3.0 — Candidate Weekly Cache Excel Materialization

## Scope and status

This bounded slice materializes the already accepted 84-week compact-cache
candidate in the canonical `Pulse_Current.xlsx`. It does not rebuild the cache
inside Excel and does not activate it. Performance, Reports, `_Metric_Calc`,
legacy imports, and the accepted 2025/W31 datasets remain unchanged.

The candidate is identified by:

- MappingContentFingerprint `MCF-759cc92c4304a913`;
- catalog content fingerprint `ICC-5644a77c18a97437`;
- Identity Preflight `IDP-062c182f23905ae8`;
- cache version `WCV-1a34ad1f46763d9b`;
- cache fingerprint `WCC-508dd608166cdb6e`;
- Phase 2A audit mapping `MAP-34202a7a1a922bd0` as of `2026-08-12`; and
- enabled restaurant scope `RSC-08df626f217dd94b`.

Weekly freshness is gated by MappingContentFingerprint. The Phase 2A/2B
date-sensitive mapping fingerprint remains unchanged and is stored separately
as audit metadata.

## Workbook surface

One hidden engineering sheet, `_Weekly_Cache`, contains four independent Excel
Tables:

| Table | Grain / purpose | Rows |
|---|---|---:|
| `tblWeeklyCacheVersions` | Candidate version, status, audit fingerprints and row counts | 1 |
| `tblWeeklyPeriodManifest` | One accepted source period/week | 84 |
| `tblWeeklyScopeCache` | CacheVersion × SourcePeriodKey × RestaurantID denominator/state coverage | 1,421 |
| `tblWeeklyRPGCache` | Dense CacheVersion × SourcePeriodKey × RestaurantID × ReportingGroupID numerators | 12,789 |

The four tables contain 14,295 data rows in total. The sheet is not part of the
six-sheet normal Pulse workflow. The version row remains `Candidate` / `Not
Active`; no active formula references these tables.

## Safe materialization contract

`src/imports/weekly-cache-materialization.mjs` converts the validated candidate
to explicit table schemas and bounded row chunks. The read-only payload audit
requires one exact fixture corpus path and all accepted fingerprints. The
Office Script uses three operations:

1. `Prepare` validates live mapping/catalog/Phase 2A/restaurant/QA state before
   creating a hidden staging surface.
2. `Write` accepts only bounded chunks at the predefined section geometry.
3. `Finalize` rereads the complete staging surface, validates exact row IDs,
   grain uniqueness, state and source reconciliation, mapping/cache
   fingerprints, range fixtures, and 16/16 Phase 2C QA before creating tables
   and renaming the surface.

A failed or interrupted write remains an explicitly named hidden staging sheet
without cache tables and cannot be mistaken for a valid candidate. `Prepare`
can discard and rebuild that staging surface. An exact existing inactive
candidate returns `Already Materialized`; an active or differing cache version
cannot be overwritten.

Excel stores ISO date strings written to date-shaped cells as serial numbers.
The fingerprint postcondition canonicalizes only those materialized period
dates back to `YYYY-MM-DD`. The raw Excel-serialized comparison fingerprint is
`WCC-1f82cc085787d84d`; it matches the repository candidate after applying the
same Excel date coercion. The authoritative canonical cache fingerprint remains
`WCC-508dd608166cdb6e`.

## Accepted reconciliation and fixtures

The materialized candidate must prove:

- 84/84 periods reconcile;
- 245,632 source facts;
- NOK 484,728,367.25 Sales;
- Quantity 2,469,988.09;
- Identity Pending: 120 facts / NOK 114,876 / Quantity 951;
- Conflict and Inactive Target: zero;
- dense RPG numerators equal Mapped scope components; and
- enabled plus excluded restaurant scope equals the complete source corpus.

Add-ons is calculated directly from the materialized additive rows:

| Range | Current share | Comparison share | PP change | Current Sales NOK | NOK Impact |
|---|---:|---:|---:|---:|---:|
| W31 2026 vs W31 2025 | 1.1883% | 0.6701% | +0.5182 pp | 80,263 | +35,001.58 |
| W01-W32 2026 vs W01-W32 2025 | 1.3074% | 0.7336% | +0.5737 pp | 2,383,679 | +1,046,091.46 |
| W20-W30 2026 vs W20-W30 2025 | 1.2494% | 0.6825% | +0.5669 pp | 835,122 | +378,929.36 |

All percentages and NOK Impact are calculated after aggregating additive
components. Weekly percentages are never averaged.

## Non-regression boundary

Materialization is additive infrastructure only. Phase 2C Interaction QA must
remain 16/16 PASS. Protected fingerprints for Performance, Reports, and
`tblMetricRPGResults` must be identical before and after. No period selector,
cache activation, import supersession, source publication, or Performance
cutover is part of this slice.

## Live Excel materialization evidence

The candidate was materialized successfully on `2026-08-14` in the canonical
`Pulse_Current.xlsx` workbook. The hidden `_Weekly_Cache` surface contains the
four tables and exact row counts above; `_Weekly_Cache_Staging` is absent after
the atomic rename. The version row is `Candidate` / `Not Active` with cache
fingerprint `WCC-508dd608166cdb6e`.

Live table-level checks confirmed:

- 245,632 source facts, NOK 484,728,367.25 Sales, and Quantity 2,469,988.09;
- 226,485 Mapped facts, 19,027 Unmapped facts, and 120 Identity Pending facts;
- Identity Pending NOK 114,876 and Quantity 951;
- zero Conflict and zero Inactive Target;
- dense RPG totals equal the Mapped scope totals;
- the two Performance-ineligible Test Department identities remain separate as
  `RST-0017` and `RST-0018`, with their 983 facts, NOK 348,692, and Quantity
  1,717.48 preserved outside Performance scope; and
- the accepted Add-ons range fixtures were reproduced from the staged candidate
  immediately before the four validated ranges were converted to tables and the
  unchanged sheet was atomically renamed.

The workbook contains 48 sheets after adding the hidden engineering surface.
The only visible sheets remain Overview, Performance, Reports, Imports,
Mapping, and Settings. Phase 2C Interaction QA remained 16/16 PASS. Protected
before/after fingerprints were identical for Performance
(`3988af7020d27700`), Reports (`35b417d137ae52ed`), and
`tblMetricRPGResults` (`4e01bd9915be5067`).
