# Build 0.3.0 — Candidate Compact Weekly Cache

## Scope

This bounded slice builds and validates a compact weekly analytical candidate in
repository/test space. It consumes the accepted 84-week parser output, the
accepted Weekly Identity Preflight, and the current Pulse hierarchical mapping
resolver. It does not write `Pulse_Current.xlsx`, activate a cache, alter
Performance, create period selectors, supersede the legacy 2025/W31 imports, or
implement production intake.

The source corpus is 2025 W01-W52 plus 2026 W01-W32 under the accepted source
scope contract `SCOPE-030-WEEKLY-SALES-PER-ITEM`. Filenames and folders remain
non-authoritative.

## Candidate contract

The build is a pure, all-or-nothing in-memory operation. It first requires the
accepted identity-preflight fingerprint plus the exact current mapping
fingerprint and MappingAsOfDate. A difference fails visibly before analytical
rows are returned. A candidate with the same version as an already active cache
cannot overwrite it.

The candidate version is `WCV-2cd012763d86a794`. It is recorded as `Candidate`
and `Not Active`; no active Performance model consumes it.
Reporting Group membership is fingerprinted by active stable RPG IDs. Editing a
display name or sort order does not change cache grain or numeric membership.
Because the accepted preflight fingerprints its complete catalog snapshot, any
catalog edit still requires an explicit refreshed preflight/candidate version
rather than bypassing staleness review.

### Weekly RPG cache

Grain:

`CacheVersion × SourcePeriodKey × RestaurantID × ReportingGroupID`

The cache is dense across the nine active Reporting Groups for every observed
Restaurant/week scope. Zero rows are intentional. It stores only:

- deterministic `WeeklyRPGCacheRowID`;
- the four grain fields;
- Mapped fact count;
- Mapped Sales NOK; and
- Mapped Quantity.

Only facts in `Mapped` state contribute to these numerators. The cache does not
store or repeat a denominator.

### Weekly scope cache

Grain:

`CacheVersion × SourcePeriodKey × RestaurantID`

One row stores the source denominator once for that Restaurant/week plus
fact-count, Sales NOK, and Quantity components for each explicit state:

- Mapped;
- Unmapped;
- Identity Pending;
- Conflict; and
- Inactive Target.

`PerformanceEligible` preserves the shared Performance restaurant-scope
contract (`Status=Active` and `ReportingEnabled=Yes`) without deleting facts.
The two Test Department identities remain separate and ineligible.

### Minimal manifests

The period/source manifest contains one row per accepted week with its period,
source identities/fingerprints, scope contract, source totals, and source
restaurant count. The cache-version manifest holds version status, schema and
source contract versions, the corpus/preflight/catalog/mapping/group/scope/cache
fingerprints, MappingAsOfDate, and row counts. These values are not duplicated
onto every analytical row.

## Identity and mapping behavior

- The accepted exact identity preflight remains the only weekly identity
  contract.
- The existing hierarchical resolver remains the only mapping engine.
- Current Pulse Product hierarchy remains mapping authority; weekly alternate
  paths are lineage/QA evidence only.
- The six unresolved products remain `Identity Pending`. Their 120 facts, NOK
  114,876, and Quantity 951 stay in the scope denominator and source
  reconciliation but never enter an RPG numerator.
- The nine accepted hierarchy-divergence ProductIDs retain current hierarchy and
  mapping, including PRD-000689 Red Curry as RPG-0001.
- RST-0017 and RST-0018 preserve 983 facts while remaining outside normal
  Performance scope.

## Frozen 84-week result

| Measure | Result |
|---|---:|
| Period manifest rows | 84 |
| Cache-version rows | 1 |
| Weekly scope rows | 1,421 |
| Dense weekly RPG rows | 12,789 |
| Nonzero weekly RPG rows | 11,688 |
| Analytical rows (scope + RPG) | 14,210 |
| Complete rows including manifests | 14,295 |
| Candidate cells | 136,359 |

Deterministic fingerprints:

- cache version: `WCV-2cd012763d86a794`;
- cache contents: `WCC-7bd0c5f845b2a36d`;
- source corpus: `WSC-349b8bfd096ace2e`;
- identity preflight: `IDP-4cd1159238339096`;
- mapping: `MAP-342029f71a922b47` as of `2026-08-12`; and
- shared enabled-restaurant scope: `RSC-08df626f217dd94b`.

## Reconciliation

All 84 individual weeks pass. The year and corpus totals also reconcile exactly
across Mapped + Unmapped + Identity Pending + Conflict + Inactive Target.

| Scope/state | Facts | Sales NOK | Quantity |
|---|---:|---:|---:|
| 2025 source | 155,349 | 302,378,436.48 | 1,525,252.09 |
| 2025 Mapped | 144,069 | 296,471,168.13 | 1,451,008.70 |
| 2025 Unmapped | 11,228 | 5,905,861.35 | 74,159.39 |
| 2025 Identity Pending | 52 | 1,407.00 | 84.00 |
| 2026 W01-W32 source | 90,283 | 182,349,930.77 | 944,736.00 |
| 2026 W01-W32 Mapped | 82,416 | 177,689,373.47 | 889,097.91 |
| 2026 W01-W32 Unmapped | 7,799 | 4,547,088.30 | 54,771.09 |
| 2026 W01-W32 Identity Pending | 68 | 113,469.00 | 867.00 |
| Complete source | 245,632 | 484,728,367.25 | 2,469,988.09 |
| Complete Mapped | 226,485 | 474,160,541.60 | 2,340,106.61 |
| Complete Unmapped | 19,027 | 10,452,949.65 | 128,930.48 |
| Complete Identity Pending | 120 | 114,876.00 | 951.00 |
| Complete Conflict / Inactive Target | 0 / 0 | 0 / 0 | 0 / 0 |

The dense RPG numerators reconcile exactly to the Mapped scope components.
Normal Performance scope contains 244,649 facts, NOK 484,379,675.25, and
Quantity 2,468,270.61. The two excluded Test Department IDs contain 983 facts,
NOK 348,692, and Quantity 1,717.48. Enabled plus excluded equals the complete
source corpus.

## Independent range evidence

Ranges aggregate weekly additive components first. Shares and impacts are then
calculated once; weekly percentages are never averaged.

| Range | RPG | Current share | Comparison share | PP change | Current Sales NOK | NOK Impact |
|---|---|---:|---:|---:|---:|---:|
| W31 2026 vs W31 2025 | Add-ons | 1.1883% | 0.6701% | +0.5182 pp | 80,263.00 | +35,001.58 |
| W31 2026 vs W31 2025 | Beer & Cider | 7.2264% | 7.2917% | -0.0653 pp | 488,112.00 | -4,412.45 |
| W31 2026 vs W31 2025 | Mains | 62.1651% | 62.9249% | -0.7597 pp | 4,198,984.00 | -51,316.73 |
| W01-W32 2026 vs W01-W32 2025 | Add-ons | 1.3074% | 0.7336% | +0.5737 pp | 2,383,679.00 | +1,046,091.46 |
| W20-W30 2026 vs W20-W30 2025 | Add-ons | 1.2494% | 0.6825% | +0.5669 pp | 835,122.00 | +378,929.36 |

NOK Impact uses the accepted current-sales-base formula:

`Current RPG Sales - (Comparison RPG Sales / Comparison Total Sales) × Current Total Sales`

## Scale conclusion

The candidate serializes to about 4.20 MB as verbose JSON. The projected
incremental `.xlsx` footprint is approximately 0.76-2.31 MB for 136,359 cells.
At 14,210 analytical rows, the compact weekly model is suitable for Excel for
the web and is materially smaller than publishing 245,632 detailed rows twice.
Runtime/materialization still requires a bounded live Excel test before any
activation decision.

## Running the read-only cache audit

The command requires one exact fixture corpus path and explicit accepted state:

```text
node src/imports/audit-weekly-compact-cache.mjs <exact-read-only-corpus-path> \
  --catalog tests/fixtures/build-0.3.0-weekly-identity-catalog.json \
  --mapping-fingerprint MAP-342029f71a922b47 \
  --mapping-as-of-date 2026-08-12 \
  --preflight-fingerprint IDP-4cd1159238339096 \
  --expected tests/expected-build-0.3.0-weekly-compact-cache.json
```

It reads fixtures and emits JSON only. It never searches for reports, changes a
source file, writes a workbook, or activates a cache.

## Next live Excel boundary

No live Excel action is required for this repository-only slice. A separately
approved materialization step must first read the connected canonical
`Pulse_Current.xlsx`, verify its exact current catalog/mapping date and
fingerprint, and fail before writing if stale. It should write a complete
candidate version to new bounded tables, validate the frozen counts and
reconciliations, leave it inactive, prove the existing Phase 2C 16/16 QA and
Performance results unchanged, and only then expose it for human review.
