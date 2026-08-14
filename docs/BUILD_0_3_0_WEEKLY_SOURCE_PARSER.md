# Build 0.3.0 — Weekly Source Parser Foundation

## Scope

This slice implements only the read-only parser and source-manifest contract for
the candidate authoritative weekly POS corpus. It does not publish facts, build
the weekly analytical cache, mutate `Pulse_Current.xlsx`, supersede legacy
imports, move OneDrive files, or implement a Power Automate flow.

Production intake remains:

`OneDrive/Pulse/Incoming reports`

The local corpus is a caller-supplied, read-only development fixture. Repository
code never searches Desktop, Downloads, OneDrive, or other folders to locate it.

## Exact source scope that can be proven

Every accepted file contains one worksheet named `Sales per Item` with:

- `Period: YYYY-MM-DD - YYYY-MM-DD` in `A1`;
- the exact row-2 headers `Restaurant`, `Main Category`, `Sub Category`,
  `Sales Account`, `Item`, `Quantity`, and `Amount`; and
- one normalized observation for every populated source row after the header.

The file therefore proves a weekly POS **Sales per Item export scope** containing
the rows emitted by that export configuration. It preserves the POS hierarchy,
Sales Account, item, restaurant, Quantity, and Amount.

The file does **not** contain:

- a Channel column;
- an In-house/Takeaway flag;
- a persisted export-filter definition; or
- other metadata that proves whether a particular sales type was included or
  excluded before export.

Delivery-fee, delivery-named, gift-card, VAT-rate, campaign, or similar rows are
business observations, not reliable evidence of the report's channel filter.
The parser never infers scope from them.

The configured scope contract is therefore:

- `ScopeID = SCOPE-030-WEEKLY-SALES-PER-ITEM`;
- `SourceSystemID = SRC-TEST-SALES`, preserving the accepted mapping source ID;
- `ChannelScope = Source-defined scope; channel not encoded`;
- authority = the human-approved POS export configuration; and
- status = candidate authoritative weekly source until the historical model is
  explicitly cut over.

If the business later supplies a precise export label such as In-house-only,
that label should update the human scope contract. It must not be inferred by
the deterministic parser.

## Parser contract

`src/imports/weekly-sales-parser.mjs` is the deterministic contract used by
automated tests and corpus audits. `office-scripts/Parse_Weekly_Sales_Report.ts`
is its read-only Excel/Power Automate adapter.

The parser:

1. reads period identity only from `A1`;
2. requires exactly one Monday–Sunday ISO week;
3. requires the exact seven-column schema;
4. preserves source spelling, punctuation, `*`, capitalization, and spaces;
5. requires finite numeric Quantity and Amount cells;
6. rejects blank rows inside the populated body and unexpected populated
   columns;
7. returns one normalized row per source row with source-row lineage;
8. calculates source row count, Sales NOK, Quantity, restaurant/item/account
   coverage, and internal reconciliation status;
9. creates filename-independent semantic file and row fingerprints, using a
   fingerprint-only canonical representation of Quantity at six decimals and
   Sales NOK at two decimals so Excel and raw OOXML produce identical identity;
   and
10. records locator and binary fingerprint only as audit metadata.

Canonicalization never replaces the parsed numeric source values used for row
output, totals, reconciliation, or lineage fields.

No ProductID, RestaurantID, classification ID, mapping, Reporting Group, or
publication state is guessed or created in this slice.

## Corpus manifest contract

`buildWeeklyCorpusManifest()` combines parsed file manifests and reports:

- period coverage derived only from PeriodStart/PeriodEnd;
- duplicate periods and semantic source files;
- overlaps and gaps;
- schema/scope-contract consistency;
- source totals and row count;
- restaurant and Sales Account set signatures; and
- a deterministic corpus fingerprint.

The manifest applies no legacy-dataset reconciliation requirement. The new
weekly corpus must reconcile to its own source rows and approved scope. The old
2025 baseline and W31 facts remain intact and auditable until a separate,
explicit cutover.

## Validated local fixture evidence

The complete read-only development corpus currently produces:

| Check | Result |
|---|---:|
| Reports / unique periods | 84 / 84 |
| 2025 coverage | W01–W52 |
| 2026 coverage | W01–W32 |
| Source rows | 245,632 |
| Sales NOK | 484,728,367.25 |
| Quantity | 2,469,988.09 |
| Duplicate periods / files | 0 / 0 |
| Overlaps / gaps | 0 / 0 |
| Unique exact restaurant values | 18 |
| Unique exact Item values | 795 |
| Corpus fingerprint | `WSC-349b8bfd096ace2e` |
| Result | PASS |

The 18 restaurant values comprise the 16 operational names plus the two exact
test-department spellings already surfaced by the audit. They remain distinct
until the existing Pulse restaurant-resolution configuration explicitly
resolves them.

## Read-only fixture audit

The development command requires one exact path and never searches for files:

```text
node src/imports/audit-weekly-corpus.mjs <exact-read-only-fixture-path> \
  --expected tests/expected-build-0.3.0-weekly-source.json
```

The command reads ZIP/XML content, produces an in-memory manifest, and writes
nothing to the source corpus or workbook.

## Cutover boundary

This parser passing does not activate the weekly model. Before cutover, the next
approved slices must still provide:

- stable catalog reuse and unresolved-identity review;
- current hierarchical mapping application;
- compact weekly RPG/scope cache generation;
- per-week and corpus reconciliation for Sales NOK, Quantity, fact count, and
  mapping states;
- source-scope consistency approval; and
- atomic cache activation plus explicit legacy import supersession.

Differences from the old annual baseline/W31 fixture are evidence to explain by
scope, not parser failures when the two source scopes are not equivalent.

The first item is now implemented as the bounded read-only Weekly Identity
Preflight documented in `BUILD_0_3_0_WEEKLY_IDENTITY_PREFLIGHT.md`. It does not
activate the later cache or cutover steps.
