# imports

Build 0.3.0 begins the bounded weekly source parser foundation here.

- `weekly-sales-parser.mjs` validates the exact weekly `Sales per Item` matrix,
  preserves source strings, derives period identity only from the internal
  `Period:` field, calculates source totals, and produces deterministic
  filename-independent file/row manifests.
- `audit-weekly-corpus.mjs` is a read-only development command. It requires one
  exact caller-supplied fixture path and never searches for or modifies source
  files. Production intake remains OneDrive + Power Automate + Office Scripts.
- `weekly-identity-preflight.mjs` reuses exact current Pulse catalog identities,
  proposes deterministic IDs for new exact identities, preserves current
  Product hierarchy as mapping authority, and surfaces collisions/multi-path
  identities as `Identity Pending` with reconciled impact.
- `audit-weekly-identity-preflight.mjs` runs that contract against one exact
  caller-supplied corpus plus a frozen accepted catalog snapshot. It remains
  read-only and creates no cache or workbook state.
- `weekly-compact-cache.mjs` builds a validated candidate with one denominator
  row per Restaurant/week and nine dense additive RPG rows per Restaurant/week.
  It reuses the accepted identity preflight, current hierarchy resolver, and
  shared Performance restaurant-scope fingerprint.
- `audit-weekly-compact-cache.mjs` runs that candidate against one exact corpus
  path and explicit accepted mapping/preflight fingerprints. It emits evidence
  only; it does not materialize or activate a workbook cache.
- `weekly-cache-activation.mjs` defines the two-field activation transition and
  the minimum freshness guard future weekly consumers must pass. It treats
  mapping content, catalog/identity content, ReportingEnabled scope and cache
  QA as authority; date-only mapping audit changes are intentionally ignored.

None of these foundations publishes weekly facts or supersedes legacy imports.
Activation changes only the accepted cache-version authority; Performance does
not consume the weekly cache until a separately approved cutover. The identity
preflight may evaluate proposed ProductIDs through the existing mapping
resolver, but it never creates mapping rules or treats identity creation as a
business mapping decision.

The weekly export has no encoded Channel/filter field. Scope is a human-owned
manifest contract (`SCOPE-030-WEEKLY-SALES-PER-ITEM`), not an inference from
filenames, folders, categories, Sales Accounts, or product names.
