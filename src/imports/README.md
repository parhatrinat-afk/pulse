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

Neither slice publishes facts, builds the compact analytical cache, or
supersedes legacy imports. The preflight may evaluate proposed ProductIDs through
the existing mapping resolver, but it never creates mapping rules or treats
identity creation as a business mapping decision.

The weekly export has no encoded Channel/filter field. Scope is a human-owned
manifest contract (`SCOPE-030-WEEKLY-SALES-PER-ITEM`), not an inference from
filenames, folders, categories, Sales Accounts, or product names.
