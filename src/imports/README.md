# imports

Build 0.3.0 begins the bounded weekly source parser foundation here.

- `weekly-sales-parser.mjs` validates the exact weekly `Sales per Item` matrix,
  preserves source strings, derives period identity only from the internal
  `Period:` field, calculates source totals, and produces deterministic
  filename-independent file/row manifests.
- `audit-weekly-corpus.mjs` is a read-only development command. It requires one
  exact caller-supplied fixture path and never searches for or modifies source
  files. Production intake remains OneDrive + Power Automate + Office Scripts.

The parser does not publish facts, resolve or create stable business IDs, apply
mapping, build the compact analytical cache, or supersede legacy imports.

The weekly export has no encoded Channel/filter field. Scope is a human-owned
manifest contract (`SCOPE-030-WEEKLY-SALES-PER-ITEM`), not an inference from
filenames, folders, categories, Sales Accounts, or product names.
