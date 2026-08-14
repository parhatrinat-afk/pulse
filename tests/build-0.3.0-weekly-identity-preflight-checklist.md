# Build 0.3.0 Weekly Identity Preflight Checklist

- [x] Starts from weekly parser checkpoint `d91dff3263132d11bdccd19641532b134ef19538`.
- [x] Uses exact SourceSystemID + source-string keys; filenames/folders are ignored.
- [x] Reuses one exact current stable ID and never overwrites it.
- [x] Proposes deterministic IDs for new unambiguous exact identities.
- [x] Keeps identity creation independent from Reporting Group mapping.
- [x] Sends catalog collisions and new multi-hierarchy ProductKeys to Identity Pending.
- [x] Creates separate RST-0017/RST-0018 Test Department candidates with ReportingEnabled=No.
- [x] Preserves existing Product hierarchy as mapping authority; weekly paths are lineage/QA.
- [x] Surfaces exactly nine accepted existing-Product hierarchy divergence cases.
- [x] Highlights PRD-000689 Add-ons versus Mains without resolving it.
- [x] Accounts for every source row exactly once.
- [x] Reconciles source fact count, Sales NOK, and Quantity exactly.
- [x] Produces no duplicate candidate stable IDs or exact keys.
- [x] Rerun against accepted candidates is idempotent.
- [x] Full 84-week frozen checkpoint passes.
- [x] No source report, workbook, publication, cache, Power Automate, cutover, or Performance mutation.

Automated command:

```text
node --test tests/*.test.mjs
```

Read-only corpus command:

```text
node src/imports/audit-weekly-identity-preflight.mjs <exact-read-only-corpus-path> \
  --catalog tests/fixtures/build-0.3.0-weekly-identity-catalog.json \
  --expected tests/expected-build-0.3.0-weekly-identity.json
```
