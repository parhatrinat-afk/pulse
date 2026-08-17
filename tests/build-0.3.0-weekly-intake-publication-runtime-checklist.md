# Build 0.3.0 — New-week Publication Runtime Checklist

- [ ] One exact fresh Active cache is resolved dynamically.
- [ ] Hidden accepted identity registry is present, fingerprinted, unique and
      bound to the Active Identity Preflight fingerprint.
- [ ] Accepted 84-week identities are reused before any new ID allocation.
- [ ] `Build Candidate` accepts only a complete reconciled parser result.
- [ ] Candidate has one additional unique period and remains Not Active.
- [ ] All period, five-state, mapped-RPG, Sales NOK, Quantity and fact-count
      reconciliation checks pass.
- [ ] Candidate WCV/WCC are deterministic and match repository evidence.
- [ ] Candidate writes are bounded and hidden.
- [ ] Candidate identity registry preserves all prior IDs and adds only genuine
      exact candidates.
- [ ] Prior Active remains authoritative after Build Candidate.
- [ ] Activation validates exact WCV/WCC, row counts and freshness again.
- [ ] Prior full version is retained on the hidden rollback surface.
- [ ] Canonical write failure restores prior Active cache and identity control.
- [ ] Canonical write failure also restores the accepted identity registry.
- [ ] Authority transition is the final analytical mutation.
- [ ] Exactly one Active / Active cache remains after publication.
- [ ] Published period appears exactly once and is selectable in Performance.
- [ ] Retry after completed activation returns Published idempotently.
- [ ] `Published` ledger evidence retains source locator, period and versions.
- [ ] Weekly Performance QA remains 16/16 PASS.
- [ ] Phase 2C Interaction QA remains 16/16 PASS.
- [ ] Performance and Reports still use existing canonical weekly table names.
- [ ] No facts, Mapping Rules, Products, Restaurants, Imports, legacy results,
      Performance formulas or Reports formulas are changed by the runtime.
- [ ] Duplicate, Correction Review, Rejected and Cache Stale routes are unchanged.
- [ ] Power Automate archives only Published + archiveReady=true.
- [ ] Existing 20-second stabilization and six-minute archive delays remain.
- [ ] W33 content resolves to `PERIOD-2026-08-10-2026-08-16`, 2,940 facts,
      NOK 5,636,773.50, Quantity 29,654.35 and `WSF-641061337dfbfd59` without
      filename/folder identity.
- [ ] W33 Candidate equals the independent 85-week rebuild at every business
      period/scope/RPG row and reconciles to `WCV-1b0b195c210da456` /
      `WCC-26c195956ebc2823`.
- [ ] Complete automated suite and Office Scripts static guards pass.
- [ ] `git diff --check` passes.
