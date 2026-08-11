# tests

This directory contains deterministic validation and regression assets for Pulse.

Current assets include Build 0.2.0 validation material such as the build checklist and expected output data. Future build work should extend these tests rather than treating this directory as a placeholder.

Build 0.3.0 Phase 1 adds a Node test fixture for deterministic hierarchy
resolution and a workbook validation checklist. Run:

`node --test tests/build-0.3.0-phase1.test.mjs`

Build 0.3.0 Phase 2A adds deterministic bridge, fingerprint, reconciliation,
mapping-state, Reporting Group total, stale-state, and legacy-comparison
fixtures. Run the complete suite with:

`node --test tests/*.test.mjs`

The Lovable migration suite validates the frozen 129-decision contract,
Product exclusion precedence/lineage/denominator behavior, migration
idempotency and `MAP-000001` reuse, Office Scripts compatibility guards, and
the accepted checkpoint reconciliation targets.

For Build 0.3.0, tests should cover at minimum:

- Reporting Group registry integrity.
- Hierarchical mapping resolution.
- More-specific override precedence.
- Inherited mapping behavior.
- Unmapped handling.
- Active/inactive Reporting Group behavior.
- Reconciliation of mapped aggregates to immutable source sales facts.
- Regression protection for the validated 0.2.0 metric behavior where applicable.
- One derived row per immutable fact with full lineage.
- Reproducible mapping fingerprints and stale Effective Mapping rejection.
- Separate Mapped, Unmapped, Conflict, and Inactive Target coverage.
- Human-configured ID-based CAT/RPG equivalence and visible variance.

Tests should be deterministic, explainable, and tied to a reproducible build or workbook checkpoint.
