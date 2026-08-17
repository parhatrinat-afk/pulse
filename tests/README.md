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

Phase 2B adds bridge-cutover freshness checks, deterministic centralized result
IDs/grain, all-state denominator tests, independent nine-RPG dataset fixtures,
accepted Add-ons targets, enabled/deselected restaurant scope reconciliation,
dynamic 306/288 cardinality, Performance formula guards, Reports linkage, and
Office Scripts compiler/performance compatibility checks.

Phase 2C adds authoritative stable-ID Include-state preservation, newly eligible
default-No behavior, additive arbitrary-restaurant aggregation, weighted Grand
Total, five display modes including current-sales-base NOK Impact,
zero-denominator display guards, Company-control reconciliation,
numeric-helper/text-facade isolation, removal of conditional number formats,
selected-RPG Total semantics, full-precision presentation-only sorting,
formula-only interaction guards, and Office Scripts static compatibility checks.

UX Information Architecture Slice 1 adds deterministic coverage for the exact
six-sheet visible workflow/order, preservation of all 47 accepted sheets,
ordinary Overview navigation, A1 saved views, 16/16 Phase 2C preflight,
restricted visibility/position/hyperlink mutation, accepted Phase 2C source
hashes, and Office Scripts iterator/read-in-loop compatibility guards.

UX Visual Slice 2A adds accepted-source hash locks, exact checkpoint/IA/QA
preflight, primary-only gridline guards, formula and table content/schema
fingerprints, structural Reports clipping safeguards, compact user-facing
Performance Explain checks, restricted value-write checks, idempotence guards,
and Office Scripts static compatibility/syntax checks.

UX Visual Slice 2B locks the accepted Phase 2C/IA/Slice 2A sources and proves
the pre-checkpoint cleanup can change only bounded Performance/Mapping layout,
semantic presentation colors, and the Mapping date display. Formula, value,
table, validation, IA, and 16/16 QA guards remain mandatory.

The weekly source parser suite validates internal Period identity, exact schema,
Monday–Sunday ISO weeks, exact-string preservation, filename independence,
runtime-independent numeric canonicalization for deterministic
semantic/file/row/corpus fingerprints, self-reconciliation,
duplicate/overlap/gap detection, explicit non-inference of channel scope, and a
read-only Office Scripts adapter. The local 84-file fixture corpus is not stored
in Git; its accepted aggregate evidence is frozen in
`expected-build-0.3.0-weekly-source.json`.

The weekly identity preflight suite validates exact stable-key reuse,
deterministic new Restaurant/Product/classification candidates, explicit
ambiguity handling, Test Department separation, idempotent reruns, current
Product-hierarchy authority, the frozen nine-item divergence review, and exact
five-state fact/Sales/Quantity reconciliation. The 84-week accepted identity
checkpoint is frozen in `expected-build-0.3.0-weekly-identity.json`; its catalog
input is the read-only `fixtures/build-0.3.0-weekly-identity-catalog.json`
snapshot rather than a workbook artifact.

The candidate compact weekly cache suite validates deterministic candidate
versioning and row IDs, one denominator per Restaurant/week, dense nine-RPG
numerators, all five mapping/identity states, Test Department scope exclusion,
current Product-hierarchy authority, stale-state rejection, complete weekly and
range aggregation, and the frozen 84-week reconciliation checkpoint.

The Excel materialization suite validates the four explicit hidden-table
schemas, bounded chunk geometry, inactive/idempotent staging contract,
date-neutral live freshness gate, Excel date canonicalization, exact frozen
row counts, five-state reconciliation, range fixtures, Phase 2C isolation, and
Office Scripts compiler/performance guards. Materialization remains inactive;
Performance still does not consume the weekly cache.

The weekly-cache activation suite validates the exact two-field authority
transition, idempotent reruns, single-Active-version rule, full materialized
cache/reconciliation preflight, date-neutral consumer freshness, deterministic
stale responses for mapping/catalog/identity/ReportingEnabled changes, Phase 2C
protection, and Office Scripts compatibility. Activation does not cut
Performance over or introduce period selectors.

The weekly Performance suite validates independent complete ISO-week ranges,
blocked incomplete/invalid ranges, same- and different-length comparisons,
accepted Add-ons fixtures, active-cache freshness, additive aggregation before
shares, the four bounded component-block replacements, Reports linkage and
preservation of the Phase 2C numeric/presentation architecture.

The weekly intake publisher/adapter suites validate dynamic Active authority,
all five intake outcomes, deterministic inactive Candidate construction,
idempotent hidden-ledger behavior, bounded Candidate payloads, typed
Power-Automate routing results, Active-cache isolation, and Office Scripts
compiler/performance compatibility. They do not implement a flow or activate a
new cache.

The New-week publication-runtime suite adds the separate inactive Candidate and
activation calls, deterministic Published authority plan, prior-version
rollback retention, retry idempotency, restore-on-write-failure guards, bounded
canonical writes, compact incremental identity evidence, accepted exact-ID
registry carry-forward, and strict isolation
from facts, mapping/configuration, Performance, Reports and Imports. The
read-only publication audit requires exact caller-supplied Active-corpus,
incoming-report and catalog paths and independently compares business rows with
a full rebuild. The frozen accepted 84-week registry is stored as derived
evidence in `fixtures/build-0.3.0-weekly-identity-registry.json`; the W33
85-week publication checkpoint is recorded in
`expected-build-0.3.0-weekly-w33-publication.json`.

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
- Centralized KPI results consumed by Performance/Reports without direct legacy
  CAT calculations.

Tests should be deterministic, explainable, and tied to a reproducible build or workbook checkpoint.
