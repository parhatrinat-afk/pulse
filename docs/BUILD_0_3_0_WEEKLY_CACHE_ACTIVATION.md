# Build 0.3.0 — Weekly Cache Activation and Freshness Guard

## Scope

This bounded slice activates the exact weekly analytical cache already
materialized and validated in the canonical `Pulse_Current.xlsx`. It does not
add period selectors or change Performance, Reports, imports, facts, mapping,
or any analytical cache row.

The accepted authority is:

- CacheVersion `WCV-1a34ad1f46763d9b`;
- CacheFingerprint `WCC-508dd608166cdb6e`;
- MappingContentFingerprint `MCF-759cc92c4304a913`;
- CatalogContentFingerprint `ICC-5644a77c18a97437`;
- Identity Preflight `IDP-062c182f23905ae8`; and
- Performance restaurant scope `RSC-08df626f217dd94b`.

Activation changes only `CacheStatus` and `ActivationState` in
`tblWeeklyCacheVersions`, from `Candidate` / `Not Active` to `Active` /
`Active`. `ValidationStatus` remains `PASS`. The 84 period rows, 1,421 scope
rows, and 12,789 dense RPG rows are not rewritten.

## Activation preflight

`Activate_Weekly_Compact_Cache.ts` stops before its single bounded write unless
all of the following are true:

- the hidden `_Weekly_Cache` surface and four accepted tables exist;
- there is exactly one row for the accepted CacheVersion and no other Active
  cache authority;
- the exact table schemas, row counts, cache fingerprint and validation status
  match the accepted candidate;
- all cache rows reconcile to 245,632 facts, NOK 484,728,367.25 and Quantity
  2,469,988.09;
- Mapped RPG numerators reconcile to Mapped scope, Identity Pending remains
  120 facts / NOK 114,876 / Quantity 951, and Conflict/Inactive Target remain
  zero;
- live mapping content, catalog content and ReportingEnabled restaurant scope
  match the accepted fingerprints; and
- Phase 2C Interaction QA is still 16/16 PASS.

The script is idempotent: an exact fresh `Active` / `Active` version returns
`Already Active`. A differing or split authority state fails visibly.

## Consumer freshness contract

Future weekly-cache consumers must refuse calculation unless exactly one
version is `Active` / `Active`, its validation is `PASS`, and these live values
match the active version:

1. MappingContentFingerprint;
2. CatalogContentFingerprint and accepted IdentityPreflightFingerprint; and
3. PerformanceRestaurantScopeFingerprint.

This is an availability gate, not an elaborate invalidation subsystem. A
genuine mapping, identity/catalog, or ReportingEnabled scope change returns
`Stale` until a new candidate is built, validated, and activated.

`MappingAsOfDate` and the Phase 2A date-sensitive mapping audit fingerprint are
not freshness inputs for weekly consumers. Advancing the audit date without
changing mapping content therefore does not invalidate the weekly cache. The
existing Phase 2A/2B stale-state contracts remain unchanged.

## Non-regression boundary

The activation script fingerprints Performance, Reports,
`tblMetricRPGResults`, and `tblImports` before and after the authority update.
It also revalidates Phase 2C QA and live semantic freshness. It does not add
period selectors, activate a new Performance calculation path, supersede a
legacy import, change the six-sheet visible workflow, or begin Phase 3.

## Live evidence

Activated successfully on `2026-08-14` in the canonical
`Pulse_Current.xlsx`. The accepted version is the only `Active` / `Active`
weekly cache authority and retains `ValidationStatus = PASS`.

Live preflight and postconditions confirmed:

- exact version, cache, mapping-content, catalog-content, identity-preflight,
  and restaurant-scope fingerprints;
- 84 period rows, 1,421 scope rows, 12,789 dense RPG rows and the unchanged
  hidden `_Weekly_Cache` surface;
- 245,632 source facts, NOK 484,728,367.25 and Quantity 2,469,988.09;
- 226,485 Mapped, 19,027 Unmapped and 120 Identity Pending facts, with zero
  Conflict and Inactive Target;
- Mapped RPG additive components still equal Mapped scope components;
- Phase 2C Interaction QA remains 16/16 PASS; and
- Performance (`3988af7020d27700`), Reports (`35b417d137ae52ed`),
  `tblMetricRPGResults` (`3b467e3929d8f06a`) and `tblImports`
  (`201d0869568c427b`) retained their exact pre-activation fingerprints.

The workbook still contains 48 sheets with only Overview, Performance,
Reports, Imports, Mapping and Settings visible. Performance remains on the
accepted Phase 2C engine; no weekly period selector or cache consumer has been
introduced.
