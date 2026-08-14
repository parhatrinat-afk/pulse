# Build 0.3.0 — Candidate Compact Weekly Cache Checklist

## Repository/test-space validation

- [x] Uses the accepted weekly parser and accepted identity preflight.
- [x] Uses the existing Pulse hierarchical resolver/current Product hierarchy.
- [x] Requires exact date-neutral MappingContentFingerprint and accepted
  preflight fingerprint before candidate construction.
- [x] Keeps MappingAsOfDate and the Phase 2A/2B mapping fingerprint as audit
  metadata without treating a date-only advance as stale weekly content.
- [x] Produces a candidate version only; it cannot overwrite an active version.
- [x] Weekly scope grain is CacheVersion × SourcePeriodKey × RestaurantID.
- [x] Weekly RPG grain is CacheVersion × SourcePeriodKey × RestaurantID ×
  ReportingGroupID.
- [x] Denominator/source scope exists once per Restaurant/week and is absent
  from RPG rows.
- [x] Dense RPG output has nine rows per observed Restaurant/week.
- [x] Row IDs, version, and cache fingerprint are deterministic.
- [x] Filename/locator and source-row order do not affect cache identity.
- [x] Identity Pending remains in denominator and outside RPG numerators.
- [x] Six pending ProductIDs remain unresolved and auditable.
- [x] Nine hierarchy divergences retain current Pulse Product hierarchy.
- [x] RST-0017 and RST-0018 remain separate, reconciled, and Performance-ineligible.
- [x] Mapped RPG numerators reconcile exactly to Mapped scope coverage.
- [x] Enabled plus excluded restaurant scope reconciles to the complete corpus.
- [x] All 84 weekly reconciliations pass for facts, Sales NOK, and Quantity.
- [x] 2025, 2026 W01-W32, and complete-corpus reconciliations pass.
- [x] W31, W01-W32, and W20-W30 comparisons aggregate components before shares.
- [x] The audit is read-only and requires one exact caller-supplied corpus path.
- [x] No source workbook, canonical workbook, Performance model, cache activation,
  period selector, legacy import, Power Automate flow, or Phase 3 surface changes.

## Frozen checkpoint

- Mapping content fingerprint: `MCF-759cc92c4304a913`
- Cache version: `WCV-1a34ad1f46763d9b`
- Cache fingerprint: `WCC-508dd608166cdb6e`
- Period / scope / dense RPG rows: `84 / 1,421 / 12,789`
- Complete candidate rows: `14,295`
- Complete source: `245,632` facts / `484,728,367.25` Sales NOK /
  `2,469,988.09` Quantity
- Validation: `PASS`

## Future materialization gate (not implemented)

- [ ] Connected workbook is exactly `OneDrive/Pulse/Development/Pulse_Current.xlsx`.
- [ ] Live MappingContentFingerprint matches the candidate; MappingAsOfDate and
  the Phase 2A mapping fingerprint are recorded separately for audit.
- [ ] New candidate tables are written completely and remain inactive.
- [ ] Workbook rows/counts/fingerprint/reconciliation match frozen evidence.
- [ ] Phase 2C Interaction QA remains 16/16 PASS and active Performance is unchanged.
