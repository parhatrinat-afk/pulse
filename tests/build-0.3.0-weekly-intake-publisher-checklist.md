# Build 0.3.0 Weekly Intake Publisher Foundation checklist

- [x] Weekly Performance resolves exactly one `Active / Active` cache.
- [x] Zero/multiple active authority is unavailable without guessing.
- [x] Mapping, catalog/identity, validation and ReportingEnabled freshness are required.
- [x] Existing active cache rows are never mutated in place.
- [x] A valid new report creates a complete deterministic inactive Candidate.
- [x] New Candidate additive rows match a full rebuild over the same reports.
- [x] Duplicate content is a no-op.
- [x] Same-period changed content is Correction Review.
- [x] Invalid reports are Rejected without cache change.
- [x] Stale authority/content blocks without cache change.
- [x] Identity Pending remains accepted denominator coverage and is excluded from RPG numerators.
- [x] Intake event IDs, WCV, WCC and row IDs are deterministic.
- [x] One Active plus one prior full rollback version is the bounded retention target.
- [x] All manifests remain historical evidence and no cleanup occurs in this slice.
- [x] Existing 84-week fixture IDs, row counts and reconciliation remain unchanged.
- [x] No Power Automate flow, OneDrive move, Imports redesign, legacy cutover or Phase 3 work is included.

## Live canonical-workbook evidence — 2026-08-16

- [x] `Pulse_Current.xlsx` resolves one `Active / Active / PASS` cache dynamically.
- [x] The installed authority formula contains no frozen `WCV-1a34ad1f46763d9b` dependency.
- [x] Current/Compare remain `2026 W01–W32` and `2025 W01–W32`.
- [x] Weekly Performance QA remains 16/16 PASS.
- [x] Phase 2C Interaction QA remains 16/16 PASS.
- [x] Performance, Reports, Imports and `tblMetricRPGResults` were unchanged by the bounded authority update.
- [x] No intake-ledger table was materialized; the ledger remains a repository contract pending a runtime publisher adapter.
