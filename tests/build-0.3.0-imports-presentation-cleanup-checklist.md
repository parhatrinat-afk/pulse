# Imports presentation cleanup — live Excel checklist

- [ ] Run `Build_0_3_0_Imports_Presentation_Cleanup.ts` once in canonical
      `OneDrive/Pulse/Development/Pulse_Current.xlsx`.
- [ ] Confirm only Overview, Performance, Reports, Imports, Mapping and Settings
      are visible.
- [ ] Confirm Imports shows latest published `2026 W33`, coverage
      `2025 W01–W52` / `2026 W01–W33`, and `Up to date`.
- [ ] Confirm `Weekly Import Activity` contains exactly four genuine rows:
      W33 duplicate, W33 published, W32 duplicate and W31 duplicate, newest first.
- [ ] Confirm no WCV/WCC/fingerprint/source-period identifiers or long OneDrive
      source identifiers are visible.
- [ ] Confirm legacy `tblImports` remains at `Imports!A4:S6` with unchanged values,
      formulas, schema and stable IDs, while its rows are hidden from normal use.
- [ ] At 100% zoom confirm headers, Processed, Source report, Sales NOK and Message
      are readable with no clipping, spill, `#######`, or formula errors.
- [ ] Confirm Weekly Performance QA and Phase 2C Interaction QA remain 16/16 PASS.
- [ ] Confirm active W33 authority and rollback version remain unchanged internally.
- [ ] Confirm Performance, Reports and Mapping remain unchanged.
- [ ] Rerun the script and confirm the result is idempotent.
