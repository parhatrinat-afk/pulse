# Build 0.3.0 — Create Reporting Group live checklist

- [ ] Canonical `Pulse_Current.xlsx` is connected and contains the accepted nine active groups.
- [ ] Run **Create Reporting Group** with the form blank; Settings installs and no authority row is added.
- [ ] Enter a neutral QA name and optional description/notes; run the script once.
- [ ] Exactly one `RPG-0010` row is created with `DOM-SALES`, Active `Yes`, SortOrder `100`.
- [ ] Mapping shows the QA business name in browse/Assign-to and zero reverse-membership Products.
- [ ] Mapping Rules remain unchanged.
- [ ] Effective Mapping remains unchanged.
- [ ] Performance Reporting Group selection has ten rows; `RPG-0010` defaults `No` and prior selections remain unchanged.
- [ ] Detail Reporting Group and Sort selectors contain the QA business name.
- [ ] Mapping says `Performance refresh required`; weekly Performance is stale/unavailable rather than using the nine-group cache.
- [ ] Source facts, imports, active/rollback cache rows, legacy 306-row results, and six visible sheets are unchanged.
- [ ] A duplicate-name attempt fails without another row.
- [ ] Remove the controlled QA row using the QA-only restoration procedure, restore the nine-group derived selectors, and recalculate.
- [ ] Canonical workbook finishes with nine groups, no `RPG-0010`, and Mapping/Phase 2C/Weekly Performance QA restored.

## Observed canonical-workbook result — 2026-08-22

- PASS — temporary `RPG-0010` was created exactly once at SortOrder `100`.
- PASS — Mapping and Settings exposed the tenth business name with zero members.
- PASS — Performance selection expanded to ten rows and the new row was `No`.
- PASS — existing nine selections were preserved by stable ReportingGroupID.
- PASS — Mapping Rules / Effective Mapping remained 133 / 1,041.
- PASS — cache staleness was user-visible; no cache version was rebuilt or activated.
- PASS — the temporary authority row and all derived ten-group surfaces were removed.
- PASS — final state restored nine groups, no `RPG-0010`, Mapping 9/9,
  Phase 2C 16/16, Weekly Performance 16/16, and weekly Performance Available.
