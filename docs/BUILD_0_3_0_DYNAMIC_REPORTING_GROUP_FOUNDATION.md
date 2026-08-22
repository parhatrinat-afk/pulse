# Build 0.3.0 — Dynamic Reporting Group Foundation

## Purpose

This slice removes the current weekly runtime assumption that exactly nine
Reporting Groups are active. Reporting Groups remain the single business-owned
semantic catalog; mapping precedence, weekly-cache grains, and KPI-0001 meaning
are unchanged.

The accepted live business state remains nine active groups. Ten-group tests use
repository fixtures only and do not create `RPG-0010` in `Pulse_Current.xlsx`.

## Runtime contract

Current weekly-cache and Weekly Performance consumers require:

- at least one active Reporting Group;
- unique stable `ReportingGroupID` values across the catalog;
- unique finite `SortOrder` values among active groups;
- deterministic active-group order by `SortOrder`, then stable ID;
- a complete dense weekly grain of `weekly scope rows × active groups`.

A group with no mapped products still receives deterministic zero-valued dense
RPG rows. Source facts, Sales NOK, Quantity, coverage states, and scope
denominators are unchanged.

## Weekly Performance

The current Weekly Performance installer sizes its Reporting Group selection,
detail/sort lists, component blocks, numeric helpers, visible matrix, Total, and
Grand Total from the active catalog. Existing selections are restored by stable
ID. A newly eligible group defaults to `No`, so catalog creation does not silently
expand a user's analytical scope.

The formulas remain the accepted Phase 2C contract:

`active weekly cache -> additive Current/Compare components -> numeric helpers -> text facade`

No alternate metric engine is introduced. Reports continues to link to the
Performance detail group and generated period summaries.

## QA separation

The following remain frozen nine-group historical regression evidence:

- `tblMetricRPGResults` and the Phase 2B 306-row checkpoint;
- accepted Phase 2B/Phase 2C fixture documentation;
- the original one-time 84-week cache materialization/activation scripts.

The following are current, count-driven runtime contracts:

- repository weekly-cache construction and validation;
- repository and Office Script new-week publication;
- active weekly-cache validation;
- Weekly Performance helpers, selectors, matrix, and Weekly Performance QA.

This separation preserves rollback/audit history without treating nine groups as
the permanent current-runtime capacity.

## Ten-group fixture

Repository tests add an active zero-member `RPG-0010` fixture and prove:

- dense row count equals scope rows × 10;
- zero numerator for the empty group;
- source totals and reconciliation remain unchanged;
- current nine selections survive by ID and the new group defaults `No`;
- helper/matrix geometry expands deterministically;
- new-week publication produces a complete reconciled ten-group candidate;
- duplicate IDs, duplicate active sort orders, and zero active groups fail.

## Boundary

This slice does not add Settings CRUD, create a Reporting Group, change mapping
semantics, rebuild the live cache to ten groups, redesign Performance, or modify
the fixed historical regression layer.
