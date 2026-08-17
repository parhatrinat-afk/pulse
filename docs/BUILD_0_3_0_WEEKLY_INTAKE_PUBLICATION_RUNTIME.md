# Build 0.3.0 — Automated New-week Publication Runtime

## Scope

This bounded slice completes the Pulse-side runtime for an intake already
classified `New`. It does not change parser semantics, Performance formulas,
KPI semantics, mapping configuration, Imports, legacy facts, or the Power
Automate duplicate/correction/rejection routes.

The existing flow stays deliberately thin:

`OneDrive report -> Parse -> Process -> Build Candidate -> Activate Candidate -> Archive`

The existing 20-second post-create delay and six-minute pre-archive delay are
operational safeguards and remain required.

## Accepted identity registry prerequisite

The 84-week identity preflight approved 2 exact Restaurant identities, 19 exact
Source Classification identities, and 193 exact Product identities that are
not part of the user-maintained Pulse catalogs. Incremental publication must
reuse those stable IDs; reallocating only the subset seen in an incoming week
would make identity depend on week order.

`Install_Weekly_Identity_Registry.ts` materializes that frozen evidence once on
hidden `_Weekly_Identity`. Its control row binds the rows to active preflight
`IDP-062c182f23905ae8`, registry fingerprint `WIR-776953cb0144af11`, and exact
counts `2 / 19 / 193`.

The registry is not a second identity resolver. It does not mutate
`tblProducts`, `tblRestaurants`, `tblSourceClassifications`, mappings, or facts.
The exact-key resolver stays authoritative. Missing, colliding, changed, or
stale registry evidence fails before Candidate construction.

## Runtime architecture

`Publish_Weekly_Intake.ts` is one Office Script with two explicit operations.

### 1. Build Candidate

The first call accepts the complete reconciled `Parse_Weekly_Sales_Report`
result. It:

1. resolves the single fresh Active cache dynamically;
2. reads the current Pulse identity/mapping catalogs in batches;
3. applies exact identity and current hierarchical mapping semantics to the
   incoming rows;
4. carries forward the validated Active period/scope/RPG components;
5. adds the incoming period at the accepted dense cache grains;
6. reconciles every period, all five mapping states, and mapped RPG totals;
7. derives deterministic WSC/IDP/WCV/WCC values; and
8. writes a hidden `Candidate / Not Active / PASS` surface in bounded chunks.

The Candidate also carries the complete inactive accepted identity registry,
including genuinely new exact candidates from the incoming week.

No authority fields or canonical analytical rows change in this operation.
An exact rerun reuses the same Candidate; a different partial Candidate fails
closed.

### 2. Activate Candidate

The second call receives the exact Candidate WCV/WCC and SourcePeriodKey. It:

1. revalidates current Active freshness and the complete Candidate;
2. copies the prior Active full version to hidden `_Weekly_Cache_Rollback`;
3. writes the Candidate rows into the existing canonical weekly table objects;
4. updates the accepted identity-control fingerprint;
5. performs the final two-field authority transition, leaving exactly one
   `Active / Active` row and one `Rollback / Not Active` row;
6. recalculates and verifies cache authority, the new period, Weekly Performance
   QA 16/16, Phase 2C QA 16/16, and rollback availability; and
7. records `Published` in `tblWeeklyIntakeLog` and returns `archiveReady=true`.

If any canonical mutation fails, the prior Active cache and identity control
are restored before the call fails. Retrying an already completed activation
returns the same `Published` outcome safely.

The identity registry participates in the same backup/restore path. The cache
authority switch remains the final analytical write.

## Power Automate contract

Keep all current Process branches. Add only the following actions to the `New`
branch:

1. **Run script — Publish Weekly Intake / Build Candidate**
   - `operation`: `Build Candidate`
   - `payloadJson`: JSON containing the complete parser result as
     `parsedReport` (or the parser result object itself).
2. **Condition**: continue only when `status = Candidate Ready`,
   `archiveReady = false`, and the returned Candidate WCV/WCC are nonblank.
3. **Run script — Publish Weekly Intake / Activate Candidate**
   - `operation`: `Activate Candidate`
   - `payloadJson`: Candidate version, Candidate fingerprint, and period from
     the first call.
4. **Condition**: archive only when `status = Published`,
   `cacheChanged = true`, and `archiveReady = true`.
5. Preserve the existing six-minute delay before moving the parsed source.

Any error, stale state, Candidate mismatch, or failed postcondition leaves the
source outside Archive and the prior cache authoritative.

## Identity and mapping boundary

The runtime reproduces the accepted exact-key identity preflight and current
hierarchical resolver contract. Identity Pending remains in scope denominators
and outside RPG numerators. Weekly observations do not rewrite current Pulse
Product hierarchy, mapping rules, Effective Mapping, or Reporting Groups.

The incremental identity fingerprint stores compact deterministic evidence for
the incoming assignments and exact identity candidates alongside the accepted
base-cache identity fingerprint. It does not infer names or mapping targets.

## Retention and workbook surfaces

- `_Weekly_Cache_Candidate`: hidden, inactive, one complete validated Candidate;
- `_Weekly_Cache_Rollback`: hidden, prior full Active version retained;
- `_Weekly_Cache`: existing canonical tables and intake ledger;
- `_Weekly_Identity`: hidden accepted exact identity registry;
- normal six-sheet navigation: unchanged.

The script preserves existing table names so weekly Performance and Reports
continue to resolve the canonical Active version dynamically. It never writes
raw facts, legacy imports, Mapping Rules, Products, Restaurants, Performance,
Reports, or `tblMetricRPGResults`.

## Required live acceptance

Before the first genuine publication:

1. run the idempotent `Install_Weekly_Identity_Registry.ts` once and verify
   `WIR-776953cb0144af11` / `2 / 19 / 193`;
2. paste the committed `Publish_Weekly_Intake.ts` into Excel Automate;
3. run `Build Candidate` with the genuine parsed W33 result;
4. verify Candidate PASS/Not Active and exact source/cache reconciliation;
5. run `Activate Candidate` with the returned WCV/WCC and W33 period;
6. verify 85 unique periods, exactly one Active authority, rollback retained,
   W33 selectable without another script, Reports linkage, Weekly Performance
   QA 16/16, and Phase 2C QA 16/16; and
7. only then allow Power Automate to archive the source.

The canonical workbook is `OneDrive/Pulse/Development/Pulse_Current.xlsx`.
No workbook binary or source report belongs in Git.

## Frozen W33 repository evidence

The untouched report whose internal period is `2026-08-10`–`2026-08-16`
parses as 2,940 rows, NOK 5,636,773.50, Quantity 29,654.35 and
`WSF-641061337dfbfd59`. Filename and folder are not identity inputs.

The accepted 85-week Candidate is:

- `WCV-1b0b195c210da456` / `WCC-26c195956ebc2823`;
- 85 period rows, 1,438 scope rows, and 12,942 dense RPG rows;
- 248,572 source facts, NOK 490,365,140.75, Quantity 2,499,642.44;
- Identity Pending 126 facts / NOK 120,048 / Quantity 999;
- 0 Conflict and 0 Inactive Target; and
- registry `WIR-00fc39ff746cb4d1` with 2 Restaurants, 19 classifications,
  and 196 Products.

An independent full 85-week rebuild produces identical period, scope, and RPG
business rows and identical source totals.
