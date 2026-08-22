# Pulse 0.3.0 operations runbook

This runbook covers routine operation of the Pulse 0.3.0 release candidate.
The canonical development workbook is
`OneDrive/Pulse/Development/Pulse_Current.xlsx`. GitHub remains authoritative
for scripts, tests and documentation.

The release candidate is validated through 2026 W33. The final production
gate is one controlled Power Automate `New -> Published` run. Do not treat the
flow as production-ready until that pilot passes.

## Required production configuration

The `Pulse Weekly Intake` trigger must have **Concurrency control enabled** and
**Degree of parallelism = 1**. Pulse publishes through one canonical workbook;
concurrent flow runs or simultaneous manual edits are not supported.

Keep the existing 20-second delay after file creation. Use the normal
exponential retry policy on Excel `Run script` actions; four attempts are the
accepted minimum. A retry is safe because intake events, Candidates and
activation are idempotent. Do not configure later publication or archive
actions to run after a failed prerequisite.

Transient OneDrive or Excel locks may keep a workbook unavailable for several
minutes. If retries are exhausted, the run fails, the source stays in
`Incoming reports`, and the prior Active cache remains authoritative. Wait
until no intake run or manual workbook edit is active, then resubmit the failed
run. Do not weaken freshness or publication checks to force a retry through.

Microsoft documents trigger concurrency control and the Excel connector's
single-writer/locking limits here:

- <https://learn.microsoft.com/en-us/power-automate/guidance/coding-guidelines/optimize-power-automate-triggers>
- <https://learn.microsoft.com/en-us/connectors/excelonlinebusiness/>

## Normal weekly operation

1. Place one untouched weekly POS workbook in
   `OneDrive/Pulse/Incoming reports`.
2. The flow waits 20 seconds, then runs `Parse Weekly Sales Report` against the
   source workbook. The trigger's stable OneDrive item identifier is passed as
   `sourceLocator`; filenames and folders do not determine period identity.
3. The flow runs `Process Weekly Intake` against `Pulse_Current.xlsx` with
   `operation = Process` and the complete parser result as `payloadJson`.
4. Branch on the returned status.

### New

1. Run `Publish Weekly Intake` with `operation = Build Candidate` and the
   complete parser result.
2. Continue only when status is `Candidate Ready`, `archiveReady = false`, and
   both Candidate WCV and WCC are nonblank.
3. Run `Publish Weekly Intake` with `operation = Activate Candidate` and the
   returned Candidate WCV, WCC and SourcePeriodKey.
4. Archive only when status is `Published`, `cacheChanged = true`, and
   `archiveReady = true`.
5. Wait six minutes, then move the source to
   `OneDrive/Pulse/Archive/Weekly reports`.
6. Verify Imports and Overview show the new latest week. The week is now
   available to Performance selectors; existing Current/Compare selections do
   not change automatically.

The authority switch is the last analytical write. A failed Candidate build
does not change the Active cache. A failed activation restores the prior
Active cache and identity control before returning an error.

### Duplicate

Expected result: `Duplicate`, `cacheChanged = false`, `archiveReady = true`.
Wait six minutes, then archive the source. No cache or Performance result
changes.

### Rejected, Cache Stale or Correction Review

Do not archive the source. The file remains in Incoming and the current Active
cache remains authoritative.

- `Rejected`: correct the source/schema problem outside Pulse and submit a new
  untouched export.
- `Cache Stale`: complete the required mapping/catalog/restaurant-scope cache
  rebuild before retrying.
- `Correction Review`: do not overwrite the published week. Explicit
  supersession remains a later controlled workflow.

## Mapping changes

1. In Mapping, select the intended hierarchy members.
2. Choose one valid action and, when required, a Reporting Group.
3. Run the accepted `Build_0_3_0_Phase1` Automate script.
4. Verify Mapping QA is 9/9 PASS and the action controls return to blank.

A real mapping-content change must show `Performance refresh required`. Rebuild
the complete weekly cache through the accepted identity/resolver/cache process,
validate the Candidate, activate it, then rerun the weekly Performance
installer if its structural surface requires refresh. Do not manipulate hidden
fingerprints or cache tables manually.

## New Reporting Group

1. Enter the business name and optional description/notes in Settings.
2. Run `Create Reporting Group`.
3. Confirm the group appears in Mapping and receives a stable generated ID.

The group has no Mapping Rule automatically, defaults to `No` in the future
Performance selection, and makes the current weekly cache stale. Build and
activate a complete cache version against the new active group set before
Performance becomes authoritative again.

## Recovery checklist

Administrators should inspect only normal product surfaces and flow history:

1. Check the failed Power Automate action and its returned status/message.
2. Check Imports for the intake event and current latest published week.
3. Check Overview data status and Performance availability.
4. Confirm the source is still in Incoming unless the result was safely
   `Duplicate` or `Published`.
5. Confirm no other flow run or manual workbook edit is active, then resubmit
   the failed run when the error was transient.

Do not edit hidden cache, identity, mapping, fingerprint or authority fields.
If a business-wrong week has already been published, stop normal intake and
use the retained rollback evidence with repository-assisted recovery. Do not
invent a manual Correction Review replacement.

## Accepted W33 lineage exception

The authoritative W33 `Published` intake event has a blank `SourceLocator`.
Its period, source file ID, semantic fingerprint, totals, Candidate lineage and
Active cache version remain intact. A later Duplicate event proves identical
content but not the original OneDrive item, so its locator must not be copied
back as historical fact. Future New processing must pass the trigger/source
locator through the parser and intake ledger; publication preserves that
supplied locator.
