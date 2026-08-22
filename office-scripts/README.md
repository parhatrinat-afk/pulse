# Pulse Office Scripts

Only a small subset of this directory is required in Excel for the web during
normal operation. Keep the Excel Automate copies identical to the accepted
repository sources.

## Recurring production scripts

These three scripts must exist in the operator's Excel Automate library:

| Excel script name | Repository source | Runs against | Purpose |
|---|---|---|---|
| Parse Weekly Sales Report | `Parse_Weekly_Sales_Report.ts` | Untouched weekly POS workbook | Returns content-derived period, rows, totals and fingerprints without modifying the source. |
| Process Weekly Intake | `Process_Weekly_Intake.ts` | Canonical Pulse workbook | Validates Active freshness, classifies the event and records non-New outcomes. |
| Publish Weekly Intake | `Publish_Weekly_Intake.ts` | Canonical Pulse workbook | Builds a complete inactive Candidate and activates it only after full validation. |

The production Power Automate flow is serialized at concurrency 1. It must pass
the source OneDrive item identifier as `sourceLocator`, branch on the typed
script results, and archive only a safe Duplicate or Published outcome. See the
[operations runbook](../docs/BUILD_0_3_0_OPERATIONS_RUNBOOK.md).

## Administrative scripts

Install these in Excel Automate when the corresponding administrator workflow
is available:

| User-facing purpose | Repository source | When to run |
|---|---|---|
| Apply Mapping Changes | `Build_0_3_0_Phase1.ts` | After a validated Mapping selection/action. It rebuilds Effective Mapping and marks Performance refresh required when content changes. |
| Create Reporting Group | `Create_Reporting_Group.ts` | From the Settings creation form. It creates one stable active group and does not create Mapping Rules. |
| Refresh Weekly Performance | `Build_0_3_0_Weekly_Performance.ts` | After a structurally changed active group/cache state when the accepted refresh process requires reinstalling the formula surface. Ordinary selectors never require it. |

## Installer, build and QA scripts

The remaining scripts are repository-controlled installation, migration,
presentation or validation assets. Operators do not run them during a normal
weekly intake:

- `Install_Weekly_Identity_Registry.ts`
- `Materialize_Weekly_Compact_Cache.ts`
- `Activate_Weekly_Compact_Cache.ts`
- `Materialize_Weekly_Mapping_Attention.ts`
- `Migrate_Lovable_Mapping.ts`
- `Build_0_3_0_Phase2A.ts`, `Build_0_3_0_Phase2B.ts`,
  `Build_0_3_0_Phase2C.ts`
- `Build_0_3_0_Overview.ts`
- `Build_0_3_0_UX_*.ts`, presentation-cleanup and Imports-cleanup scripts
- `Build_0_3_0_Release_State_Cleanup.ts`
- pre-0.3.0 and Build 0.2.0 build/QA scripts

These scripts remain valuable reproducible evidence. Do not delete them merely
because their workbook surfaces are already installed.

## Safety boundary

- Untouched source reports are never modified by the parser.
- A Candidate is never authoritative.
- Activation retains one prior full rollback version and changes authority only
  after all canonical rows are written and revalidated.
- Mapping and Reporting Group changes make stale weekly results unavailable;
  administrators must rebuild rather than edit hidden fingerprints.
- Workbook binaries and source reports do not belong in this repository.
