import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const root = new URL("../", import.meta.url);
const script = await read("office-scripts/Build_0_3_0_Release_State_Cleanup.ts");
const readme = await read("README.md");
const changelog = await read("CHANGELOG.md");
const runbook = await read("docs/BUILD_0_3_0_OPERATIONS_RUNBOOK.md");
const cleanupDoc = await read("docs/BUILD_0_3_0_RELEASE_STATE_CLEANUP.md");
const officeManifest = await read("office-scripts/README.md");
const obsoleteUpload = await read("UPLOAD_INSTRUCTIONS.md");
const gitignore = await read(".gitignore");

test("release cleanup clears only transient Mapping action state", () => {
  assert.match(script, /getColumnByName\("Select"\)/);
  assert.match(script, /getRangeBetweenHeaderAndTotal\(\)\.clear\(ExcelScript\.ClearApplyTo\.contents\)/);
  assert.match(script, /getRange\("B15:C15"\)\.clear\(ExcelScript\.ClearApplyTo\.contents\)/);
  assert.match(script, /getRange\("E15:F15"\)\.clear\(ExcelScript\.ClearApplyTo\.contents\)/);
  assert.match(script, /No pending mapping action\./);
  assert.match(script, /fingerprintTable\(rules\) !== rulesBefore/);
  assert.match(script, /fingerprintTable\(effective\) !== effectiveBefore/);
});

test("release cleanup validates and preserves accepted authorities", () => {
  assert.match(script, /dataRowCount\(rules\) !== 133/);
  assert.match(script, /dataRowCount\(effective\) !== 1041/);
  assert.match(script, /dataRowCount\(groups\) !== 9/);
  assert.match(script, /WCV-1b0b195c210da456/);
  assert.match(script, /WCV-1a34ad1f46763d9b/);
  assert.match(script, /qaPassCount\(workbook, "tblMappingQA"\) !== 9/);
  assert.match(script, /qaPassCount\(workbook, "tblPerformanceInteractionQA"\) !== 16/);
  assert.match(script, /qaPassCount\(workbook, "tblWeeklyPerformanceQA"\) !== 16/);
});

test("release cleanup Office Script remains bounded and Excel-web compatible", () => {
  assert.match(script, /^function main\(/m);
  assert.doesNotMatch(script, /\b(?:export|async)\s+function main\(/);
  assert.doesNotMatch(script, /new Map|new Set|\.entries\(|\.keys\(|\.values\(|\.\.\./);
  assert.doesNotMatch(script, /fetch\(|console\./);
  assert.match(script, /getRange\("B15:C15"\)/);
  assert.match(script, /getRange\("E15:F15"\)/);
  assertBalanced(script);
});

test("organisation currency is configured through tblApplicationSettings", () => {
  assert.match(script, /requiredTable\(workbook, "tblApplicationSettings"\)/);
  assert.match(script, /updateTableValue\(settings, "Currency", "NOK"/);
  assert.doesNotMatch(script, /Performance.*setValue\("NOK"\)/s);
});

test("environment is descriptive and weekly manifests stay authoritative", () => {
  assert.match(script, /0\.3\.0-Release-Candidate/);
  assert.match(script, /final Power Automate New-to-Published pilot pending/);
  assert.match(script, /weekly freshness is governed by the weekly cache manifests/);
  assert.match(script, /Legacy import evidence only/);
  assert.doesNotMatch(script, /updateTableValue\(environment, "LatestFinalizedImport"/);
  assert.match(cleanupDoc, /does not release Pulse/);
});

test("operations runbook serializes the single-workbook production flow", () => {
  assert.match(runbook, /Concurrency control enabled/);
  assert.match(runbook, /Degree of parallelism = 1/);
  assert.match(runbook, /exponential retry policy/);
  assert.match(runbook, /four attempts/);
  assert.match(runbook, /source stays in[\s\S]*Incoming reports/);
  assert.match(runbook, /prior Active cache remains authoritative/);
  assert.match(runbook, /Wait six minutes/);
});

test("runbook documents every accepted intake outcome and recovery boundary", () => {
  for (const outcome of ["New", "Duplicate", "Rejected", "Cache Stale", "Correction Review"]) {
    assert.match(runbook, new RegExp(outcome));
  }
  assert.match(runbook, /Do not edit hidden cache, identity, mapping, fingerprint or authority fields/);
  assert.match(runbook, /controlled Power Automate `New -> Published` run/);
});

test("production Office Scripts manifest separates recurring and administrative use", () => {
  for (const name of [
    "Parse Weekly Sales Report", "Process Weekly Intake", "Publish Weekly Intake",
    "Build_0_3_0_Phase1.ts", "Create_Reporting_Group.ts",
    "Build_0_3_0_Weekly_Performance.ts",
  ]) assert.match(officeManifest, new RegExp(name.replaceAll(".", "\\.")));
  assert.match(officeManifest, /Recurring production scripts/);
  assert.match(officeManifest, /Administrative scripts/);
  assert.match(officeManifest, /Installer, build and QA scripts/);
});

test("README and changelog describe the accepted 0.3.0 candidate", () => {
  assert.match(readme, /Current candidate:[\s\S]*Pulse 0\.3\.0/);
  assert.match(readme, /2025 W01–W52 and 2026 W01–W33/);
  assert.match(readme, /Overview[\s\S]*Performance[\s\S]*Reports[\s\S]*Imports[\s\S]*Mapping[\s\S]*Settings/);
  assert.match(readme, /Remaining release gate/);
  assert.match(changelog, /Release-candidate state through 2026 W33/);
  assert.match(changelog, /dynamic\s+positive number of active Reporting Groups/);
  assert.match(changelog, /W33 Published ledger row/);
});

test("W33 blank SourceLocator is documented without invented backfill", () => {
  assert.match(runbook, /W33 `Published` intake event has a blank `SourceLocator`/);
  assert.match(runbook, /proves identical\s+content but not the original OneDrive item/);
  assert.match(runbook, /must not be copied\s+back as historical fact/);
  assert.doesNotMatch(script, /SourceLocator/);
});

test("obsolete upload wrapper cannot be mistaken for current operations", () => {
  assert.match(obsoleteUpload, /^# Obsolete repository-bootstrap instructions/m);
  assert.match(obsoleteUpload, /not.*weekly intake procedure/i);
  assert.match(obsoleteUpload, /BUILD_0_3_0_OPERATIONS_RUNBOOK\.md/);
});

test("macOS metadata is ignored without broad binary ignores", () => {
  assert.equal(gitignore.trim(), ".DS_Store");
});

async function read(path) {
  return readFile(new URL(path, root), "utf8");
}

function assertBalanced(source) {
  for (const [open, close] of [["{", "}"], ["(", ")"], ["[", "]"]]) {
    let depth = 0;
    for (const character of source) {
      if (character === open) depth += 1;
      if (character === close) depth -= 1;
      assert.ok(depth >= 0, `${close} closed before ${open}`);
    }
    assert.equal(depth, 0, `${open}${close} are not balanced`);
  }
}
