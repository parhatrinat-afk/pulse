/**
 * Pulse Build 0.3.0 — bounded release-state cleanup.
 *
 * Clears transient Mapping inputs, sets the validated organisation currency,
 * and labels _Environment as descriptive metadata. It does not write Mapping
 * Rules, Effective Mapping, weekly cache authority, facts or reporting results.
 */
function main(workbook: ExcelScript.Workbook): ReleaseCleanupResult {
  const mapping = requiredSheet(workbook, "Mapping");
  const memberTable = requiredTable(workbook, "tblMappingMemberWorkspace");
  const settings = requiredTable(workbook, "tblApplicationSettings");
  const environment = requiredTable(workbook, "tblEnvironment");
  const rules = requiredTable(workbook, "tblMappingRules");
  const effective = requiredTable(workbook, "tblEffectiveMapping");
  const groups = requiredTable(workbook, "tblReportingGroups");
  const versions = requiredTable(workbook, "tblWeeklyCacheVersions");

  validateWorkbookState(workbook, rules, effective, groups, versions);
  const rulesBefore = fingerprintTable(rules);
  const effectiveBefore = fingerprintTable(effective);

  const selectColumn = memberTable.getColumnByName("Select");
  if (!selectColumn) throw new Error("PUL-030RC-001: Mapping Select column is missing.");
  selectColumn.getRangeBetweenHeaderAndTotal().clear(ExcelScript.ClearApplyTo.contents);
  mapping.getRange("B15:C15").clear(ExcelScript.ClearApplyTo.contents);
  mapping.getRange("E15:F15").clear(ExcelScript.ClearApplyTo.contents);
  mapping.getRange("H15:J16").clear(ExcelScript.ClearApplyTo.contents);
  mapping.getRange("A17").setValue("No pending mapping action.");

  updateTableValue(settings, "Currency", "NOK",
    "Configured organisation currency");
  updateTableValue(environment, "BuildID", "0.3.0-Release-Candidate",
    "Installed 0.3.0 release candidate; final Power Automate New-to-Published pilot pending.");
  updateTableValue(environment, "BuildVersion", "0.3.0",
    "Installed product version; weekly freshness is governed by the weekly cache manifests.");
  updateTableValue(environment, "BuildDate", 46256,
    "Release-state metadata refreshed 2026-08-22; not a weekly data-freshness authority.");
  updateTableNote(environment, "LatestFinalizedImport",
    "Legacy import evidence only; current weekly authority is tblWeeklyCacheVersions and tblWeeklyPeriodManifest.");
  updateTableNote(environment, "LatestFinalizedPeriod",
    "Legacy import evidence only; current weekly coverage is authoritative in tblWeeklyPeriodManifest.");

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  if (selectedCount(memberTable) !== 0 || text(mapping.getRange("B15").getValue()) ||
      text(mapping.getRange("E15").getValue())) {
    throw new Error("PUL-030RC-002: Mapping transient inputs were not cleared.");
  }
  if (tableValue(settings, "Currency") !== "NOK") {
    throw new Error("PUL-030RC-003: Organisation currency is not NOK.");
  }
  if (fingerprintTable(rules) !== rulesBefore || fingerprintTable(effective) !== effectiveBefore) {
    throw new Error("PUL-030RC-004: Mapping authority changed during release cleanup.");
  }
  validateWorkbookState(workbook, rules, effective, groups, versions);

  return {
    status: "PASS",
    selectedMappingMembers: 0,
    currency: "NOK",
    buildId: tableValue(environment, "BuildID"),
    activeCacheVersion: cacheVersion(versions, "Active", "Active"),
    rollbackCacheVersion: cacheVersion(versions, "Rollback", "Not Active"),
    mappingRuleCount: dataRowCount(rules),
    effectiveMappingCount: dataRowCount(effective)
  };
}

interface ReleaseCleanupResult {
  status: string;
  selectedMappingMembers: number;
  currency: string;
  buildId: string;
  activeCacheVersion: string;
  rollbackCacheVersion: string;
  mappingRuleCount: number;
  effectiveMappingCount: number;
}

function validateWorkbookState(workbook: ExcelScript.Workbook, rules: ExcelScript.Table,
  effective: ExcelScript.Table, groups: ExcelScript.Table, versions: ExcelScript.Table): void {
  const visible = workbook.getWorksheets()
    .filter(sheet => sheet.getVisibility() === ExcelScript.SheetVisibility.visible)
    .map(sheet => sheet.getName());
  const expectedVisible = ["Overview", "Performance", "Reports", "Imports", "Mapping", "Settings"];
  if (visible.join("|") !== expectedVisible.join("|")) {
    throw new Error("PUL-030RC-010: Visible-sheet architecture differs.");
  }
  if (dataRowCount(rules) !== 133 || dataRowCount(effective) !== 1041 || dataRowCount(groups) !== 9) {
    throw new Error("PUL-030RC-011: Mapping or Reporting Group authority differs.");
  }
  if (cacheVersion(versions, "Active", "Active") !== "WCV-1b0b195c210da456" ||
      cacheVersion(versions, "Rollback", "Not Active") !== "WCV-1a34ad1f46763d9b") {
    throw new Error("PUL-030RC-012: Active or rollback weekly authority differs.");
  }
  if (qaPassCount(workbook, "tblMappingQA") !== 9 ||
      qaPassCount(workbook, "tblPerformanceInteractionQA") !== 16 ||
      qaPassCount(workbook, "tblWeeklyPerformanceQA") !== 16) {
    throw new Error("PUL-030RC-013: Required workbook QA is not fully passing.");
  }
}

function updateTableValue(table: ExcelScript.Table, key: string,
  value: string | number, note: string): void {
  const row = tableRowIndex(table, key);
  const body = table.getRangeBetweenHeaderAndTotal();
  body.getCell(row, 1).setValue(value);
  body.getCell(row, 2).setValue(note);
}

function updateTableNote(table: ExcelScript.Table, key: string, note: string): void {
  const row = tableRowIndex(table, key);
  table.getRangeBetweenHeaderAndTotal().getCell(row, 2).setValue(note);
}

function tableValue(table: ExcelScript.Table, key: string): string {
  const body = table.getRangeBetweenHeaderAndTotal();
  return text(body.getCell(tableRowIndex(table, key), 1).getValue());
}

function tableRowIndex(table: ExcelScript.Table, key: string): number {
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  for (let index = 0; index < values.length; index += 1) {
    if (text(values[index][0]) === key) return index;
  }
  throw new Error(`PUL-030RC-020: ${key} is missing from ${table.getName()}.`);
}

function selectedCount(table: ExcelScript.Table): number {
  const column = table.getColumnByName("Select");
  if (!column) return -1;
  const values = column.getRangeBetweenHeaderAndTotal().getValues();
  let count = 0;
  for (let index = 0; index < values.length; index += 1) {
    if (text(values[index][0]) === "Yes") count += 1;
  }
  return count;
}

function cacheVersion(table: ExcelScript.Table, status: string, authority: string): string {
  const headers = table.getHeaderRowRange().getTexts()[0];
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  const versionIndex = headers.indexOf("CacheVersion");
  const statusIndex = headers.indexOf("CacheStatus");
  const authorityIndex = headers.indexOf("ActivationState");
  const matches: string[] = [];
  for (let index = 0; index < values.length; index += 1) {
    if (text(values[index][statusIndex]) === status &&
        text(values[index][authorityIndex]) === authority) {
      matches.push(text(values[index][versionIndex]));
    }
  }
  return matches.length === 1 ? matches[0] : "";
}

function qaPassCount(workbook: ExcelScript.Workbook, tableName: string): number {
  const table = requiredTable(workbook, tableName);
  const headers = table.getHeaderRowRange().getTexts()[0];
  let statusIndex = headers.indexOf("Status");
  if (statusIndex < 0) statusIndex = headers.indexOf("Result");
  if (statusIndex < 0) throw new Error(`PUL-030RC-021: ${tableName} has no Status/Result column.`);
  const values = table.getRangeBetweenHeaderAndTotal().getTexts();
  let pass = 0;
  for (let index = 0; index < values.length; index += 1) {
    if (values[index][statusIndex] === "PASS") pass += 1;
  }
  return pass;
}

function dataRowCount(table: ExcelScript.Table): number {
  return table.getRangeBetweenHeaderAndTotal().getRowCount();
}

function fingerprintTable(table: ExcelScript.Table): string {
  const values = table.getRange().getValues();
  let left = 0;
  let right = 0;
  for (let row = 0; row < values.length; row += 1) {
    for (let column = 0; column < values[row].length; column += 1) {
      const value = `${text(values[row][column])}|`;
      for (let index = 0; index < value.length; index += 1) {
        const code = value.charCodeAt(index);
        left = (left * 131 + code) % 2147483647;
        right = (right * 137 + code + index) % 2147483629;
      }
    }
  }
  return `${left}:${right}`;
}

function requiredSheet(workbook: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
  const sheet = workbook.getWorksheet(name);
  if (!sheet) throw new Error(`PUL-030RC-030: Missing worksheet ${name}.`);
  return sheet;
}

function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error(`PUL-030RC-031: Missing table ${name}.`);
  return table;
}

function text(value: unknown): string {
  return value === null || value === undefined ? "" : String(value).trim();
}
