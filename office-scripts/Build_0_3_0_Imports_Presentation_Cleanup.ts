/**
 * Pulse Build 0.3.0 — Imports presentation cleanup.
 *
 * Presentation only. The legacy import table remains intact for rollback and
 * lineage. The visible weekly activity is a formula view over the genuine
 * tblWeeklyIntakeLog events; the active cache manifest supplies coverage.
 */
function main(workbook: ExcelScript.Workbook): string {
  const imports = requiredSheet(workbook, "Imports");
  const performance = requiredSheet(workbook, "Performance");
  const reports = requiredSheet(workbook, "Reports");
  const mapping = requiredSheet(workbook, "Mapping");
  const importsTable = requiredTable(workbook, "tblImports");
  const versions = requiredTable(workbook, "tblWeeklyCacheVersions");
  const periods = requiredTable(workbook, "tblWeeklyPeriodManifest");
  const intake = requiredTable(workbook, "tblWeeklyIntakeLog");
  const weeklyQa = requiredTable(workbook, "tblWeeklyPerformanceQA");
  const interactionQa = requiredTable(workbook, "tblPerformanceInteractionQA");

  validateAcceptedStartingPoint(imports, importsTable, versions, periods, intake, weeklyQa, interactionQa);

  const protectedBefore = protectedFingerprint(
    importsTable,
    versions,
    periods,
    intake,
    weeklyQa,
    interactionQa,
    performance.getUsedRange(true),
    reports.getUsedRange(true),
    mapping.getUsedRange(true)
  );
  const sheetCountBefore = workbook.getWorksheets().length;

  installPresentation(imports);
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const protectedAfter = protectedFingerprint(
    importsTable,
    versions,
    periods,
    intake,
    weeklyQa,
    interactionQa,
    performance.getUsedRange(true),
    reports.getUsedRange(true),
    mapping.getUsedRange(true)
  );
  if (protectedAfter !== protectedBefore) {
    throw new Error("PUL-030IM-001: A protected table, value, or formula changed during Imports presentation cleanup.");
  }
  if (workbook.getWorksheets().length !== sheetCountBefore) {
    throw new Error("PUL-030IM-002: Worksheet count changed during Imports presentation cleanup.");
  }

  validatePostconditions(workbook, imports, importsTable, weeklyQa, interactionQa);
  imports.activate();
  imports.getRange("A1").select();

  return (
    "Pulse 0.3.0 Imports presentation cleanup applied. Active weekly coverage and four genuine intake events are visible; " +
    "legacy imports, cache authority, Performance, Reports, Mapping, and both QA gates remain unchanged."
  );
}

const NAVY = "#17233A";
const BLUE = "#4F86F7";
const LIGHT_BLUE = "#EAF2FF";
const PALE_GREEN = "#E7F4EA";
const NEUTRAL = "#F3F6FA";
const WHITE = "#FFFFFF";
const MUTED = "#5B6677";
const OLD_SUBTITLE = "Published imports are retained with certificates, lineage and recovery states.";
const NEW_SUBTITLE = "Weekly sales reports processed by Pulse.";

const LATEST_FORMULA =
  '=LET(v,XLOOKUP(1,(tblWeeklyCacheVersions[CacheStatus]="Active")*(tblWeeklyCacheVersions[ActivationState]="Active"),tblWeeklyCacheVersions[CacheVersion],""),' +
  'y,MAXIFS(tblWeeklyPeriodManifest[ISOYear],tblWeeklyPeriodManifest[CacheVersion],v),' +
  'w,MAXIFS(tblWeeklyPeriodManifest[ISOWeek],tblWeeklyPeriodManifest[CacheVersion],v,tblWeeklyPeriodManifest[ISOYear],y),' +
  'IF(v="","Check required",y&" W"&TEXT(w,"00")))';

const COVERAGE_FORMULA =
  '=LET(v,XLOOKUP(1,(tblWeeklyCacheVersions[CacheStatus]="Active")*(tblWeeklyCacheVersions[ActivationState]="Active"),tblWeeklyCacheVersions[CacheVersion],""),' +
  'ys,SORT(UNIQUE(FILTER(tblWeeklyPeriodManifest[ISOYear],tblWeeklyPeriodManifest[CacheVersion]=v))),' +
  'TEXTJOIN(CHAR(10),TRUE,MAP(ys,LAMBDA(y,y&" W"&TEXT(MINIFS(tblWeeklyPeriodManifest[ISOWeek],tblWeeklyPeriodManifest[CacheVersion],v,tblWeeklyPeriodManifest[ISOYear],y),"00")&"–W"&TEXT(MAXIFS(tblWeeklyPeriodManifest[ISOWeek],tblWeeklyPeriodManifest[CacheVersion],v,tblWeeklyPeriodManifest[ISOYear],y),"00")))))';

const STATUS_FORMULA =
  '=LET(n,SUMPRODUCT((tblWeeklyCacheVersions[CacheStatus]="Active")*(tblWeeklyCacheVersions[ActivationState]="Active")),' +
  'v,XLOOKUP(1,(tblWeeklyCacheVersions[CacheStatus]="Active")*(tblWeeklyCacheVersions[ActivationState]="Active"),tblWeeklyCacheVersions[CacheVersion],""),' +
  'qa,XLOOKUP(v,tblWeeklyCacheVersions[CacheVersion],tblWeeklyCacheVersions[ValidationStatus],""),' +
  'k,XLOOKUP(MAXIFS(tblWeeklyPeriodManifest[PeriodEnd],tblWeeklyPeriodManifest[CacheVersion],v),tblWeeklyPeriodManifest[PeriodEnd],tblWeeklyPeriodManifest[SourcePeriodKey],""),' +
  'IF(AND(n=1,qa="PASS",COUNTIFS(tblWeeklyIntakeLog[SourcePeriodKey],k,tblWeeklyIntakeLog[IntakeStatus],"Published")>0),"Up to date","Check required"))';

const ACTIVITY_FORMULA =
  '=LET(keep,tblWeeklyIntakeLog[IntakeEventID]<>"",' +
  'keys,FILTER(tblWeeklyIntakeLog[SourcePeriodKey],keep),' +
  'states,FILTER(tblWeeklyIntakeLog[IntakeStatus],keep),' +
  'times,FILTER(tblWeeklyIntakeLog[ProcessedAt],keep),' +
  'locators,FILTER(tblWeeklyIntakeLog[SourceLocator],keep),' +
  'rows,FILTER(tblWeeklyIntakeLog[SourceRowCount],keep),' +
  'sales,FILTER(tblWeeklyIntakeLog[SourceSalesNOK],keep),' +
  'stamp,DATE(VALUE(LEFT(times,4)),VALUE(MID(times,6,2)),VALUE(MID(times,9,2)))+TIME(VALUE(MID(times,12,2)),VALUE(MID(times,15,2)),VALUE(MID(times,18,2))),' +
  'period,XLOOKUP(keys,tblWeeklyPeriodManifest[SourcePeriodKey],tblWeeklyPeriodManifest[ISOYear],"")&" W"&TEXT(XLOOKUP(keys,tblWeeklyPeriodManifest[SourcePeriodKey],tblWeeklyPeriodManifest[ISOWeek],""),"00"),' +
  'friendly,SWITCH(states,"Published","Published","Duplicate","Duplicate — no data change","Correction Review","Review required","Rejected","Rejected","Cache Stale","Blocked — cache stale",states),' +
  'report,IF(RIGHT(LOWER(locators),5)=".xlsx",IFERROR(TEXTAFTER(SUBSTITUTE(locators,CHAR(92),"/"),"/",-1),locators),"Source report"),' +
  'message,SWITCH(states,"Published","Weekly report published successfully.","Duplicate","Already processed; no data changed.","Correction Review","A different report already exists for this period.","Rejected","Report could not be processed.","Cache Stale","Processing is temporarily blocked; maintenance required.","See processing details."),' +
  'SORTBY(HSTACK(period,friendly,stamp,report,rows,sales,message),stamp,-1))';

function installPresentation(sheet: ExcelScript.Worksheet): void {
  const subtitle = text(sheet.getRange("A2").getValue());
  if (subtitle !== OLD_SUBTITLE && subtitle !== NEW_SUBTITLE) {
    throw new Error("PUL-030IM-003: Imports subtitle is not an accepted pre/post value.");
  }
  // Row 7 is intentionally left blank between the legacy table and the
  // presentation. Without that buffer Excel auto-expands tblImports when A7 is
  // populated, even though the legacy rows are hidden.
  const destination = sheet.getRange("A8:I53");
  const destinationValues = destination.getValues();
  const alreadyInstalled = text(sheet.getRange("A12").getValue()) === "Weekly Import Activity";
  if (!alreadyInstalled && !matrixIsBlank(destinationValues)) {
    throw new Error("PUL-030IM-004: Imports presentation destination A8:I53 is not empty.");
  }

  destination.clear(ExcelScript.ClearApplyTo.all);
  sheet.getRange("A1:I1").getFormat().getFill().setColor(NAVY);
  sheet.getRange("A1:I1").getFormat().getFont().setColor(WHITE);
  sheet.getRange("A1:I1").getFormat().getFont().setBold(true);
  sheet.getRange("A1:I1").getFormat().getFont().setSize(18);
  sheet.getRange("A1:I1").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  sheet.getRange("A1:I1").getFormat().setRowHeight(34);
  sheet.getRange("A1").setValue("Imports");

  sheet.getRange("A2:I2").getFormat().getFill().setColor(LIGHT_BLUE);
  sheet.getRange("A2:I2").getFormat().getFont().setColor(NAVY);
  sheet.getRange("A2:I2").getFormat().getFont().setSize(11);
  sheet.getRange("A2:I2").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  sheet.getRange("A2:I2").getFormat().setRowHeight(30);
  sheet.getRange("A2").setValue(NEW_SUBTITLE);

  sheet.getRange("A8:C10").getFormat().getFill().setColor(LIGHT_BLUE);
  sheet.getRange("D8:F10").getFormat().getFill().setColor(LIGHT_BLUE);
  sheet.getRange("G8:I10").getFormat().getFill().setColor(PALE_GREEN);
  sheet.getRange("A8:I10").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  sheet.getRange("A8:I10").getFormat().setWrapText(true);
  sheet.getRange("A8").setValue("Latest published");
  sheet.getRange("D8").setValue("Coverage");
  sheet.getRange("G8").setValue("Status");
  sheet.getRange("A8:I8").getFormat().getFont().setColor(MUTED);
  sheet.getRange("A8:I8").getFormat().getFont().setBold(true);
  sheet.getRange("A9").setFormula(LATEST_FORMULA);
  sheet.getRange("D9").setFormula(COVERAGE_FORMULA);
  sheet.getRange("G9").setFormula(STATUS_FORMULA);
  sheet.getRange("A9:I10").getFormat().getFont().setColor(NAVY);
  sheet.getRange("A9:I10").getFormat().getFont().setBold(true);
  sheet.getRange("A9:I10").getFormat().getFont().setSize(12);
  sheet.getRange("8:8").getFormat().setRowHeight(22);
  sheet.getRange("9:10").getFormat().setRowHeight(24);

  sheet.getRange("A12:I12").getFormat().getFill().setColor(BLUE);
  sheet.getRange("A12:I12").getFormat().getFont().setColor(WHITE);
  sheet.getRange("A12:I12").getFormat().getFont().setBold(true);
  sheet.getRange("A12:I12").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  sheet.getRange("A12:I12").getFormat().setRowHeight(26);
  sheet.getRange("A12").setValue("Weekly Import Activity");

  sheet.getRange("A13:G13").setValues([[
    "Period", "Status", "Processed", "Source report", "Rows", "Sales NOK", "Message"
  ]]);
  sheet.getRange("A13:G13").getFormat().getFill().setColor(NAVY);
  sheet.getRange("A13:G13").getFormat().getFont().setColor(WHITE);
  sheet.getRange("A13:G13").getFormat().getFont().setBold(true);
  sheet.getRange("A13:G13").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  sheet.getRange("A13:G13").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  sheet.getRange("A13:G13").getFormat().setRowHeight(28);
  sheet.getRange("A14").setFormula(ACTIVITY_FORMULA);

  sheet.getRange("A14:G53").getFormat().getFill().setColor(WHITE);
  sheet.getRange("A14:G53").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  sheet.getRange("A14:G53").getFormat().setWrapText(true);
  sheet.getRange("A14:B53").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  sheet.getRange("C14:C53").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  sheet.getRange("D14:D53").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  sheet.getRange("E14:F53").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  sheet.getRange("G14:G53").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  sheet.getRange("C14:C53").setNumberFormat("dd.mm.yyyy hh:mm");
  sheet.getRange("E14:E53").setNumberFormat("#,##0");
  sheet.getRange("F14:F53").setNumberFormat('#,##0.00 "NOK"');
  sheet.getRange("14:53").getFormat().setRowHeight(28);

  const widths: number[] = [95, 170, 130, 145, 75, 125, 270, 18, 18];
  for (let index = 0; index < widths.length; index++) {
    sheet.getRangeByIndexes(0, index, 1, 1).getEntireColumn().getFormat().setColumnWidth(widths[index]);
  }

  // Retain the legacy import table and its identifiers exactly; remove it only
  // from the normal operational view.
  sheet.getRange("4:7").setRowHidden(true);
  sheet.getRange("J:S").setColumnHidden(true);
  sheet.setShowGridlines(false);
  sheet.getFreezePanes().unfreeze();
  sheet.getFreezePanes().freezeRows(2);
}

function validateAcceptedStartingPoint(
  imports: ExcelScript.Worksheet,
  importsTable: ExcelScript.Table,
  versions: ExcelScript.Table,
  periods: ExcelScript.Table,
  intake: ExcelScript.Table,
  weeklyQa: ExcelScript.Table,
  interactionQa: ExcelScript.Table
): void {
  if (importsTable.getRange().getAddress() !== "Imports!A4:S6") {
    throw new Error("PUL-030IM-005: Legacy tblImports is not at the accepted range Imports!A4:S6.");
  }
  if (!matrixIsBlank(imports.getRange("A7:S7").getValues())) {
    throw new Error("PUL-030IM-023: Required blank buffer row 7 is not empty.");
  }
  assertHeaders(importsTable, ["ImportID", "PublicationState", "ImportFingerprint"]);
  assertHeaders(versions, ["CacheVersion", "CacheStatus", "ActivationState", "ValidationStatus", "PeriodRowCount"]);
  assertHeaders(periods, ["CacheVersion", "SourcePeriodKey", "PeriodEnd", "ISOYear", "ISOWeek"]);
  assertHeaders(intake, ["IntakeEventID", "SourceLocator", "SourcePeriodKey", "IntakeStatus", "SourceRowCount", "SourceSalesNOK", "ProcessedAt"]);

  const versionRows = tableRows(versions);
  const activeRows = versionRows.filter(row =>
    rowValue(versions, row, "CacheStatus") === "Active" &&
    rowValue(versions, row, "ActivationState") === "Active"
  );
  const rollbackRows = versionRows.filter(row => rowValue(versions, row, "CacheStatus") === "Rollback");
  if (
    activeRows.length !== 1 ||
    rowValue(versions, activeRows[0], "ValidationStatus") !== "PASS" ||
    Number(rowValue(versions, activeRows[0], "PeriodRowCount")) !== 85 ||
    rollbackRows.length !== 1
  ) {
    throw new Error("PUL-030IM-006: Accepted single Active/Active 85-week cache and one rollback version are not present.");
  }
  const activeVersion = rowValue(versions, activeRows[0], "CacheVersion");
  const activePeriods = tableRows(periods).filter(row => rowValue(periods, row, "CacheVersion") === activeVersion);
  if (activePeriods.length !== 85) {
    throw new Error(`PUL-030IM-007: Expected 85 active period rows; found ${activePeriods.length}.`);
  }
  const periodLabels = activePeriods.map(row =>
    `${rowValue(periods, row, "ISOYear")} W${pad2(Number(rowValue(periods, row, "ISOWeek")))}`
  );
  if (periodLabels.indexOf("2025 W01") < 0 || periodLabels.indexOf("2025 W52") < 0 || periodLabels.indexOf("2026 W01") < 0 || periodLabels.indexOf("2026 W33") < 0) {
    throw new Error("PUL-030IM-008: Accepted 2025 W01–W52 and 2026 W01–W33 coverage is not present.");
  }
  if (tableRows(intake).length !== 4) {
    throw new Error(`PUL-030IM-009: Expected four genuine intake events; found ${tableRows(intake).length}.`);
  }
  validateQaTable(weeklyQa, "QA-030WP-", 16);
  validateQaTable(interactionQa, "QA-0302C-", 16);

  const subtitle = text(imports.getRange("A2").getValue());
  if (subtitle !== OLD_SUBTITLE && subtitle !== NEW_SUBTITLE) {
    throw new Error("PUL-030IM-010: Imports subtitle is not an accepted pre/post value.");
  }
}

function validatePostconditions(
  workbook: ExcelScript.Workbook,
  imports: ExcelScript.Worksheet,
  importsTable: ExcelScript.Table,
  weeklyQa: ExcelScript.Table,
  interactionQa: ExcelScript.Table
): void {
  if (
    text(imports.getRange("A1").getValue()) !== "Imports" ||
    text(imports.getRange("A2").getValue()) !== NEW_SUBTITLE ||
    text(imports.getRange("A12").getValue()) !== "Weekly Import Activity"
  ) {
    throw new Error("PUL-030IM-011: Imports presentation labels were not installed.");
  }
  if (importsTable.getRange().getAddress() !== "Imports!A4:S6") {
    throw new Error("PUL-030IM-012: Legacy tblImports moved or resized.");
  }
  if (
    imports.getRange("4:7").getRowHidden() !== true ||
    imports.getRange("J:S").getColumnHidden() !== true ||
    imports.getRange("A8:I10").getFormat().getWrapText() !== true ||
    imports.getRange("A13:G13").getFormat().getRowHeight() < 27.5
  ) {
    throw new Error("PUL-030IM-013: Imports visibility or readability postcondition failed.");
  }
  if (
    text(imports.getRange("A9").getText()) !== "2026 W33" ||
    text(imports.getRange("D9").getText()).indexOf("2025 W01–W52") < 0 ||
    text(imports.getRange("D9").getText()).indexOf("2026 W01–W33") < 0 ||
    text(imports.getRange("G9").getText()) !== "Up to date"
  ) {
    throw new Error("PUL-030IM-014: Summary values do not match the accepted active weekly cache.");
  }
  if (text(imports.getRange("A14").getText()) !== "2026 W33") {
    throw new Error("PUL-030IM-015: Genuine intake activity is not sorted newest first.");
  }
  validateQaTable(weeklyQa, "QA-030WP-", 16);
  validateQaTable(interactionQa, "QA-0302C-", 16);
  const visible = workbook.getWorksheets()
    .filter(sheet => sheet.getVisibility() === ExcelScript.SheetVisibility.visible)
    .map(sheet => sheet.getName());
  const expected = ["Overview", "Performance", "Reports", "Imports", "Mapping", "Settings"];
  if (visible.join("|") !== expected.join("|")) {
    throw new Error(`PUL-030IM-016: Visible sheet architecture changed: ${visible.join(" | ")}.`);
  }
}

function protectedFingerprint(
  imports: ExcelScript.Table,
  versions: ExcelScript.Table,
  periods: ExcelScript.Table,
  intake: ExcelScript.Table,
  weeklyQa: ExcelScript.Table,
  interactionQa: ExcelScript.Table,
  performance: ExcelScript.Range,
  reports: ExcelScript.Range,
  mapping: ExcelScript.Range
): string {
  return [
    tableFingerprint(imports), tableFingerprint(versions), tableFingerprint(periods),
    tableFingerprint(intake), tableFingerprint(weeklyQa), tableFingerprint(interactionQa),
    rangeFingerprint(performance), rangeFingerprint(reports), rangeFingerprint(mapping)
  ].join("||");
}

function tableFingerprint(table: ExcelScript.Table): string {
  return `${table.getName()}@${table.getRange().getAddress()}@${rangeFingerprint(table.getRange())}`;
}

function rangeFingerprint(range: ExcelScript.Range): string {
  return JSON.stringify(range.getValues()) + "::" + JSON.stringify(range.getFormulas());
}

function validateQaTable(table: ExcelScript.Table, prefix: string, expectedCount: number): void {
  const rows = tableRows(table);
  const idIndex = columnIndex(table, "CheckID");
  const resultIndex = columnIndex(table, "Result");
  if (rows.length !== expectedCount) {
    throw new Error(`PUL-030IM-017: ${table.getName()} has ${rows.length} rows, not ${expectedCount}.`);
  }
  for (let index = 0; index < rows.length; index++) {
    if (text(rows[index][idIndex]).indexOf(prefix) !== 0 || text(rows[index][resultIndex]) !== "PASS") {
      throw new Error(`PUL-030IM-018: ${table.getName()} row ${index + 1} is not an accepted PASS result.`);
    }
  }
}

function assertHeaders(table: ExcelScript.Table, required: string[]): void {
  const headers = table.getHeaderRowRange().getValues()[0].map(value => text(value));
  for (let index = 0; index < required.length; index++) {
    if (headers.indexOf(required[index]) < 0) {
      throw new Error(`PUL-030IM-019: ${table.getName()} is missing ${required[index]}.`);
    }
  }
}

function tableRows(table: ExcelScript.Table): (string | number | boolean)[][] {
  return table.getRangeBetweenHeaderAndTotal().getValues();
}

function rowValue(table: ExcelScript.Table, row: (string | number | boolean)[], column: string): string {
  return text(row[columnIndex(table, column)]);
}

function columnIndex(table: ExcelScript.Table, column: string): number {
  const headers = table.getHeaderRowRange().getValues()[0].map(value => text(value));
  const index = headers.indexOf(column);
  if (index < 0) throw new Error(`PUL-030IM-020: ${table.getName()} is missing ${column}.`);
  return index;
}

function matrixIsBlank(values: (string | number | boolean)[][]): boolean {
  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    for (let columnIndex = 0; columnIndex < values[rowIndex].length; columnIndex++) {
      if (text(values[rowIndex][columnIndex]) !== "") return false;
    }
  }
  return true;
}

function requiredSheet(workbook: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
  const sheet = workbook.getWorksheet(name);
  if (!sheet) throw new Error(`PUL-030IM-021: Missing worksheet ${name}.`);
  return sheet;
}

function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error(`PUL-030IM-022: Missing table ${name}.`);
  return table;
}

function pad2(value: number): string {
  return value < 10 ? `0${value}` : `${value}`;
}

function text(value: unknown): string {
  return value === undefined || value === null ? "" : String(value).trim();
}
