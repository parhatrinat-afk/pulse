/**
 * Pulse Build 0.3.0 UX — Information Architecture Slice 1.
 *
 * Prerequisite:
 * Run this after the accepted Phase 2C script and verify that
 * tblPerformanceInteractionQA contains QA-0302C-01..16 as PASS.
 *
 * This script owns only workbook information architecture: primary-sheet
 * visibility/order, simple Overview hyperlinks, and saved A1 starting views.
 * It does not change analytical formulas, tables, source data, mappings, or
 * Phase 2C interaction behavior.
 */
function main(workbook: ExcelScript.Workbook): string {
  const expectedSheetCount = 47;
  const sheetCountBefore = workbook.getWorksheets().length;
  const tableCountBefore = workbook.getTables().length;

  // Resolve every accepted checkpoint sheet before any workbook mutation.
  // The explicit calls keep worksheet reads outside loops for Excel for web.
  const primarySheets: ExcelScript.Worksheet[] = [
    requiredSheet(workbook, "Overview"),
    requiredSheet(workbook, "Performance"),
    requiredSheet(workbook, "Reports"),
    requiredSheet(workbook, "Imports"),
    requiredSheet(workbook, "Mapping"),
    requiredSheet(workbook, "Settings")
  ];
  const supportingSheets: ExcelScript.Worksheet[] = [
    requiredSheet(workbook, "KPI Registry"),
    requiredSheet(workbook, "Context"),
    requiredSheet(workbook, "Views"),
    requiredSheet(workbook, "Restaurants"),
    requiredSheet(workbook, "Reporting Categories"),
    requiredSheet(workbook, "Products"),
    requiredSheet(workbook, "Source Classifications"),
    requiredSheet(workbook, "Remap Assistant"),
    requiredSheet(workbook, "Remap Rules"),
    requiredSheet(workbook, "Import Exclusions"),
    requiredSheet(workbook, "Expected Coverage"),
    requiredSheet(workbook, "Publication Control"),
    requiredSheet(workbook, "Import Actions"),
    requiredSheet(workbook, "Import Certificates"),
    requiredSheet(workbook, "Domains"),
    requiredSheet(workbook, "Source Systems"),
    requiredSheet(workbook, "Adapters"),
    requiredSheet(workbook, "Adapter Contract"),
    requiredSheet(workbook, "Import Control"),
    requiredSheet(workbook, "Test Run Control"),
    requiredSheet(workbook, "_Raw_2025_Baseline"),
    requiredSheet(workbook, "_Raw_2026_Week31"),
    requiredSheet(workbook, "_Sales_Facts"),
    requiredSheet(workbook, "_Standard_Staging"),
    requiredSheet(workbook, "_Remap_Audit"),
    requiredSheet(workbook, "_Import_Action_Audit"),
    requiredSheet(workbook, "_Build_Log"),
    requiredSheet(workbook, "_Environment"),
    requiredSheet(workbook, "_Lists"),
    requiredSheet(workbook, "Effective Categories"),
    requiredSheet(workbook, "_Metric_Calc"),
    requiredSheet(workbook, "Reporting Groups"),
    requiredSheet(workbook, "Mapping Rules"),
    requiredSheet(workbook, "Effective Mapping"),
    requiredSheet(workbook, "_Mapping_Lists"),
    requiredSheet(workbook, "Mapping QA"),
    requiredSheet(workbook, "Metric Contract"),
    requiredSheet(workbook, "Metric Equivalence"),
    requiredSheet(workbook, "_Metric_RPG_Facts"),
    requiredSheet(workbook, "Metric Migration QA"),
    requiredSheet(workbook, "Metric Results QA")
  ];
  const interactionQa = requiredTable(workbook, "tblPerformanceInteractionQA");

  if (sheetCountBefore !== expectedSheetCount) {
    throw new Error(
      `PUL-030UX-001: Expected the accepted 47-sheet Phase 2C checkpoint; found ${sheetCountBefore} sheets. ` +
      "No information-architecture changes were applied."
    );
  }
  if (primarySheets.length + supportingSheets.length !== expectedSheetCount) {
    throw new Error("PUL-030UX-002: The information-architecture sheet contract is incomplete.");
  }
  validateAcceptedPhase2C(interactionQa);

  // Resolve all target ranges before loops so no worksheet read is repeated.
  const primaryAnchors: ExcelScript.Range[] = [
    primarySheets[0].getRange("A1"),
    primarySheets[1].getRange("A1"),
    primarySheets[2].getRange("A1"),
    primarySheets[3].getRange("A1"),
    primarySheets[4].getRange("A1"),
    primarySheets[5].getRange("A1")
  ];
  const overviewNavigation: ExcelScript.Range[] = [
    primarySheets[0].getRange("E8"),
    primarySheets[0].getRange("E9"),
    primarySheets[0].getRange("E10"),
    primarySheets[0].getRange("E11"),
    primarySheets[0].getRange("E12")
  ];

  // Keep a visible sheet active while normal-hidden support sheets are updated.
  primarySheets[0].setVisibility(ExcelScript.SheetVisibility.visible);
  primarySheets[0].activate();
  for (const sheet of primarySheets) {
    sheet.setVisibility(ExcelScript.SheetVisibility.visible);
  }
  for (const sheet of supportingSheets) {
    sheet.setVisibility(ExcelScript.SheetVisibility.hidden);
  }
  for (let index = 0; index < primarySheets.length; index++) {
    primarySheets[index].setPosition(index);
  }

  setInternalLink(overviewNavigation[0], "Performance", "Open Sales Performance");
  setInternalLink(overviewNavigation[1], "Reports", "Open reporting output");
  setInternalLink(overviewNavigation[2], "Imports", "Open import administration");
  setInternalLink(overviewNavigation[3], "Mapping", "Open mapping administration");
  setInternalLink(overviewNavigation[4], "Settings", "Open workbook settings");

  // Selecting A1 while each sheet is active saves its intended opening view.
  for (let index = 0; index < primarySheets.length; index++) {
    primarySheets[index].activate();
    primaryAnchors[index].select();
  }
  primarySheets[0].activate();
  primaryAnchors[0].select();

  const sheetCountAfter = workbook.getWorksheets().length;
  const tableCountAfter = workbook.getTables().length;
  if (sheetCountAfter !== sheetCountBefore || tableCountAfter !== tableCountBefore) {
    throw new Error(
      "PUL-030UX-003: Worksheet or table count changed unexpectedly while applying information architecture."
    );
  }

  return (
    "Pulse 0.3.0 UX IA Slice 1 applied. Visible tabs: Overview | Performance | Reports | " +
    "Imports | Mapping | Settings. Supporting sheets remain present and normally hidden; primary views start at A1."
  );
}

function validateAcceptedPhase2C(table: ExcelScript.Table): void {
  const headers = table.getHeaderRowRange().getValues()[0];
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  const checkIdColumn = headerIndex(headers, "CheckID");
  const resultColumn = headerIndex(headers, "Result");
  const expectedIds: string[] = [
    "QA-0302C-01", "QA-0302C-02", "QA-0302C-03", "QA-0302C-04",
    "QA-0302C-05", "QA-0302C-06", "QA-0302C-07", "QA-0302C-08",
    "QA-0302C-09", "QA-0302C-10", "QA-0302C-11", "QA-0302C-12",
    "QA-0302C-13", "QA-0302C-14", "QA-0302C-15", "QA-0302C-16"
  ];
  const observed: { [key: string]: string } = {};

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const checkId = text(rows[rowIndex][checkIdColumn]);
    if (checkId) observed[checkId] = text(rows[rowIndex][resultColumn]);
  }
  if (rows.length !== expectedIds.length) {
    throw new Error(
      `PUL-030UX-004: tblPerformanceInteractionQA must contain 16 checks; found ${rows.length}. ` +
      "No information-architecture changes were applied."
    );
  }
  for (const checkId of expectedIds) {
    if (observed[checkId] !== "PASS") {
      throw new Error(
        `PUL-030UX-005: Accepted Phase 2C check ${checkId} is ${observed[checkId] || "missing"}, not PASS. ` +
        "No information-architecture changes were applied."
      );
    }
  }
}

function setInternalLink(target: ExcelScript.Range, sheetName: string, screenTip: string): void {
  target.setHyperlink({
    documentReference: `${sheetName}!A1`,
    screenTip,
    textToDisplay: sheetName
  });
}

function headerIndex(headers: (string | number | boolean)[], requiredHeader: string): number {
  for (let index = 0; index < headers.length; index++) {
    if (text(headers[index]) === requiredHeader) return index;
  }
  throw new Error(`PUL-030UX-006: tblPerformanceInteractionQA is missing column ${requiredHeader}.`);
}

function requiredSheet(workbook: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
  const sheet = workbook.getWorksheet(name);
  if (!sheet) {
    throw new Error(`PUL-030UX-007: Required accepted-checkpoint worksheet is missing: ${name}.`);
  }
  return sheet;
}

function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error(`PUL-030UX-008: Required table is missing: ${name}.`);
  return table;
}

function text(value: string | number | boolean): string {
  return value === undefined || value === null ? "" : String(value).trim();
}
