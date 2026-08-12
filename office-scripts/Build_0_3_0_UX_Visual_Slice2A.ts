/**
 * Pulse Build 0.3.0 UX — Visual Slice 2A.
 *
 * Prerequisite:
 * - Accepted Phase 2C workbook with QA-0302C-01..16 all PASS.
 * - Accepted UX Information Architecture Slice 1 state.
 *
 * This rerunnable script owns presentation only. It normalizes the six primary
 * sheets, repairs Reports sizing, and replaces the clipped Performance Explain
 * block with concise user-facing copy. Formula, table, mapping, import, and
 * Phase 2C interaction fingerprints must remain unchanged.
 */
function main(workbook: ExcelScript.Workbook): string {
  const worksheetCountBefore = workbook.getWorksheets().length;
  const tableCountBefore = workbook.getTables().length;

  const overview = requiredSheet(workbook, "Overview");
  const performance = requiredSheet(workbook, "Performance");
  const reports = requiredSheet(workbook, "Reports");
  const imports = requiredSheet(workbook, "Imports");
  const mapping = requiredSheet(workbook, "Mapping");
  const settings = requiredSheet(workbook, "Settings");
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
  const restaurantSelection = requiredTable(workbook, "tblPerformanceRestaurantSelection");
  const groupSelection = requiredTable(workbook, "tblPerformanceRPGSelection");
  const importsTable = requiredTable(workbook, "tblImports");
  const mappingMainTable = requiredTable(workbook, "tblMappingMainNodes");
  const mappingSubTable = requiredTable(workbook, "tblMappingSubcategoryNodes");
  const mappingProductTable = requiredTable(workbook, "tblMappingProducts");
  const settingsTable = requiredTable(workbook, "tblApplicationSettings");

  if (worksheetCountBefore !== 47) {
    throw new Error(
      `PUL-030UX2A-001: Expected the accepted 47-sheet checkpoint; found ${worksheetCountBefore}. ` +
      "No visual changes were applied."
    );
  }
  if (supportingSheets.length !== 41) {
    throw new Error("PUL-030UX2A-002: Supporting-sheet contract is incomplete.");
  }
  validateAcceptedPhase2C(interactionQa);
  validateIaState(overview, performance, reports, imports, mapping, settings, supportingSheets);
  validateOverviewNavigation(overview);
  validateReportsLinkage(reports);

  const restaurantCount = restaurantSelection.getRangeBetweenHeaderAndTotal().getRowCount();
  const explainStartRow = 23 + restaurantCount + 3;
  const explainHeader = performance.getRangeByIndexes(explainStartRow, 0, 1, 8);
  const explainBody = performance.getRangeByIndexes(explainStartRow + 1, 0, 8, 8);
  const explainFormulasBefore = explainBody.getFormulas();
  if (text(explainHeader.getCell(0, 0).getValue()) !== "Explain" || hasFormula(explainFormulasBefore)) {
    throw new Error(
      "PUL-030UX2A-003: Performance Explain block is not the accepted presentation-only Phase 2C surface."
    );
  }
  const detailNoteBefore = text(performance.getRange("A8").getValue());
  if (detailNoteBefore && detailNoteBefore !== DETAIL_NOTE) {
    throw new Error("PUL-030UX2A-004: Performance A8 is no longer available for the approved detail note.");
  }

  const performanceFormulaBefore = formulaFingerprint(performance.getUsedRange(true).getFormulas());
  const reportsFormulaBefore = formulaFingerprint(reports.getUsedRange(true).getFormulas());
  const primaryTableContentBefore = protectedTableContentFingerprint(
    importsTable,
    mappingMainTable,
    mappingSubTable,
    mappingProductTable,
    settingsTable,
    restaurantSelection,
    groupSelection,
    interactionQa
  );
  const primaryTableStructureBefore = protectedTableStructureFingerprint(
    importsTable,
    mappingMainTable,
    mappingSubTable,
    mappingProductTable,
    settingsTable,
    restaurantSelection,
    groupSelection,
    interactionQa
  );

  // All checkpoint validation above must succeed before presentation mutation.
  overview.setShowGridlines(false);
  performance.setShowGridlines(false);
  reports.setShowGridlines(false);
  imports.setShowGridlines(false);
  mapping.setShowGridlines(false);
  settings.setShowGridlines(false);

  formatSharedTitle(overview, "H", 36);
  formatSharedTitle(performance, "T", 38);
  formatSharedTitle(reports, "H", 38);
  formatSharedTitle(imports, "S", 42);
  formatSharedTitle(mapping, "N", 46);
  formatSharedTitle(settings, "C", 40);

  formatOverview(overview);
  formatPerformance(performance, restaurantSelection, groupSelection, explainStartRow, explainHeader, explainBody);
  formatReports(reports);
  formatImports(imports);
  formatMapping(mapping);
  formatSettings(settings);

  const primaryAnchors: ExcelScript.Range[] = [
    overview.getRange("A1"),
    performance.getRange("A1"),
    reports.getRange("A1"),
    imports.getRange("A1"),
    mapping.getRange("A1"),
    settings.getRange("A1")
  ];
  const primarySheets: ExcelScript.Worksheet[] = [
    overview, performance, reports, imports, mapping, settings
  ];
  for (let index = 0; index < primarySheets.length; index++) {
    primarySheets[index].activate();
    primaryAnchors[index].select();
  }
  overview.activate();
  primaryAnchors[0].select();

  const worksheetCountAfter = workbook.getWorksheets().length;
  const tableCountAfter = workbook.getTables().length;
  // Ignore formatting-only used-range expansion when checking formula immutability.
  const performanceFormulaAfter = formulaFingerprint(performance.getUsedRange(true).getFormulas());
  const reportsFormulaAfter = formulaFingerprint(reports.getUsedRange(true).getFormulas());
  const primaryTableContentAfter = protectedTableContentFingerprint(
    importsTable,
    mappingMainTable,
    mappingSubTable,
    mappingProductTable,
    settingsTable,
    restaurantSelection,
    groupSelection,
    interactionQa
  );
  const primaryTableStructureAfter = protectedTableStructureFingerprint(
    importsTable,
    mappingMainTable,
    mappingSubTable,
    mappingProductTable,
    settingsTable,
    restaurantSelection,
    groupSelection,
    interactionQa
  );

  if (worksheetCountAfter !== worksheetCountBefore || tableCountAfter !== tableCountBefore) {
    throw new Error("PUL-030UX2A-005: Worksheet or table count changed unexpectedly.");
  }
  if (performanceFormulaAfter !== performanceFormulaBefore || reportsFormulaAfter !== reportsFormulaBefore) {
    throw new Error("PUL-030UX2A-006: Performance or Reports formulas changed during visual formatting.");
  }
  if (primaryTableContentAfter !== primaryTableContentBefore) {
    throw new Error("PUL-030UX2A-007: A protected primary-sheet table value or formula changed.");
  }
  if (primaryTableStructureAfter !== primaryTableStructureBefore) {
    throw new Error("PUL-030UX2A-008: A protected primary-sheet table schema or range changed.");
  }
  validateAcceptedPhase2C(interactionQa);
  validateIaState(overview, performance, reports, imports, mapping, settings, supportingSheets);
  validateOverviewNavigation(overview);
  validateVisualPostconditions(overview, performance, reports, imports, mapping, settings, explainStartRow);

  return (
    "Pulse 0.3.0 UX Visual Slice 2A applied. Reports repaired; Performance Explain compacted; " +
    "six primary sheets normalized with gridlines off. Phase 2C formulas, tables, selections, mappings, and imports are unchanged."
  );
}

const NAVY = "#172033";
const BLUE = "#4F8CFF";
const LIGHT_BLUE = "#EAF2FF";
const GREY = "#EEF1F5";
const PALE_YELLOW = "#FFF4D6";
const WHITE = "#FFFFFF";
const MUTED = "#5B6677";
const DETAIL_NOTE =
  "Detail selection controls the result card and Reports; matrix Reporting Group selection is separate.";

function formatSharedTitle(
  sheet: ExcelScript.Worksheet,
  endColumn: string,
  subtitleHeight: number
): void {
  const title = sheet.getRange(`A1:${endColumn}1`);
  title.getFormat().getFill().setColor(NAVY);
  title.getFormat().getFont().setColor(WHITE);
  title.getFormat().getFont().setBold(true);
  title.getFormat().getFont().setSize(18);
  title.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  title.getFormat().setRowHeight(34);

  const subtitle = sheet.getRange(`A2:${endColumn}2`);
  subtitle.getFormat().getFill().setColor(LIGHT_BLUE);
  subtitle.getFormat().getFont().setColor(NAVY);
  subtitle.getFormat().getFont().setSize(11);
  subtitle.getFormat().setWrapText(true);
  subtitle.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  subtitle.getFormat().setRowHeight(subtitleHeight);
}

function formatOverview(sheet: ExcelScript.Worksheet): void {
  sheet.getRange("A1:H43").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  sheet.getRange("A20:H20").getFormat().setWrapText(true);
  sheet.getRange("20:20").getFormat().setRowHeight(30);
}

function formatPerformance(
  sheet: ExcelScript.Worksheet,
  restaurantSelection: ExcelScript.Table,
  groupSelection: ExcelScript.Table,
  explainStartRow: number,
  explainHeader: ExcelScript.Range,
  explainBody: ExcelScript.Range
): void {
  sheet.getRange("B7").getFormat().getFill().setColor(PALE_YELLOW);
  sheet.getRange("B10").getFormat().getFill().setColor(PALE_YELLOW);
  sheet.getRange("G6").getFormat().getFill().setColor(PALE_YELLOW);
  sheet.getRange("G10").getFormat().getFill().setColor(PALE_YELLOW);
  sheet.getRange("I6").getFormat().getFill().setColor(PALE_YELLOW);
  sheet.getRange("I7").getFormat().getFill().setColor(PALE_YELLOW);
  sheet.getRange("B7").getFormat().getFont().setBold(true);
  sheet.getRange("B10").getFormat().getFont().setBold(true);
  sheet.getRange("G6").getFormat().getFont().setBold(true);
  sheet.getRange("G10").getFormat().getFont().setBold(true);
  sheet.getRange("I6").getFormat().getFont().setBold(true);
  sheet.getRange("I7").getFormat().getFont().setBold(true);

  const restaurantInclude = requiredColumn(restaurantSelection, "Include");
  const groupInclude = requiredColumn(groupSelection, "Include");
  restaurantInclude.getRangeBetweenHeaderAndTotal().getFormat().getFill().setColor(PALE_YELLOW);
  groupInclude.getRangeBetweenHeaderAndTotal().getFormat().getFill().setColor(PALE_YELLOW);
  restaurantInclude.getRangeBetweenHeaderAndTotal().getFormat().getFont().setBold(true);
  groupInclude.getRangeBetweenHeaderAndTotal().getFormat().getFont().setBold(true);
  restaurantInclude.getRangeBetweenHeaderAndTotal().getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  groupInclude.getRangeBetweenHeaderAndTotal().getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  sheet.getRange("G7:G8").getFormat().getFill().setColor(LIGHT_BLUE);
  sheet.getRange("I8").getFormat().getFill().setColor(LIGHT_BLUE);
  sheet.getRange("G16:G19").getFormat().getFill().setColor(LIGHT_BLUE);
  sheet.getRange("A8:D8").getFormat().getFill().setColor(LIGHT_BLUE);
  sheet.getRange("A8").setValue(DETAIL_NOTE);
  sheet.getRange("A8").getFormat().getFont().setColor(MUTED);
  sheet.getRange("A8").getFormat().getFont().setItalic(true);
  sheet.getRange("A8").getFormat().getFont().setSize(10);
  sheet.getRange("A8").getFormat().setWrapText(false);

  sheet.getRange("3:4").getFormat().setRowHeight(8);
  sheet.getRange("11:14").getFormat().setRowHeight(8);
  sheet.getRange("20:20").getFormat().setRowHeight(8);
  sheet.getRange(`${explainStartRow - 1}:${explainStartRow}`).getFormat().setRowHeight(8);

  explainHeader.getFormat().getFill().setColor(BLUE);
  explainHeader.getFormat().getFont().setColor(WHITE);
  explainHeader.getFormat().getFont().setBold(true);
  explainHeader.getFormat().setRowHeight(26);

  sheet.getRangeByIndexes(explainStartRow + 1, 1, 4, 7).unmerge();
  explainBody.setValues([
    ["Metric", "Reporting Group Sales Share", "", "", "", "", "", ""],
    ["Definition", "Selected Reporting Group sales as a share of total sales for the selected restaurant and dataset scope.", "", "", "", "", "", ""],
    ["Total", "Total combines the currently selected Reporting Groups.", "", "", "", "", "", ""],
    ["Comparison", "PP Change compares current share with comparison share. NOK Impact estimates the sales difference at the current sales level.", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", ""]
  ]);
  sheet.getRangeByIndexes(explainStartRow + 1, 1, 1, 7).merge();
  sheet.getRangeByIndexes(explainStartRow + 2, 1, 1, 7).merge();
  sheet.getRangeByIndexes(explainStartRow + 3, 1, 1, 7).merge();
  sheet.getRangeByIndexes(explainStartRow + 4, 1, 1, 7).merge();
  sheet.getRangeByIndexes(explainStartRow + 1, 0, 4, 1).getFormat().getFill().setColor(GREY);
  sheet.getRangeByIndexes(explainStartRow + 1, 0, 4, 1).getFormat().getFont().setBold(true);
  sheet.getRangeByIndexes(explainStartRow + 1, 0, 4, 1).getFormat().getFont().setColor(MUTED);
  sheet.getRangeByIndexes(explainStartRow + 1, 1, 4, 7).getFormat().getFill().setColor(WHITE);
  sheet.getRangeByIndexes(explainStartRow + 1, 1, 4, 7).getFormat().setWrapText(true);
  sheet.getRangeByIndexes(explainStartRow + 1, 0, 4, 8).getFormat()
    .setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  sheet.getRangeByIndexes(explainStartRow + 1, 0, 1, 8).getFormat().setRowHeight(26);
  sheet.getRangeByIndexes(explainStartRow + 2, 0, 1, 8).getFormat().setRowHeight(48);
  sheet.getRangeByIndexes(explainStartRow + 3, 0, 1, 8).getFormat().setRowHeight(30);
  sheet.getRangeByIndexes(explainStartRow + 4, 0, 1, 8).getFormat().setRowHeight(52);
  sheet.getRangeByIndexes(explainStartRow + 5, 0, 4, 8).getFormat().setRowHeight(5);
}

function formatReports(sheet: ExcelScript.Worksheet): void {
  sheet.getRange("A1:H22").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  sheet.getRange("A1:H22").getFormat().setWrapText(false);
  sheet.getRange("A:A").getFormat().setColumnWidth(205);
  sheet.getRange("B:B").getFormat().setColumnWidth(315);
  sheet.getRange("C:D").getFormat().setColumnWidth(90);
  sheet.getRange("E:H").getFormat().setColumnWidth(70);
  sheet.getRange("A2:H2").getFormat().setWrapText(true);
  sheet.getRange("A6:A11").getFormat().setWrapText(true);
  sheet.getRange("B6:B11").getFormat().setWrapText(true);
  sheet.getRange("B10:D11").getFormat().setWrapText(true);
  sheet.getRange("A15:A18").getFormat().setWrapText(true);
  sheet.getRange("B6:B11").getFormat().getFill().setColor(WHITE);
  sheet.getRange("B10:D11").getFormat().getFill().setColor(LIGHT_BLUE);
  sheet.getRange("B15:B18").getFormat().getFill().setColor(WHITE);
  sheet.getRange("B15:B18").getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  sheet.getRange("3:4").getFormat().setRowHeight(8);
  sheet.getRange("5:5").getFormat().setRowHeight(26);
  sheet.getRange("6:9").getFormat().setRowHeight(26);
  sheet.getRange("10:11").getFormat().setRowHeight(42);
  sheet.getRange("12:13").getFormat().setRowHeight(10);
  sheet.getRange("14:14").getFormat().setRowHeight(26);
  sheet.getRange("15:17").getFormat().setRowHeight(27);
  sheet.getRange("18:18").getFormat().setRowHeight(34);
}

function formatImports(sheet: ExcelScript.Worksheet): void {
  sheet.getRange("A1:S6").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  sheet.getRange("A2:S2").getFormat().setWrapText(true);
  sheet.getRange("A4:S4").getFormat().setWrapText(true);
  sheet.getRange("4:4").getFormat().setRowHeight(32);
  sheet.getFreezePanes().freezeRows(4);
}

function formatMapping(sheet: ExcelScript.Worksheet): void {
  sheet.getRange("A2:N2").getFormat().setWrapText(true);
  sheet.getRange("A4:N12").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  sheet.getRange("B5:B12").getFormat().getFill().setColor(PALE_YELLOW);
  sheet.getRange("B5:B12").getFormat().getFont().setBold(true);
  sheet.getRange("C5:C12").getFormat().setWrapText(true);
  sheet.getRange("E5:N7").getFormat().setWrapText(true);
  sheet.getRange("5:7").getFormat().setRowHeight(46);
  sheet.getRange("8:9").getFormat().setRowHeight(36);
  sheet.getRange("10:12").getFormat().setRowHeight(32);
  sheet.getFreezePanes().freezeRows(4);
}

function formatSettings(sheet: ExcelScript.Worksheet): void {
  sheet.getRange("A:A").getFormat().setColumnWidth(210);
  sheet.getRange("B:B").getFormat().setColumnWidth(260);
  sheet.getRange("C:C").getFormat().setColumnWidth(360);
  sheet.getRange("A1:C20").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  sheet.getRange("A2:C2").getFormat().setWrapText(true);
  sheet.getRange("C5:C12").getFormat().setWrapText(true);
  sheet.getRange("B5:B12").getFormat().getFill().setColor(PALE_YELLOW);
  sheet.getRange("B5:B12").getFormat().getFont().setBold(true);
  sheet.getRange("4:4").getFormat().setRowHeight(30);
  sheet.getRange("5:12").getFormat().setRowHeight(28);
  sheet.getFreezePanes().freezeRows(4);
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
    observed[text(rows[rowIndex][checkIdColumn])] = text(rows[rowIndex][resultColumn]);
  }
  if (rows.length !== expectedIds.length) {
    throw new Error(`PUL-030UX2A-009: Expected 16 Phase 2C QA rows; found ${rows.length}.`);
  }
  for (const checkId of expectedIds) {
    if (observed[checkId] !== "PASS") {
      throw new Error(`PUL-030UX2A-010: ${checkId} is ${observed[checkId] || "missing"}, not PASS.`);
    }
  }
}

function validateIaState(
  overview: ExcelScript.Worksheet,
  performance: ExcelScript.Worksheet,
  reports: ExcelScript.Worksheet,
  imports: ExcelScript.Worksheet,
  mapping: ExcelScript.Worksheet,
  settings: ExcelScript.Worksheet,
  supporting: ExcelScript.Worksheet[]
): void {
  if (
    overview.getVisibility() !== ExcelScript.SheetVisibility.visible || overview.getPosition() !== 0 ||
    performance.getVisibility() !== ExcelScript.SheetVisibility.visible || performance.getPosition() !== 1 ||
    reports.getVisibility() !== ExcelScript.SheetVisibility.visible || reports.getPosition() !== 2 ||
    imports.getVisibility() !== ExcelScript.SheetVisibility.visible || imports.getPosition() !== 3 ||
    mapping.getVisibility() !== ExcelScript.SheetVisibility.visible || mapping.getPosition() !== 4 ||
    settings.getVisibility() !== ExcelScript.SheetVisibility.visible || settings.getPosition() !== 5
  ) {
    throw new Error("PUL-030UX2A-011: UX IA Slice 1 primary visibility/order is not present.");
  }
  assertHidden(supporting[0]); assertHidden(supporting[1]); assertHidden(supporting[2]);
  assertHidden(supporting[3]); assertHidden(supporting[4]); assertHidden(supporting[5]);
  assertHidden(supporting[6]); assertHidden(supporting[7]); assertHidden(supporting[8]);
  assertHidden(supporting[9]); assertHidden(supporting[10]); assertHidden(supporting[11]);
  assertHidden(supporting[12]); assertHidden(supporting[13]); assertHidden(supporting[14]);
  assertHidden(supporting[15]); assertHidden(supporting[16]); assertHidden(supporting[17]);
  assertHidden(supporting[18]); assertHidden(supporting[19]); assertHidden(supporting[20]);
  assertHidden(supporting[21]); assertHidden(supporting[22]); assertHidden(supporting[23]);
  assertHidden(supporting[24]); assertHidden(supporting[25]); assertHidden(supporting[26]);
  assertHidden(supporting[27]); assertHidden(supporting[28]); assertHidden(supporting[29]);
  assertHidden(supporting[30]); assertHidden(supporting[31]); assertHidden(supporting[32]);
  assertHidden(supporting[33]); assertHidden(supporting[34]); assertHidden(supporting[35]);
  assertHidden(supporting[36]); assertHidden(supporting[37]); assertHidden(supporting[38]);
  assertHidden(supporting[39]); assertHidden(supporting[40]);
}

function assertHidden(sheet: ExcelScript.Worksheet): void {
  if (sheet.getVisibility() !== ExcelScript.SheetVisibility.hidden) {
    throw new Error("PUL-030UX2A-012: A supporting sheet is not in accepted normal-hidden state.");
  }
}

function validateOverviewNavigation(sheet: ExcelScript.Worksheet): void {
  const performance = sheet.getRange("E8").getHyperlink();
  const reports = sheet.getRange("E9").getHyperlink();
  const imports = sheet.getRange("E10").getHyperlink();
  const mapping = sheet.getRange("E11").getHyperlink();
  const settings = sheet.getRange("E12").getHyperlink();
  if (
    performance.documentReference !== "Performance!A1" ||
    reports.documentReference !== "Reports!A1" ||
    imports.documentReference !== "Imports!A1" ||
    mapping.documentReference !== "Mapping!A1" ||
    settings.documentReference !== "Settings!A1"
  ) {
    throw new Error("PUL-030UX2A-013: Overview navigation does not match accepted IA Slice 1.");
  }
}

function validateReportsLinkage(sheet: ExcelScript.Worksheet): void {
  const context = sheet.getRange("B7:B11").getFormulas();
  const results = sheet.getRange("B15:B18").getFormulas();
  const expectedContext = [
    "=Performance!B7", "=Performance!B10", "=Performance!G10", "=Performance!G18", "=Performance!G19"
  ];
  const expectedResults = [
    "=Performance!B16", "=Performance!B17", "=Performance!B18", "=Performance!B19"
  ];
  for (let index = 0; index < expectedContext.length; index++) {
    if (text(context[index][0]) !== expectedContext[index]) {
      throw new Error("PUL-030UX2A-014: Reports context linkage differs from accepted Phase 2C.");
    }
  }
  for (let index = 0; index < expectedResults.length; index++) {
    if (text(results[index][0]) !== expectedResults[index]) {
      throw new Error("PUL-030UX2A-015: Reports result linkage differs from accepted Phase 2C.");
    }
  }
}

function validateVisualPostconditions(
  overview: ExcelScript.Worksheet,
  performance: ExcelScript.Worksheet,
  reports: ExcelScript.Worksheet,
  imports: ExcelScript.Worksheet,
  mapping: ExcelScript.Worksheet,
  settings: ExcelScript.Worksheet,
  explainStartRow: number
): void {
  if (
    overview.getShowGridlines() || performance.getShowGridlines() || reports.getShowGridlines() ||
    imports.getShowGridlines() || mapping.getShowGridlines() || settings.getShowGridlines()
  ) {
    throw new Error("PUL-030UX2A-016: A primary sheet still shows gridlines.");
  }
  if (reports.getRange("A:A").getFormat().getColumnWidth() < 200 ||
      reports.getRange("B:B").getFormat().getColumnWidth() < 300 ||
      reports.getRange("10:11").getFormat().getRowHeight() < 40 ||
      reports.getRange("18:18").getFormat().getRowHeight() < 30) {
    throw new Error("PUL-030UX2A-017: Reports clipping safeguards were not applied.");
  }
  const explain = performance.getRangeByIndexes(explainStartRow + 1, 0, 4, 2).getValues();
  if (
    text(explain[0][0]) !== "Metric" || text(explain[0][1]) !== "Reporting Group Sales Share" ||
    text(explain[1][0]) !== "Definition" ||
    text(explain[2][0]) !== "Total" ||
    text(explain[3][0]) !== "Comparison"
  ) {
    throw new Error("PUL-030UX2A-018: Compact Performance Explain content is incomplete.");
  }
}

function protectedTableContentFingerprint(
  imports: ExcelScript.Table,
  mappingMain: ExcelScript.Table,
  mappingSub: ExcelScript.Table,
  mappingProduct: ExcelScript.Table,
  settings: ExcelScript.Table,
  restaurantSelection: ExcelScript.Table,
  groupSelection: ExcelScript.Table,
  qa: ExcelScript.Table
): string {
  const state = newHashState();
  updateTableContent(state, "tblImports", imports.getRange().getValues(), imports.getRange().getFormulas());
  updateTableContent(state, "tblMappingMainNodes", mappingMain.getRange().getValues(), mappingMain.getRange().getFormulas());
  updateTableContent(state, "tblMappingSubcategoryNodes", mappingSub.getRange().getValues(), mappingSub.getRange().getFormulas());
  updateTableContent(state, "tblMappingProducts", mappingProduct.getRange().getValues(), mappingProduct.getRange().getFormulas());
  updateTableContent(state, "tblApplicationSettings", settings.getRange().getValues(), settings.getRange().getFormulas());
  updateTableContent(state, "tblPerformanceRestaurantSelection", restaurantSelection.getRange().getValues(), restaurantSelection.getRange().getFormulas());
  updateTableContent(state, "tblPerformanceRPGSelection", groupSelection.getRange().getValues(), groupSelection.getRange().getFormulas());
  updateTableContent(state, "tblPerformanceInteractionQA", qa.getRange().getValues(), qa.getRange().getFormulas());
  return finishHash(state, "UX-CONTENT-");
}

function protectedTableStructureFingerprint(
  imports: ExcelScript.Table,
  mappingMain: ExcelScript.Table,
  mappingSub: ExcelScript.Table,
  mappingProduct: ExcelScript.Table,
  settings: ExcelScript.Table,
  restaurantSelection: ExcelScript.Table,
  groupSelection: ExcelScript.Table,
  qa: ExcelScript.Table
): string {
  const state = newHashState();
  updateTableStructure(state, imports.getName(), imports.getRange().getAddress(), imports.getHeaderRowRange().getValues());
  updateTableStructure(state, mappingMain.getName(), mappingMain.getRange().getAddress(), mappingMain.getHeaderRowRange().getValues());
  updateTableStructure(state, mappingSub.getName(), mappingSub.getRange().getAddress(), mappingSub.getHeaderRowRange().getValues());
  updateTableStructure(state, mappingProduct.getName(), mappingProduct.getRange().getAddress(), mappingProduct.getHeaderRowRange().getValues());
  updateTableStructure(state, settings.getName(), settings.getRange().getAddress(), settings.getHeaderRowRange().getValues());
  updateTableStructure(state, restaurantSelection.getName(), restaurantSelection.getRange().getAddress(), restaurantSelection.getHeaderRowRange().getValues());
  updateTableStructure(state, groupSelection.getName(), groupSelection.getRange().getAddress(), groupSelection.getHeaderRowRange().getValues());
  updateTableStructure(state, qa.getName(), qa.getRange().getAddress(), qa.getHeaderRowRange().getValues());
  return finishHash(state, "UX-STRUCTURE-");
}

function updateTableContent(
  state: HashState,
  name: string,
  values: (string | number | boolean)[][],
  formulas: (string | number | boolean)[][]
): void {
  updateHash(state, `${name}\n`);
  updateHashMatrix(state, values);
  updateHashMatrix(state, formulas);
}

function updateTableStructure(
  state: HashState,
  name: string,
  address: string,
  headers: (string | number | boolean)[][]
): void {
  updateHash(state, `${name}|${address}\n`);
  updateHashMatrix(state, headers);
}

function formulaFingerprint(values: (string | number | boolean)[][]): string {
  const state = newHashState();
  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    for (let columnIndex = 0; columnIndex < values[rowIndex].length; columnIndex++) {
      const value = text(values[rowIndex][columnIndex]);
      if (value.indexOf("=") !== 0) continue;
      updateHash(state, `${rowIndex}:${columnIndex}:${value.length}:${value}|`);
    }
  }
  return finishHash(state, "UX-FORMULA-");
}

function hasFormula(values: (string | number | boolean)[][]): boolean {
  for (const row of values) {
    for (const value of row) {
      if (text(value).indexOf("=") === 0) return true;
    }
  }
  return false;
}

type HashState = { left: number; right: number };

function newHashState(): HashState {
  return { left: 0, right: 0 };
}

function updateHash(state: HashState, value: string): void {
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index);
    state.left = (state.left * 131 + code) % 2147483647;
    state.right = (state.right * 137 + code) % 2147483629;
  }
}

function updateHashMatrix(state: HashState, values: (string | number | boolean)[][]): void {
  for (const row of values) {
    for (const value of row) {
      const normalized = text(value);
      updateHash(state, `${normalized.length}:${normalized}|`);
    }
    updateHash(state, "\n");
  }
}

function finishHash(state: HashState, prefix: string): string {
  return `${prefix}${state.left.toString(16).padStart(8, "0")}${state.right.toString(16).padStart(8, "0")}`;
}

function requiredSheet(workbook: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
  const sheet = workbook.getWorksheet(name);
  if (!sheet) throw new Error(`PUL-030UX2A-019: Required worksheet missing: ${name}.`);
  return sheet;
}

function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error(`PUL-030UX2A-020: Required table missing: ${name}.`);
  return table;
}

function requiredColumn(table: ExcelScript.Table, name: string): ExcelScript.TableColumn {
  const column = table.getColumnByName(name);
  if (!column) throw new Error(`PUL-030UX2A-021: ${table.getName()} is missing column ${name}.`);
  return column;
}

function headerIndex(headers: (string | number | boolean)[], requiredHeader: string): number {
  for (let index = 0; index < headers.length; index++) {
    if (text(headers[index]) === requiredHeader) return index;
  }
  throw new Error(`PUL-030UX2A-022: QA table is missing column ${requiredHeader}.`);
}

function text(value: string | number | boolean): string {
  return value === undefined || value === null ? "" : String(value).trim();
}
