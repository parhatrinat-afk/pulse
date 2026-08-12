/**
 * Pulse Build 0.3.0 UX — Visual Slice 2B.
 *
 * Prerequisite:
 * - Accepted Phase 2C + UX IA Slice 1 + UX Visual Slice 2A workbook.
 * - QA-0302C-01..16 all PASS.
 *
 * This rerunnable repair owns the final pre-checkpoint presentation cleanup on
 * Performance and Mapping. It does not change workbook values, formulas,
 * tables, validations, selections, calculations, navigation, or sheet
 * visibility/order.
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
      `PUL-030UX2B-001: Expected the accepted 47-sheet workbook; found ${worksheetCountBefore}. ` +
      "No presentation changes were applied."
    );
  }
  validatePrimaryIa(overview, performance, reports, imports, mapping, settings);
  validateAcceptedPhase2C(interactionQa);
  validateSlice2AState(performance, restaurantSelection);

  const restaurantBody = restaurantSelection.getRangeBetweenHeaderAndTotal();
  const groupBody = groupSelection.getRangeBetweenHeaderAndTotal();
  if (
    restaurantSelection.getRange().getAddress() !== "Performance!N5:P21" ||
    groupSelection.getRange().getAddress() !== "Performance!R5:T14"
  ) {
    throw new Error("PUL-030UX2B-002: Performance selection-table ranges differ from the accepted Phase 2C layout.");
  }

  const performanceValuesBefore = matrixFingerprint(performance.getUsedRange(true).getValues(), "UX2B-VALUE-");
  const performanceFormulasBefore = formulaFingerprint(performance.getUsedRange(true).getFormulas());
  const mappingValuesBefore = matrixFingerprint(mapping.getUsedRange(true).getValues(), "UX2B-MAPPING-VALUE-");
  const mappingFormulasBefore = formulaFingerprint(mapping.getUsedRange(true).getFormulas());
  const reportsValuesBefore = matrixFingerprint(reports.getUsedRange(true).getValues(), "UX2B-REPORTS-VALUE-");
  const reportsFormulasBefore = formulaFingerprint(reports.getUsedRange(true).getFormulas());
  const protectedContentBefore = protectedTableContentFingerprint(
    importsTable,
    mappingMainTable,
    mappingSubTable,
    mappingProductTable,
    settingsTable,
    restaurantSelection,
    groupSelection,
    interactionQa
  );
  const protectedStructureBefore = protectedTableStructureFingerprint(
    importsTable,
    mappingMainTable,
    mappingSubTable,
    mappingProductTable,
    settingsTable,
    restaurantSelection,
    groupSelection,
    interactionQa
  );
  const validationBefore = selectionValidationFingerprint(restaurantSelection, groupSelection);

  formatPerformance(performance, restaurantBody, groupBody);
  formatMapping(mapping);

  const worksheetCountAfter = workbook.getWorksheets().length;
  const tableCountAfter = workbook.getTables().length;
  const performanceValuesAfter = matrixFingerprint(performance.getUsedRange(true).getValues(), "UX2B-VALUE-");
  const performanceFormulasAfter = formulaFingerprint(performance.getUsedRange(true).getFormulas());
  const mappingValuesAfter = matrixFingerprint(mapping.getUsedRange(true).getValues(), "UX2B-MAPPING-VALUE-");
  const mappingFormulasAfter = formulaFingerprint(mapping.getUsedRange(true).getFormulas());
  const reportsValuesAfter = matrixFingerprint(reports.getUsedRange(true).getValues(), "UX2B-REPORTS-VALUE-");
  const reportsFormulasAfter = formulaFingerprint(reports.getUsedRange(true).getFormulas());
  const protectedContentAfter = protectedTableContentFingerprint(
    importsTable,
    mappingMainTable,
    mappingSubTable,
    mappingProductTable,
    settingsTable,
    restaurantSelection,
    groupSelection,
    interactionQa
  );
  const protectedStructureAfter = protectedTableStructureFingerprint(
    importsTable,
    mappingMainTable,
    mappingSubTable,
    mappingProductTable,
    settingsTable,
    restaurantSelection,
    groupSelection,
    interactionQa
  );
  const validationAfter = selectionValidationFingerprint(restaurantSelection, groupSelection);

  if (worksheetCountAfter !== worksheetCountBefore || tableCountAfter !== tableCountBefore) {
    throw new Error("PUL-030UX2B-003: Worksheet or table count changed during presentation repair.");
  }
  if (
    performanceValuesAfter !== performanceValuesBefore ||
    mappingValuesAfter !== mappingValuesBefore ||
    reportsValuesAfter !== reportsValuesBefore
  ) {
    throw new Error("PUL-030UX2B-004: A protected Performance, Mapping, or Reports value changed.");
  }
  if (
    performanceFormulasAfter !== performanceFormulasBefore ||
    mappingFormulasAfter !== mappingFormulasBefore ||
    reportsFormulasAfter !== reportsFormulasBefore
  ) {
    throw new Error("PUL-030UX2B-005: A protected Performance, Mapping, or Reports formula changed.");
  }
  if (protectedContentAfter !== protectedContentBefore || protectedStructureAfter !== protectedStructureBefore) {
    throw new Error("PUL-030UX2B-006: A protected table value, formula, schema, or range changed.");
  }
  if (validationAfter !== validationBefore) {
    throw new Error("PUL-030UX2B-007: A Performance selection validation rule changed.");
  }

  validatePrimaryIa(overview, performance, reports, imports, mapping, settings);
  validateAcceptedPhase2C(interactionQa);
  validateSlice2AState(performance, restaurantSelection);
  validateRepairPostconditions(performance, mapping, restaurantBody, groupBody);

  return (
    "Pulse 0.3.0 UX Visual Slice 2B applied. Performance and Mapping presentation compacted for 100% zoom. " +
    "All values, formulas, tables, validations, selections, QA, and sheet state are unchanged."
  );
}

const DETAIL_NOTE =
  "Detail selection controls the result card and Reports; matrix Reporting Group selection is separate.";
const PERFORMANCE_SUBTITLE =
  "Interactive Reporting Group Sales Share — restaurant, Reporting Group, and dataset selections recalculate without Office Scripts.";
const NAVY = "#17233A";
const BLUE = "#4F86F7";
const LIGHT_BLUE = "#EAF2FF";
const NEUTRAL = "#F3F6FA";
const PALE_YELLOW = "#FFF2CC";
const PALE_GREEN = "#E2F0D9";
const WHITE = "#FFFFFF";
const MUTED = "#5B6677";

function formatPerformance(
  performance: ExcelScript.Worksheet,
  restaurantBody: ExcelScript.Range,
  groupBody: ExcelScript.Range
): void {
  const subtitle = performance.getRange("A2:T2");
  subtitle.getFormat().setWrapText(false);
  subtitle.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  subtitle.getFormat().setRowHeight(28);
  subtitle.getFormat().getFill().setColor(LIGHT_BLUE);
  subtitle.getFormat().getFont().setColor(NAVY);

  performance.getRange("3:3").getFormat().setRowHeight(8);
  performance.getRange("4:4").getFormat().setRowHeight(18);
  performance.getRange("5:5").getFormat().setRowHeight(20);
  performance.getRange("6:14").getFormat().setRowHeight(19);
  performance.getRange("15:15").getFormat().setRowHeight(20);
  performance.getRange("16:20").getFormat().setRowHeight(19);
  performance.getRange("21:21").getFormat().setRowHeight(20);
  performance.getRange("22:22").getFormat().setRowHeight(18);
  performance.getRange("23:23").getFormat().setRowHeight(24);
  performance.getRange("24:40").getFormat().setRowHeight(19);
  performance.getRange("41:42").getFormat().setRowHeight(6);
  performance.getRange("43:43").getFormat().setRowHeight(22);
  performance.getRange("44:44").getFormat().setRowHeight(20);
  performance.getRange("45:45").getFormat().setRowHeight(30);
  performance.getRange("46:46").getFormat().setRowHeight(20);
  performance.getRange("47:47").getFormat().setRowHeight(30);

  performance.getRange("A:A").getFormat().setColumnWidth(165);
  performance.getRange("B:B").getFormat().setColumnWidth(105);
  performance.getRange("C:C").getFormat().setColumnWidth(100);
  performance.getRange("D:D").getFormat().setColumnWidth(105);
  performance.getRange("E:E").getFormat().setColumnWidth(110);
  performance.getRange("F:F").getFormat().setColumnWidth(100);
  performance.getRange("G:G").getFormat().setColumnWidth(180);
  performance.getRange("H:H").getFormat().setColumnWidth(85);
  performance.getRange("I:I").getFormat().setColumnWidth(100);
  performance.getRange("J:K").getFormat().setColumnWidth(90);
  performance.getRange("L:M").getFormat().setColumnWidth(18);
  performance.getRange("N:N").getFormat().setColumnWidth(65);
  performance.getRange("O:O").getFormat().setColumnWidth(145);
  performance.getRange("P:P").getFormat().setColumnWidth(82);
  performance.getRange("Q:Q").getFormat().setColumnWidth(18);
  performance.getRange("R:R").getFormat().setColumnWidth(65);
  performance.getRange("S:S").getFormat().setColumnWidth(135);
  performance.getRange("T:T").getFormat().setColumnWidth(110);

  for (const address of ["A5:D5", "F5:J5", "A9:D9", "F9:J9", "A21:L21", "A43:J43", "N4:P4", "R4:T4"]) {
    const range = performance.getRange(address);
    range.getFormat().getFill().setColor(BLUE);
    range.getFormat().getFont().setColor(WHITE);
    range.getFormat().getFont().setBold(true);
    range.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  }
  for (const address of ["A15:D15", "F15:J15", "N5:P5", "R5:T5"]) {
    const range = performance.getRange(address);
    range.getFormat().getFill().setColor(NAVY);
    range.getFormat().getFont().setColor(WHITE);
    range.getFormat().getFont().setBold(true);
    range.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  }

  for (const address of ["A6:A7", "A10:A10", "F6:F8", "F10:F10", "H6:H8", "A16:A19", "F16:F19"]) {
    performance.getRange(address).getFormat().getFill().setColor(NEUTRAL);
  }
  for (const address of ["B6:B6", "G7:G8", "I8:I8", "B16:B19", "G16:J19", "O6:P21", "S6:T14"]) {
    performance.getRange(address).getFormat().getFill().setColor(WHITE);
  }
  for (const address of ["B7:B7", "B10:B10", "G6:G6", "G10:G10", "I6:I7", "N6:N21", "R6:R14"]) {
    performance.getRange(address).getFormat().getFill().setColor(PALE_YELLOW);
  }

  performance.getRange("A8:D8").getFormat().getFill().setColor(LIGHT_BLUE);
  performance.getRange("A8").getFormat().getFont().setColor(MUTED);
  performance.getRange("A23:K23").getFormat().getFill().setColor(NEUTRAL);
  performance.getRange("A23:K23").getFormat().getFont().setColor(NAVY);
  performance.getRange("A23:K23").getFormat().getFont().setBold(true);
  performance.getRange("A40:K40").getFormat().getFill().setColor(LIGHT_BLUE);
  performance.getRange("A40:K40").getFormat().getFont().setBold(true);
  performance.getRange("A15:J19").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  performance.getRange("G16:J19").getFormat().setWrapText(false);
  restaurantBody.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  groupBody.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
}

function formatMapping(mapping: ExcelScript.Worksheet): void {
  const subtitle = mapping.getRange("A2:N2");
  subtitle.getFormat().setWrapText(false);
  subtitle.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  subtitle.getFormat().setRowHeight(28);
  subtitle.getFormat().getFill().setColor(LIGHT_BLUE);
  subtitle.getFormat().getFont().setColor(NAVY);

  mapping.getRange("3:3").getFormat().setRowHeight(8);
  mapping.getRange("4:4").getFormat().setRowHeight(20);
  mapping.getRange("5:8").getFormat().setRowHeight(32);
  mapping.getRange("9:12").getFormat().setRowHeight(22);
  mapping.getRange("13:14").getFormat().setRowHeight(8);
  mapping.getRange("15:15").getFormat().setRowHeight(22);

  for (const address of ["A4:C4", "E4:N4"]) {
    const range = mapping.getRange(address);
    range.getFormat().getFill().setColor(BLUE);
    range.getFormat().getFont().setColor(WHITE);
    range.getFormat().getFont().setBold(true);
    range.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  }
  mapping.getRange("A5:A12").getFormat().getFill().setColor(NEUTRAL);
  mapping.getRange("B5:B12").getFormat().getFill().setColor(PALE_YELLOW);
  mapping.getRange("C5:C12").getFormat().getFill().setColor(WHITE);
  mapping.getRange("E5:N7").getFormat().getFill().setColor(WHITE);
  mapping.getRange("E8:N8").getFormat().getFill().setColor(PALE_GREEN);
  mapping.getRange("C5:C8").getFormat().setWrapText(true);
  mapping.getRange("C9:C12").getFormat().setWrapText(false);
  mapping.getRange("E5:N8").getFormat().setWrapText(false);
  mapping.getRange("A4:N12").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  mapping.getRange("B10").setNumberFormat([["dd.mm.yyyy"]]);
}

function validatePrimaryIa(
  overview: ExcelScript.Worksheet,
  performance: ExcelScript.Worksheet,
  reports: ExcelScript.Worksheet,
  imports: ExcelScript.Worksheet,
  mapping: ExcelScript.Worksheet,
  settings: ExcelScript.Worksheet
): void {
  if (
    overview.getVisibility() !== ExcelScript.SheetVisibility.visible || overview.getPosition() !== 0 ||
    performance.getVisibility() !== ExcelScript.SheetVisibility.visible || performance.getPosition() !== 1 ||
    reports.getVisibility() !== ExcelScript.SheetVisibility.visible || reports.getPosition() !== 2 ||
    imports.getVisibility() !== ExcelScript.SheetVisibility.visible || imports.getPosition() !== 3 ||
    mapping.getVisibility() !== ExcelScript.SheetVisibility.visible || mapping.getPosition() !== 4 ||
    settings.getVisibility() !== ExcelScript.SheetVisibility.visible || settings.getPosition() !== 5
  ) {
    throw new Error("PUL-030UX2B-008: Accepted six-sheet IA order/visibility is not present.");
  }
}

function validateSlice2AState(
  performance: ExcelScript.Worksheet,
  restaurantSelection: ExcelScript.Table
): void {
  if (performance.getShowGridlines()) {
    throw new Error("PUL-030UX2B-009: Performance is not in the accepted gridline-free Slice 2A state.");
  }
  if (text(performance.getRange("A2").getValue()) !== PERFORMANCE_SUBTITLE) {
    throw new Error("PUL-030UX2B-010: Performance subtitle differs from the accepted Slice 2A text.");
  }
  if (text(performance.getRange("A8").getValue()) !== DETAIL_NOTE) {
    throw new Error("PUL-030UX2B-011: Performance detail-selector note differs from accepted Slice 2A.");
  }
  const explainStartRow = 23 + restaurantSelection.getRangeBetweenHeaderAndTotal().getRowCount() + 3;
  const explain = performance.getRangeByIndexes(explainStartRow + 1, 0, 4, 2).getValues();
  if (
    text(explain[0][0]) !== "Metric" || text(explain[0][1]) !== "Reporting Group Sales Share" ||
    text(explain[1][0]) !== "Definition" || text(explain[2][0]) !== "Total" ||
    text(explain[3][0]) !== "Comparison"
  ) {
    throw new Error("PUL-030UX2B-012: Accepted compact Performance Explain block is not present.");
  }
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
  for (let index = 0; index < rows.length; index++) {
    observed[text(rows[index][checkIdColumn])] = text(rows[index][resultColumn]);
  }
  if (rows.length !== 16) {
    throw new Error(`PUL-030UX2B-013: Expected 16 Phase 2C QA rows; found ${rows.length}.`);
  }
  for (const checkId of expectedIds) {
    if (observed[checkId] !== "PASS") {
      throw new Error(`PUL-030UX2B-014: ${checkId} is ${observed[checkId] || "missing"}, not PASS.`);
    }
  }
}

function validateRepairPostconditions(
  performance: ExcelScript.Worksheet,
  mapping: ExcelScript.Worksheet,
  restaurantBody: ExcelScript.Range,
  groupBody: ExcelScript.Range
): void {
  const subtitleFormat = performance.getRange("A2:T2").getFormat();
  if (
    Math.abs(subtitleFormat.getRowHeight() - 28) > 0.5 ||
    subtitleFormat.getWrapText() !== false ||
    subtitleFormat.getVerticalAlignment() !== ExcelScript.VerticalAlignment.center
  ) {
    throw new Error("PUL-030UX2B-015: Performance subtitle sizing/wrapping postcondition failed.");
  }
  if (
    performance.getRange("6:14").getFormat().getRowHeight() < 18.5 ||
    performance.getRange("16:20").getFormat().getRowHeight() < 18.5 ||
    groupBody.getFormat().getRowHeight() < 18.5 ||
    restaurantBody.getFormat().getVerticalAlignment() !== ExcelScript.VerticalAlignment.center ||
    groupBody.getFormat().getVerticalAlignment() !== ExcelScript.VerticalAlignment.center
  ) {
    throw new Error("PUL-030UX2B-016: Performance selection-row height postcondition failed.");
  }
  if (
    performance.getRange("B:B").getFormat().getColumnWidth() > 106 ||
    performance.getRange("N:N").getFormat().getColumnWidth() > 66 ||
    performance.getRange("O:O").getFormat().getColumnWidth() > 146 ||
    performance.getRange("R:R").getFormat().getColumnWidth() > 66 ||
    performance.getRange("T:T").getFormat().getColumnWidth() > 111
  ) {
    throw new Error("PUL-030UX2B-020: Performance compact-column postcondition failed.");
  }
  const mappingSubtitle = mapping.getRange("A2:N2").getFormat();
  if (
    Math.abs(mappingSubtitle.getRowHeight() - 28) > 0.5 ||
    mappingSubtitle.getWrapText() !== false ||
    mapping.getRange("C5:C8").getFormat().getWrapText() !== true ||
    mapping.getRange("E5:N8").getFormat().getWrapText() !== false ||
    mapping.getRange("5:8").getFormat().getRowHeight() < 31.5
  ) {
    throw new Error("PUL-030UX2B-021: Mapping text-layout postcondition failed.");
  }
  if (performance.getRange("G16:J19").getFormat().getWrapText() !== false) {
    throw new Error("PUL-030UX2B-023: Performance comparison/status text still wraps.");
  }
  if (mapping.getRange("B10").getNumberFormat()[0][0] !== "dd.mm.yyyy") {
    throw new Error("PUL-030UX2B-022: Mapping EffectiveFrom display format was not repaired.");
  }
}

function selectionValidationFingerprint(
  restaurantSelection: ExcelScript.Table,
  groupSelection: ExcelScript.Table
): string {
  const restaurantRule = requiredColumn(restaurantSelection, "Include")
    .getRangeBetweenHeaderAndTotal().getDataValidation().getRule();
  const groupRule = requiredColumn(groupSelection, "Include")
    .getRangeBetweenHeaderAndTotal().getDataValidation().getRule();
  return `${JSON.stringify(restaurantRule)}|${JSON.stringify(groupRule)}`;
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
  updateTableContent(state, "tblImports", imports);
  updateTableContent(state, "tblMappingMainNodes", mappingMain);
  updateTableContent(state, "tblMappingSubcategoryNodes", mappingSub);
  updateTableContent(state, "tblMappingProducts", mappingProduct);
  updateTableContent(state, "tblApplicationSettings", settings);
  updateTableContent(state, "tblPerformanceRestaurantSelection", restaurantSelection);
  updateTableContent(state, "tblPerformanceRPGSelection", groupSelection);
  updateTableContent(state, "tblPerformanceInteractionQA", qa);
  return finishHash(state, "UX2B-CONTENT-");
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
  updateTableStructure(state, imports);
  updateTableStructure(state, mappingMain);
  updateTableStructure(state, mappingSub);
  updateTableStructure(state, mappingProduct);
  updateTableStructure(state, settings);
  updateTableStructure(state, restaurantSelection);
  updateTableStructure(state, groupSelection);
  updateTableStructure(state, qa);
  return finishHash(state, "UX2B-STRUCTURE-");
}

function updateTableContent(state: HashState, name: string, table: ExcelScript.Table): void {
  updateHash(state, `${name}\n`);
  updateHashMatrix(state, table.getRange().getValues());
  updateHashMatrix(state, table.getRange().getFormulas());
}

function updateTableStructure(state: HashState, table: ExcelScript.Table): void {
  updateHash(state, `${table.getName()}|${table.getRange().getAddress()}\n`);
  updateHashMatrix(state, table.getHeaderRowRange().getValues());
}

function matrixFingerprint(values: (string | number | boolean)[][], prefix: string): string {
  const state = newHashState();
  updateHashMatrix(state, values);
  return finishHash(state, prefix);
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
  return finishHash(state, "UX2B-FORMULA-");
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
  if (!sheet) throw new Error(`PUL-030UX2B-017: Required worksheet missing: ${name}.`);
  return sheet;
}

function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error(`PUL-030UX2B-018: Required table missing: ${name}.`);
  return table;
}

function requiredColumn(table: ExcelScript.Table, name: string): ExcelScript.TableColumn {
  const column = table.getColumnByName(name);
  if (!column) throw new Error(`PUL-030UX2B-019: ${table.getName()} is missing column ${name}.`);
  return column;
}

function headerIndex(headers: (string | number | boolean)[], requiredHeader: string): number {
  for (let index = 0; index < headers.length; index++) {
    if (text(headers[index]) === requiredHeader) return index;
  }
  throw new Error(`PUL-030UX2B-020: QA table is missing column ${requiredHeader}.`);
}

function text(value: string | number | boolean): string {
  return value === undefined || value === null ? "" : String(value).trim();
}
