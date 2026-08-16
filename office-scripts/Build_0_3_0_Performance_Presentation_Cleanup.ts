/**
 * Pulse Build 0.3.0 — Performance presentation cleanup.
 *
 * Prerequisite:
 * - Accepted active weekly-cache Performance cutover.
 * - Weekly Performance QA and Phase 2C Interaction QA are both 16/16 PASS.
 *
 * This rerunnable presentation-only slice moves the two existing native
 * selection tables below Explain, balances the existing matrix columns, and
 * aligns visible labels/results. It does not rebuild tables or change metric,
 * cache, selection, mapping, Reports, or rollback semantics.
 */
function main(workbook: ExcelScript.Workbook): string {
  const performance = requiredSheet(workbook, "Performance");
  const reports = requiredSheet(workbook, "Reports");
  const restaurantSelection = requiredTable(workbook, "tblPerformanceRestaurantSelection");
  const groupSelection = requiredTable(workbook, "tblPerformanceRPGSelection");
  const interactionQa = requiredTable(workbook, "tblPerformanceInteractionQA");
  const weeklyQa = requiredTable(workbook, "tblWeeklyPerformanceQA");
  const metricResults = requiredTable(workbook, "tblMetricRPGResults");
  const imports = requiredTable(workbook, "tblImports");

  const worksheetCountBefore = workbook.getWorksheets().length;
  const tableCountBefore = workbook.getTables().length;
  if (worksheetCountBefore !== 48) {
    throw new Error(`PUL-030PC-001: Expected the accepted 48-sheet workbook; found ${worksheetCountBefore}.`);
  }
  validatePrimaryIa(workbook);
  validateQa(interactionQa, "QA-0302C-", "PUL-030PC-002");
  validateQa(weeklyQa, "QA-030WP-", "PUL-030PC-003");
  validateAcceptedLayout(performance, restaurantSelection, groupSelection);

  const restaurantName = restaurantSelection.getName();
  const groupName = groupSelection.getName();
  const restaurantContentBefore = tableContentFingerprint(restaurantSelection, "PC-RST-");
  const groupContentBefore = tableContentFingerprint(groupSelection, "PC-RPG-");
  const selectionValidationBefore = selectionValidationFingerprint(restaurantSelection, groupSelection);
  const matrixBefore = rangeFingerprint(performance.getRange("A23:K40"), "PC-MATRIX-");
  const topBefore = rangeFingerprint(performance.getRange("A1:K44"), "PC-TOP-");
  const lowerExplainBefore = rangeFingerprint(performance.getRange("A46:K49"), "PC-EXPLAIN-");
  const performanceFormulasBefore = formulaFingerprint(performance.getUsedRange(true).getFormulas());
  const reportsBefore = rangeFingerprint(reports.getUsedRange(true), "PC-REPORTS-");
  const metricResultsBefore = tableContentFingerprint(metricResults, "PC-METRIC-");
  const importsBefore = tableContentFingerprint(imports, "PC-IMPORTS-");
  const currentCompareBefore = rangeFingerprint(performance.getRange("B10:B12"), "PC-CURRENT-") +
    rangeFingerprint(performance.getRange("G10:G12"), "PC-COMPARE-");

  moveSelectionTables(performance, restaurantSelection, groupSelection);
  updateExplainCopy(performance);
  formatPerformance(performance, restaurantSelection, groupSelection);

  if (workbook.getWorksheets().length !== worksheetCountBefore || workbook.getTables().length !== tableCountBefore) {
    throw new Error("PUL-030PC-004: Worksheet or table count changed during presentation cleanup.");
  }
  if (restaurantSelection.getName() !== restaurantName || groupSelection.getName() !== groupName) {
    throw new Error("PUL-030PC-005: A selection table identity changed.");
  }
  if (
    tableContentFingerprint(restaurantSelection, "PC-RST-") !== restaurantContentBefore ||
    tableContentFingerprint(groupSelection, "PC-RPG-") !== groupContentBefore
  ) {
    throw new Error("PUL-030PC-006: Selection values, stable IDs, or formulas changed.");
  }
  if (selectionValidationFingerprint(restaurantSelection, groupSelection) !== selectionValidationBefore) {
    throw new Error("PUL-030PC-007: A Yes/No selection validation rule changed.");
  }
  if (
    rangeFingerprint(performance.getRange("A23:K40"), "PC-MATRIX-") !== matrixBefore ||
    rangeFingerprint(performance.getRange("A1:K44"), "PC-TOP-") !== topBefore ||
    rangeFingerprint(performance.getRange("A46:K49"), "PC-EXPLAIN-") !== lowerExplainBefore
  ) {
    throw new Error("PUL-030PC-008: A protected Performance result or control value changed.");
  }
  if (formulaFingerprint(performance.getUsedRange(true).getFormulas()) !== performanceFormulasBefore) {
    throw new Error("PUL-030PC-009: A Performance formula changed.");
  }
  if (
    rangeFingerprint(reports.getUsedRange(true), "PC-REPORTS-") !== reportsBefore ||
    tableContentFingerprint(metricResults, "PC-METRIC-") !== metricResultsBefore ||
    tableContentFingerprint(imports, "PC-IMPORTS-") !== importsBefore
  ) {
    throw new Error("PUL-030PC-010: Reports, rollback metric results, or Imports changed.");
  }
  if (
    rangeFingerprint(performance.getRange("B10:B12"), "PC-CURRENT-") +
    rangeFingerprint(performance.getRange("G10:G12"), "PC-COMPARE-") !== currentCompareBefore
  ) {
    throw new Error("PUL-030PC-011: A Current/Compare selector value or formula changed.");
  }

  validatePrimaryIa(workbook);
  validateQa(interactionQa, "QA-0302C-", "PUL-030PC-012");
  validateQa(weeklyQa, "QA-030WP-", "PUL-030PC-013");
  validatePostconditions(performance, restaurantSelection, groupSelection);

  return (
    "Pulse 0.3.0 Performance presentation cleanup applied. Native selection tables now sit below Explain; " +
    "matrix widths and alignment are balanced. Weekly and Phase 2C QA remain 16/16 PASS."
  );
}

const OLD_DEFINITION =
  "Selected Reporting Group sales as a share of total sales for the selected restaurant and dataset scope.";
const NEW_DEFINITION =
  "Selected Reporting Group sales as a share of total sales for the selected restaurants and period.";
const OLD_RESTAURANT_RANGE = "Performance!N5:P21";
const OLD_GROUP_RANGE = "Performance!R5:T14";
const NEW_RESTAURANT_RANGE = "Performance!B51:D67";
const NEW_GROUP_RANGE = "Performance!F51:H60";
const NAVY = "#17233A";
const BLUE = "#4F86F7";
const NEUTRAL = "#F3F6FA";
const WHITE = "#FFFFFF";

function moveSelectionTables(
  performance: ExcelScript.Worksheet,
  restaurantSelection: ExcelScript.Table,
  groupSelection: ExcelScript.Table
): void {
  const restaurantAddress = restaurantSelection.getRange().getAddress();
  const groupAddress = groupSelection.getRange().getAddress();
  if (restaurantAddress === NEW_RESTAURANT_RANGE && groupAddress === NEW_GROUP_RANGE) return;
  if (restaurantAddress !== OLD_RESTAURANT_RANGE || groupAddress !== OLD_GROUP_RANGE) {
    throw new Error(
      `PUL-030PC-014: Selection tables are neither in the accepted source nor destination layout: ` +
      `${restaurantAddress}; ${groupAddress}.`
    );
  }
  if (!rangeIsBlank(performance.getRange("B50:H67"))) {
    throw new Error("PUL-030PC-015: The below-Explain destination B50:H67 is not empty.");
  }

  // A single native cut-style move preserves both existing Table objects,
  // their relative gap, structured references, styles, and validation rules.
  performance.getRange("N4:T21").moveTo(performance.getRange("B50:H67"));

  if (
    restaurantSelection.getRange().getAddress() !== NEW_RESTAURANT_RANGE ||
    groupSelection.getRange().getAddress() !== NEW_GROUP_RANGE
  ) {
    throw new Error("PUL-030PC-016: Excel did not relocate both native selection tables as expected.");
  }
}

function updateExplainCopy(performance: ExcelScript.Worksheet): void {
  const definition = text(performance.getRange("B45").getValue());
  if (definition !== OLD_DEFINITION && definition !== NEW_DEFINITION) {
    throw new Error("PUL-030PC-017: Performance Explain definition differs from the accepted text.");
  }
  performance.getRange("B45").setValue(NEW_DEFINITION);
}

function formatPerformance(
  performance: ExcelScript.Worksheet,
  restaurantSelection: ExcelScript.Table,
  groupSelection: ExcelScript.Table
): void {
  const columnWidths: number[] = [165, 98, 98, 102, 108, 100, 105, 110, 100, 95, 95];
  for (let index = 0; index < columnWidths.length; index++) {
    performance.getRangeByIndexes(0, index, 1, 1).getEntireColumn().getFormat().setColumnWidth(columnWidths[index]);
  }

  performance.getRange("A23:A40").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  performance.getRange("B23:K40").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  performance.getRange("A23:K23").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  performance.getRange("B23:K23").getFormat().setWrapText(true);
  performance.getRange("23:23").getFormat().setRowHeight(30);

  performance.getRange("G6:G8").getFormat().setWrapText(true);
  performance.getRange("G6:G8").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  performance.getRange("6:8").getFormat().setRowHeight(28);

  for (const address of ["A10:A12", "F10:F12"]) {
    const labels = performance.getRange(address);
    labels.getFormat().getFill().setColor(NEUTRAL);
    labels.getFormat().getFont().setColor(NAVY);
    labels.getFormat().getFont().setBold(true);
    labels.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    labels.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  }

  for (const address of ["B50:D50", "F50:H50"]) {
    const title = performance.getRange(address);
    title.getFormat().getFill().setColor(BLUE);
    title.getFormat().getFont().setColor(WHITE);
    title.getFormat().getFont().setBold(true);
    title.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  }
  performance.getRange("50:50").getFormat().setRowHeight(22);
  performance.getRange("51:51").getFormat().setRowHeight(32);
  performance.getRange("52:67").getFormat().setRowHeight(28);

  const restaurantRange = restaurantSelection.getRange();
  const groupRange = groupSelection.getRange();
  restaurantRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  groupRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  restaurantSelection.getHeaderRowRange().getFormat().setWrapText(true);
  groupSelection.getHeaderRowRange().getFormat().setWrapText(true);
  requiredColumn(restaurantSelection, "Include").getRangeBetweenHeaderAndTotal()
    .getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  requiredColumn(groupSelection, "Include").getRangeBetweenHeaderAndTotal()
    .getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  requiredColumn(restaurantSelection, "Restaurant").getRangeBetweenHeaderAndTotal().getFormat().setWrapText(true);
  requiredColumn(groupSelection, "Reporting Group").getRangeBetweenHeaderAndTotal().getFormat().setWrapText(true);
}

function validateAcceptedLayout(
  performance: ExcelScript.Worksheet,
  restaurantSelection: ExcelScript.Table,
  groupSelection: ExcelScript.Table
): void {
  const restaurantAddress = restaurantSelection.getRange().getAddress();
  const groupAddress = groupSelection.getRange().getAddress();
  const acceptedOld = restaurantAddress === OLD_RESTAURANT_RANGE && groupAddress === OLD_GROUP_RANGE;
  const acceptedNew = restaurantAddress === NEW_RESTAURANT_RANGE && groupAddress === NEW_GROUP_RANGE;
  if (!acceptedOld && !acceptedNew) {
    throw new Error("PUL-030PC-018: Performance selection-table layout is not an accepted pre/post state.");
  }
  if (text(performance.getRange("A43").getValue()) !== "Explain") {
    throw new Error("PUL-030PC-019: Accepted compact Explain section is missing.");
  }
  if (
    text(performance.getRange("A23").getValue()) !== "Restaurant" ||
    text(performance.getRange("B23").getValue()) !== "Total" ||
    text(performance.getRange("G23").getValue()) !== "Beer & Cider" ||
    text(performance.getRange("A40").getValue()) !== "Grand Total"
  ) {
    throw new Error("PUL-030PC-020: Accepted Performance matrix structure is not present.");
  }
}

function validatePostconditions(
  performance: ExcelScript.Worksheet,
  restaurantSelection: ExcelScript.Table,
  groupSelection: ExcelScript.Table
): void {
  if (
    restaurantSelection.getRange().getAddress() !== NEW_RESTAURANT_RANGE ||
    groupSelection.getRange().getAddress() !== NEW_GROUP_RANGE ||
    text(performance.getRange("B50").getValue()) !== "Restaurant selection" ||
    text(performance.getRange("F50").getValue()) !== "Reporting Group selection"
  ) {
    throw new Error("PUL-030PC-021: Selection-table placement postcondition failed.");
  }
  if (!rangeIsBlank(performance.getRange("N4:T21"))) {
    throw new Error("PUL-030PC-022: The old far-right selection area is not empty.");
  }
  if (text(performance.getRange("B45").getValue()) !== NEW_DEFINITION) {
    throw new Error("PUL-030PC-023: Explain wording postcondition failed.");
  }
  if (
    performance.getRange("A23:A40").getFormat().getHorizontalAlignment() !== ExcelScript.HorizontalAlignment.left ||
    performance.getRange("B23:K40").getFormat().getHorizontalAlignment() !== ExcelScript.HorizontalAlignment.center ||
    performance.getRange("B23:K23").getFormat().getWrapText() !== true
  ) {
    throw new Error("PUL-030PC-024: Matrix alignment/wrapping postcondition failed.");
  }
  const widths: { address: string; expected: number }[] = [
    { address: "B:B", expected: 98 }, { address: "C:C", expected: 98 },
    { address: "D:D", expected: 102 }, { address: "E:E", expected: 108 },
    { address: "F:F", expected: 100 }, { address: "G:G", expected: 105 },
    { address: "H:H", expected: 110 }, { address: "I:I", expected: 100 },
    { address: "J:J", expected: 95 }, { address: "K:K", expected: 95 }
  ];
  for (const item of widths) {
    if (Math.abs(performance.getRange(item.address).getFormat().getColumnWidth() - item.expected) > 0.5) {
      throw new Error(`PUL-030PC-025: Matrix column ${item.address} width postcondition failed.`);
    }
  }
  if (
    performance.getRange("G6:G8").getFormat().getWrapText() !== true ||
    groupSelection.getHeaderRowRange().getFormat().getWrapText() !== true ||
    performance.getRange("51:51").getFormat().getRowHeight() < 31.5 ||
    performance.getRange("52:67").getFormat().getRowHeight() < 27.5
  ) {
    throw new Error("PUL-030PC-026: Top-control or selection-table readability postcondition failed.");
  }
}

function validatePrimaryIa(workbook: ExcelScript.Workbook): void {
  const expected: string[] = ["Overview", "Performance", "Reports", "Imports", "Mapping", "Settings"];
  const sheets = workbook.getWorksheets();
  let visibleCount = 0;
  for (let index = 0; index < sheets.length; index++) {
    if (sheets[index].getVisibility() === ExcelScript.SheetVisibility.visible) visibleCount++;
  }
  if (visibleCount !== 6) throw new Error(`PUL-030PC-027: Expected six visible sheets; found ${visibleCount}.`);
  for (let index = 0; index < expected.length; index++) {
    const sheet = requiredSheet(workbook, expected[index]);
    if (sheet.getVisibility() !== ExcelScript.SheetVisibility.visible || sheet.getPosition() !== index) {
      throw new Error(`PUL-030PC-028: Visible IA differs at ${expected[index]}.`);
    }
  }
}

function validateQa(table: ExcelScript.Table, prefix: string, errorCode: string): void {
  const headers = table.getHeaderRowRange().getTexts()[0];
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  const checkColumn = headerIndex(headers, "CheckID");
  const resultColumn = headerIndex(headers, "Result");
  let passCount = 0;
  for (let index = 1; index <= 16; index++) {
    const expected = `${prefix}${String(index).padStart(2, "0")}`;
    let result = "";
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
      if (text(rows[rowIndex][checkColumn]) === expected) result = text(rows[rowIndex][resultColumn]);
    }
    if (result === "PASS") passCount++;
  }
  if (passCount !== 16) throw new Error(`${errorCode}: ${prefix} QA is ${passCount}/16 PASS.`);
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

function rangeIsBlank(range: ExcelScript.Range): boolean {
  const values = range.getValues();
  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    for (let columnIndex = 0; columnIndex < values[rowIndex].length; columnIndex++) {
      if (text(values[rowIndex][columnIndex]) !== "") return false;
    }
  }
  return true;
}

function tableContentFingerprint(table: ExcelScript.Table, prefix: string): string {
  return matrixFingerprint(table.getRange().getValues(), table.getRange().getFormulas(), prefix);
}

function rangeFingerprint(range: ExcelScript.Range, prefix: string): string {
  return matrixFingerprint(range.getValues(), range.getFormulas(), prefix);
}

function matrixFingerprint(
  values: (string | number | boolean)[][],
  formulas: (string | number | boolean)[][],
  prefix: string
): string {
  const state = newHashState();
  updateHashMatrix(state, values);
  updateHashMatrix(state, formulas);
  return finishHash(state, prefix);
}

function formulaFingerprint(values: (string | number | boolean)[][]): string {
  const state = newHashState();
  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    for (let columnIndex = 0; columnIndex < values[rowIndex].length; columnIndex++) {
      const value = text(values[rowIndex][columnIndex]);
      if (value.indexOf("=") === 0) updateHash(state, `${rowIndex}:${columnIndex}:${value.length}:${value}|`);
    }
  }
  return finishHash(state, "PC-FORMULA-");
}

type HashState = { left: number; right: number };

function newHashState(): HashState { return { left: 0, right: 0 }; }

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
  if (!sheet) throw new Error(`PUL-030PC-029: Required worksheet missing: ${name}.`);
  return sheet;
}

function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error(`PUL-030PC-030: Required table missing: ${name}.`);
  return table;
}

function requiredColumn(table: ExcelScript.Table, name: string): ExcelScript.TableColumn {
  const column = table.getColumnByName(name);
  if (!column) throw new Error(`PUL-030PC-031: ${table.getName()} is missing column ${name}.`);
  return column;
}

function headerIndex(headers: string[], name: string): number {
  for (let index = 0; index < headers.length; index++) if (text(headers[index]) === name) return index;
  throw new Error(`PUL-030PC-032: Required header missing: ${name}.`);
}

function text(value: unknown): string {
  return String(value === null || value === undefined ? "" : value).trim();
}
