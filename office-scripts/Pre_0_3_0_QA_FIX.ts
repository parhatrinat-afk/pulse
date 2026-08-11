/**
 * Pulse — Pre-0.3.0 QA / UI Fix
 *
 * Run this against the latest corrected 0.2.0 workbook.
 *
 * Fixes found during workbook review:
 * - Performance and Reports column widths were far too narrow.
 * - Category validation was stored as a very long literal list, which is fragile.
 * - Period validation is moved to the helper range for consistency.
 * - User-facing pages receive consistent readable sizing/wrapping.
 *
 * This script does NOT change sales facts, mapping logic, metric formulas,
 * category membership, or current/comparison calculation rules.
 */
function main(workbook: ExcelScript.Workbook): string {
  const performance = requiredSheet(workbook, "Performance");
  const reports = requiredSheet(workbook, "Reports");
  const overview = requiredSheet(workbook, "Overview");
  const imports = requiredSheet(workbook, "Imports");
  const settings = requiredSheet(workbook, "Settings");
  const calc = requiredSheet(workbook, "_Metric_Calc");
  const environment = requiredSheet(workbook, "_Environment");
  const buildLog = requiredSheet(workbook, "_Build_Log");

  const categoryCount = countPopulated(calc.getRange("I2:I500").getValues());
  const datasetCount = countPopulated(calc.getRange("A2:A100").getValues());

  if (categoryCount < 1) {
    throw new Error("PUL-QA-001: _Metric_Calc contains no category selector values.");
  }
  if (datasetCount < 1) {
    throw new Error("PUL-QA-002: _Metric_Calc contains no dataset selector values.");
  }

  // -------------------------------------------------------------------------
  // Performance: make the working surface genuinely readable.
  // Office Scripts column width values are much smaller than Excel's visible
  // widths when values like 20–30 are used, so use explicit practical widths.
  // -------------------------------------------------------------------------
  performance.getRange("A:A").getFormat().setColumnWidth(150);
  performance.getRange("B:B").getFormat().setColumnWidth(185);
  performance.getRange("C:E").getFormat().setColumnWidth(90);
  performance.getRange("F:F").getFormat().setColumnWidth(135);
  performance.getRange("G:G").getFormat().setColumnWidth(220);
  performance.getRange("H:L").getFormat().setColumnWidth(90);

  performance.getRange("A1:L200").getFormat()
    .setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  performance.getRange("A1:L200").getFormat().setWrapText(false);

  // Longer explanatory areas may wrap.
  for (const address of ["A2:L2", "G18:I18", "B48:H60"]) {
    performance.getRange(address).getFormat().setWrapText(true);
  }

  performance.getRange("1:1").getFormat().setRowHeight(30);
  performance.getRange("2:2").getFormat().setRowHeight(28);
  performance.getRange("5:19").getFormat().setRowHeight(22);
  performance.getRange("23:200").getFormat().setRowHeight(20);

  // Replace the huge literal category list with a proper range-backed dropdown.
  performance.getRange("B7").getDataValidation().clear();
  performance.getRange("B7").getDataValidation().setRule({
    list: {
      inCellDropDown: true,
      source: `='_Metric_Calc'!$I$2:$I$${categoryCount + 1}`
    }
  });

  // Current and comparison dataset selectors use the same authoritative list.
  const datasetSource =
    `='_Metric_Calc'!$A$2:$A$${datasetCount + 1}`;

  for (const cell of ["B10", "G10"]) {
    performance.getRange(cell).getDataValidation().clear();
    performance.getRange(cell).getDataValidation().setRule({
      list: {
        inCellDropDown: true,
        source: datasetSource
      }
    });
  }

  // Give editable selector cells an understated visual cue.
  for (const cell of ["B7", "B10", "G10"]) {
    const range: ExcelScript.Range = performance.getRange(cell);
    range.getFormat().getFill().setColor("#FFFFFF");
    range.getFormat().getFont().setBold(true);

    const borders: ExcelScript.RangeBorder[] =
      range.getFormat().getBorders();

    for (const border of borders) {
      if (border.getSideIndex() === ExcelScript.BorderIndex.edgeBottom) {
        border.setStyle(ExcelScript.BorderLineStyle.continuous);
        border.setColor("#9CA3AF");
      }
    }
  }

  performance.getFreezePanes().freezeRows(2);

  // -------------------------------------------------------------------------
  // Reports: same visual language, without giant or tiny columns.
  // -------------------------------------------------------------------------
  reports.getRange("A:A").getFormat().setColumnWidth(155);
  reports.getRange("B:B").getFormat().setColumnWidth(235);
  reports.getRange("C:H").getFormat().setColumnWidth(90);
  reports.getRange("A1:H40").getFormat()
    .setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  reports.getRange("A2:H2").getFormat().setWrapText(true);
  reports.getRange("1:1").getFormat().setRowHeight(30);
  reports.getRange("2:2").getFormat().setRowHeight(28);
  reports.getRange("5:25").getFormat().setRowHeight(22);
  reports.getFreezePanes().freezeRows(2);

  // -------------------------------------------------------------------------
  // Other user-facing pages: retain the content, normalize readability.
  // -------------------------------------------------------------------------
  overview.getRange("A:H").getFormat()
    .setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  overview.getRange("A2:H2").getFormat().setWrapText(true);

  imports.getRange("A:S").getFormat()
    .setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  imports.getRange("A1:S3").getFormat().setWrapText(true);
  imports.getFreezePanes().freezeRows(4);

  settings.getRange("A:C").getFormat()
    .setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  settings.getRange("A1:C20").getFormat().setWrapText(true);
  settings.getFreezePanes().freezeRows(4);

  // -------------------------------------------------------------------------
  // Build metadata: this is QA polish, not Build 0.3.0.
  // -------------------------------------------------------------------------
  updateEnvironment(environment, "BuildID", "0.2.0-QA");
  updateEnvironment(environment, "BuildVersion", "0.2.0-QA");

  appendLog(buildLog, [
    nextLogId(buildLog),
    excelNow(),
    "0.2.0-QA",
    "Pre-0.3.0 usability review",
    "Success",
    "UI and validation",
    "Corrected Performance/Reports sizing and replaced fragile literal selector lists with range-backed dropdowns."
  ]);

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  return (
    `Pulse QA fix applied. ${categoryCount} category choices and ` +
    `${datasetCount} dataset choices are connected to range-backed dropdowns.`
  );
}

function requiredSheet(
  workbook: ExcelScript.Workbook,
  name: string
): ExcelScript.Worksheet {
  const sheet = workbook.getWorksheet(name);
  if (!sheet) throw new Error(`PUL-QA-003: Required sheet missing: ${name}`);
  return sheet;
}

function countPopulated(values: (string | number | boolean)[][]): number {
  let count = 0;
  for (const row of values) {
    if (String(row[0] ?? "").trim() !== "") count++;
  }
  return count;
}

function updateEnvironment(
  sheet: ExcelScript.Worksheet,
  key: string,
  value: string
): void {
  const used = sheet.getUsedRange();
  if (!used) return;
  const rows = used.getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0] ?? "").trim() === key) {
      sheet.getCell(i, 1).setValue(value);
      return;
    }
  }
}

function appendLog(
  sheet: ExcelScript.Worksheet,
  row: (string | number | boolean)[]
): void {
  const used = sheet.getUsedRange();
  const nextRow = used ? used.getRowCount() : 1;
  sheet.getRangeByIndexes(nextRow, 0, 1, row.length).setValues([row]);
}

function nextLogId(sheet: ExcelScript.Worksheet): string {
  const used = sheet.getUsedRange();
  if (!used) return "LOG-000001";
  const rows = used.getValues();
  let max = 0;
  for (const row of rows) {
    const id = String(row[0] ?? "").trim();
    const match = id.match(/^LOG-(\d+)$/);
    if (!match) continue;
    const n = Number(match[1]);
    if (Number.isFinite(n)) max = Math.max(max, n);
  }
  return `LOG-${String(max + 1).padStart(6, "0")}`;
}

function excelNow(): number {
  return Date.now() / 86400000 + 25569;
}
