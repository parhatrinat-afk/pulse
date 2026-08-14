/**
 * Read-only adapter for the Build 0.3.0 weekly Sales per Item parser contract.
 *
 * The script reads one untouched POS export and returns normalized source rows
 * plus a deterministic manifest. It does not stage, publish, map, or edit any
 * workbook. sourceLocator and sourceBinaryFingerprint are audit metadata only;
 * neither participates in period identity or the semantic source fingerprint.
 */

const PARSER_VERSION = "0.3.0-weekly-parser-v1";
const SCHEMA_VERSION = "sales-per-item-v1";
const SOURCE_SHEET = "Sales per Item";
const EXPECTED_HEADERS = [
  "Restaurant", "Main Category", "Sub Category", "Sales Account",
  "Item", "Quantity", "Amount"
];

interface ScopeContract {
  scopeId: string;
  scopeName: string;
  sourceSystemId: string;
  reportType: string;
  channelScope: string;
  scopeAuthority: string;
  status: string;
}

interface NormalizedWeeklySalesRow {
  sourceRowId: string;
  sourceFileId: string;
  sourceRowNumber: number;
  sourcePeriodKey: string;
  periodStart: string;
  periodEnd: string;
  isoYear: number;
  isoWeek: number;
  restaurant: string;
  mainCategory: string;
  subCategory: string;
  salesAccount: string;
  item: string;
  quantity: number;
  salesNok: number;
}

interface WeeklySourceManifest {
  sourceFileId: string;
  sourceLocator: string;
  sourceBinaryFingerprint: string;
  semanticFingerprint: string;
  sourcePeriodKey: string;
  periodStart: string;
  periodEnd: string;
  inclusiveDays: number;
  isoYear: number;
  isoWeek: number;
  sheetName: string;
  schemaVersion: string;
  parserVersion: string;
  scopeId: string;
  scopeFingerprint: string;
  channelScope: string;
  channelFieldEncoded: boolean;
  exportFilterMetadataEncoded: boolean;
  sourceRowCount: number;
  totalSalesNok: number;
  totalQuantity: number;
  restaurantCount: number;
  restaurants: string[];
  mainCategoryCount: number;
  subCategoryCount: number;
  salesAccountCount: number;
  salesAccounts: string[];
  itemCount: number;
  contentReconciliationStatus: string;
}

interface WeeklyParserResult {
  parserVersion: string;
  schemaVersion: string;
  scopeContract: ScopeContract;
  manifest: WeeklySourceManifest;
  rows: NormalizedWeeklySalesRow[];
}

interface ParsedSourceRow {
  sourceRowNumber: number;
  restaurant: string;
  mainCategory: string;
  subCategory: string;
  salesAccount: string;
  item: string;
  quantity: number;
  salesNok: number;
}

interface PeriodContract {
  periodStart: string;
  periodEnd: string;
  isoYear: number;
  isoWeek: number;
  sourcePeriodKey: string;
}

interface IsoWeek {
  year: number;
  week: number;
}

function main(
  workbook: ExcelScript.Workbook,
  sourceLocator: string = "",
  sourceBinaryFingerprint: string = ""
): WeeklyParserResult {
  const sheets = workbook.getWorksheets();
  if (sheets.length !== 1 || sheets[0].getName() !== SOURCE_SHEET) {
    throw new Error(`PUL-030I-040: Expected one worksheet named '${SOURCE_SHEET}'.`);
  }
  const used = sheets[0].getUsedRange(true);
  if (!used) throw new Error("PUL-030I-001: Source matrix is missing.");
  const matrix = used.getValues();
  return parseWeeklySalesMatrix(matrix, sourceLocator, sourceBinaryFingerprint);
}

function parseWeeklySalesMatrix(
  matrix: (string | number | boolean)[][],
  sourceLocator: string,
  sourceBinaryFingerprint: string
): WeeklyParserResult {
  if (matrix.length < 3) throw new Error("PUL-030I-002: Sales per Item report has no data rows.");
  const scopeContract = weeklyScopeContract();
  const period = parsePeriodCell(matrix[0][0]);
  const headers = normalizedRow(matrix[1], EXPECTED_HEADERS.length);
  validateHeaders(headers);
  const lastDataIndex = lastPopulatedRow(matrix, 2);
  if (lastDataIndex < 2) throw new Error("PUL-030I-002: Sales per Item report has no data rows.");

  const parsedRows: ParsedSourceRow[] = [];
  for (let matrixIndex = 2; matrixIndex <= lastDataIndex; matrixIndex++) {
    const sourceRowNumber = matrixIndex + 1;
    const source = matrix[matrixIndex] || [];
    if (blankRow(source)) {
      throw new Error(`PUL-030I-010: Blank row found inside the data region at row ${sourceRowNumber}.`);
    }
    for (let column = EXPECTED_HEADERS.length; column < source.length; column++) {
      if (!blank(source[column])) {
        throw new Error(`PUL-030I-011: Unexpected populated column after Amount at row ${sourceRowNumber}.`);
      }
    }
    const values = normalizedRow(source, EXPECTED_HEADERS.length);
    for (let column = 0; column < 5; column++) {
      if (typeof values[column] !== "string" || String(values[column]).length === 0) {
        throw new Error(`PUL-030I-012: Required source text is blank or non-text at row ${sourceRowNumber}.`);
      }
    }
    parsedRows.push({
      sourceRowNumber,
      restaurant: String(values[0]),
      mainCategory: String(values[1]),
      subCategory: String(values[2]),
      salesAccount: String(values[3]),
      item: String(values[4]),
      quantity: strictNumber(values[5], "Quantity", sourceRowNumber),
      salesNok: strictNumber(values[6], "Amount", sourceRowNumber)
    });
  }

  const rowRecords: string[] = [];
  for (let index = 0; index < parsedRows.length; index++) {
    rowRecords.push(sourceRowRecord(parsedRows[index]));
  }
  const sortedRecords = rowRecords.slice().sort();
  const semanticRecords = [
    fingerprintRecord("PERIOD", [period.periodStart, period.periodEnd]),
    fingerprintRecord("HEADER", EXPECTED_HEADERS)
  ];
  for (let index = 0; index < sortedRecords.length; index++) semanticRecords.push(sortedRecords[index]);
  const semanticFingerprint = hashStrings(semanticRecords, "WSF-");
  const sourceFileId = `SRCFILE-${semanticFingerprint.substring(4)}`;
  const occurrences: { [key: string]: number } = {};
  const rows: NormalizedWeeklySalesRow[] = [];
  for (let index = 0; index < parsedRows.length; index++) {
    const row = parsedRows[index];
    const rowRecord = rowRecords[index];
    const occurrence = (occurrences[rowRecord] || 0) + 1;
    occurrences[rowRecord] = occurrence;
    rows.push({
      sourceRowId: hashStrings([sourceFileId, rowRecord, String(occurrence)], "SROW-"),
      sourceFileId,
      sourceRowNumber: row.sourceRowNumber,
      sourcePeriodKey: period.sourcePeriodKey,
      periodStart: period.periodStart,
      periodEnd: period.periodEnd,
      isoYear: period.isoYear,
      isoWeek: period.isoWeek,
      restaurant: row.restaurant,
      mainCategory: row.mainCategory,
      subCategory: row.subCategory,
      salesAccount: row.salesAccount,
      item: row.item,
      quantity: row.quantity,
      salesNok: row.salesNok
    });
  }

  const restaurants = uniqueSorted(rows.map(row => row.restaurant));
  const mainCategories = uniqueSorted(rows.map(row => row.mainCategory));
  const subCategories = uniqueSorted(rows.map(row => row.subCategory));
  const salesAccounts = uniqueSorted(rows.map(row => row.salesAccount));
  const items = uniqueSorted(rows.map(row => row.item));
  let totalSalesNok = 0;
  let totalQuantity = 0;
  for (let index = 0; index < rows.length; index++) {
    totalSalesNok += rows[index].salesNok;
    totalQuantity += rows[index].quantity;
  }
  return {
    parserVersion: PARSER_VERSION,
    schemaVersion: SCHEMA_VERSION,
    scopeContract,
    manifest: {
      sourceFileId,
      sourceLocator,
      sourceBinaryFingerprint,
      semanticFingerprint,
      sourcePeriodKey: period.sourcePeriodKey,
      periodStart: period.periodStart,
      periodEnd: period.periodEnd,
      inclusiveDays: 7,
      isoYear: period.isoYear,
      isoWeek: period.isoWeek,
      sheetName: SOURCE_SHEET,
      schemaVersion: SCHEMA_VERSION,
      parserVersion: PARSER_VERSION,
      scopeId: scopeContract.scopeId,
      scopeFingerprint: fingerprintScopeContract(scopeContract),
      channelScope: scopeContract.channelScope,
      channelFieldEncoded: false,
      exportFilterMetadataEncoded: false,
      sourceRowCount: rows.length,
      totalSalesNok: roundNumber(totalSalesNok, 2),
      totalQuantity: roundNumber(totalQuantity, 6),
      restaurantCount: restaurants.length,
      restaurants,
      mainCategoryCount: mainCategories.length,
      subCategoryCount: subCategories.length,
      salesAccountCount: salesAccounts.length,
      salesAccounts,
      itemCount: items.length,
      contentReconciliationStatus: "PASS"
    },
    rows
  };
}

function weeklyScopeContract(): ScopeContract {
  return {
    scopeId: "SCOPE-030-WEEKLY-SALES-PER-ITEM",
    scopeName: "POS Sales per Item weekly export",
    sourceSystemId: "SRC-TEST-SALES",
    reportType: "Sales per Item",
    channelScope: "Source-defined scope; channel not encoded",
    scopeAuthority: "Human-approved export configuration",
    status: "Candidate authoritative weekly source"
  };
}

function fingerprintScopeContract(scope: ScopeContract): string {
  return hashStrings([fingerprintRecord("SCOPE", [
    scope.scopeId, scope.scopeName, scope.sourceSystemId, scope.reportType,
    scope.channelScope, scope.scopeAuthority, scope.status
  ])], "SCP-");
}

function parsePeriodCell(value: string | number | boolean): PeriodContract {
  if (typeof value !== "string") throw new Error("PUL-030I-003: A1 must contain the internal Period field.");
  const match = /^Period: (\d{4}-\d{2}-\d{2}) - (\d{4}-\d{2}-\d{2})$/.exec(value);
  if (!match) throw new Error("PUL-030I-004: Period must use 'Period: YYYY-MM-DD - YYYY-MM-DD'.");
  const start = parseIsoDate(match[1]);
  const end = parseIsoDate(match[2]);
  if (daysBetween(start, end) + 1 !== 7 || start.getUTCDay() !== 1 || end.getUTCDay() !== 0) {
    throw new Error("PUL-030I-005: Period must be exactly Monday through Sunday inclusive.");
  }
  const iso = calculateIsoWeek(start);
  const endIso = calculateIsoWeek(end);
  if (iso.year !== endIso.year || iso.week !== endIso.week) {
    throw new Error("PUL-030I-006: PeriodStart and PeriodEnd must belong to one ISO week.");
  }
  return {
    periodStart: match[1],
    periodEnd: match[2],
    isoYear: iso.year,
    isoWeek: iso.week,
    sourcePeriodKey: `PERIOD-${match[1]}-${match[2]}`
  };
}

function validateHeaders(headers: (string | number | boolean)[]): void {
  for (let index = 0; index < EXPECTED_HEADERS.length; index++) {
    if (headers[index] !== EXPECTED_HEADERS[index]) {
      throw new Error(`PUL-030I-008: Header ${index + 1} must be '${EXPECTED_HEADERS[index]}'.`);
    }
  }
}

function strictNumber(value: string | number | boolean, field: string, row: number): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`PUL-030I-013: ${field} must be a finite numeric cell at row ${row}.`);
  }
  return value;
}

function parseIsoDate(value: string): Date {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) throw new Error(`PUL-030I-014: Invalid ISO date ${value}.`);
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    throw new Error(`PUL-030I-014: Invalid ISO date ${value}.`);
  }
  return date;
}

function calculateIsoWeek(date: Date): IsoWeek {
  const value = new Date(date.getTime());
  const day = value.getUTCDay() || 7;
  value.setUTCDate(value.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(value.getUTCFullYear(), 0, 1));
  return {
    year: value.getUTCFullYear(),
    week: Math.ceil((((value.getTime() - yearStart.getTime()) / 86400000) + 1) / 7)
  };
}

function daysBetween(left: Date, right: Date): number {
  return Math.round((right.getTime() - left.getTime()) / 86400000);
}

function normalizedRow(
  row: (string | number | boolean)[],
  length: number
): (string | number | boolean)[] {
  const values = row ? row.slice(0, length) : [];
  while (values.length < length) values.push("");
  return values;
}

function lastPopulatedRow(matrix: (string | number | boolean)[][], startIndex: number): number {
  for (let index = matrix.length - 1; index >= startIndex; index--) {
    if (!blankRow(matrix[index])) return index;
  }
  return -1;
}

function blankRow(row: (string | number | boolean)[]): boolean {
  if (!row) return true;
  for (let index = 0; index < row.length; index++) if (!blank(row[index])) return false;
  return true;
}

function blank(value: string | number | boolean): boolean {
  return value === "" || value === null || value === undefined;
}

function sourceRowRecord(row: ParsedSourceRow): string {
  return fingerprintRecord("ROW", [
    row.restaurant, row.mainCategory, row.subCategory, row.salesAccount,
    row.item, canonicalFingerprintNumber(row.quantity, 6),
    canonicalFingerprintNumber(row.salesNok, 2)
  ]);
}

function canonicalFingerprintNumber(value: number, decimals: number): string {
  return String(roundNumber(value, decimals));
}

function uniqueSorted(values: string[]): string[] {
  const seen: { [key: string]: boolean } = {};
  const unique: string[] = [];
  for (let index = 0; index < values.length; index++) {
    if (seen[values[index]]) continue;
    seen[values[index]] = true;
    unique.push(values[index]);
  }
  return unique.sort();
}

function roundNumber(value: number, decimals: number): number {
  const factor = Math.pow(10, decimals);
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

function fingerprintRecord(kind: string, values: (string | number | boolean)[]): string {
  const parts: string[] = [];
  for (let index = 0; index < values.length; index++) {
    const normalized = normalizeFingerprintValue(values[index]);
    parts.push(`${normalized.length}:${normalized}`);
  }
  return `${kind}|${parts.join("|")}`;
}

function normalizeFingerprintValue(value: string | number | boolean): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") return Number.isFinite(value) ? String(value) : "";
  return String(value);
}

function hashStrings(values: string[], prefix: string): string {
  let left = 0;
  let right = 0;
  for (let itemIndex = 0; itemIndex < values.length; itemIndex++) {
    const value = `${values[itemIndex]}\n`;
    for (let index = 0; index < value.length; index++) {
      const code = value.charCodeAt(index);
      left = (left * 131 + code) % 2147483647;
      right = (right * 137 + code) % 2147483629;
    }
  }
  return `${prefix}${left.toString(16).padStart(8, "0")}${right.toString(16).padStart(8, "0")}`;
}
