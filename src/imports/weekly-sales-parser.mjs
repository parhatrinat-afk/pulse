export const WEEKLY_SALES_PARSER_VERSION = "0.3.0-weekly-parser-v1";
export const WEEKLY_SALES_SCHEMA_VERSION = "sales-per-item-v1";
export const WEEKLY_SALES_SHEET = "Sales per Item";
export const WEEKLY_SALES_HEADERS = Object.freeze([
  "Restaurant",
  "Main Category",
  "Sub Category",
  "Sales Account",
  "Item",
  "Quantity",
  "Amount",
]);

/**
 * This contract describes the human-approved export scope. The source workbook
 * has no channel/filter field, so the parser must never infer In-house,
 * Takeaway, or All channels from product/category/account text.
 */
export const WEEKLY_SALES_SCOPE_CONTRACT = Object.freeze({
  scopeId: "SCOPE-030-WEEKLY-SALES-PER-ITEM",
  scopeName: "POS Sales per Item weekly export",
  sourceSystemId: "SRC-TEST-SALES",
  reportType: "Sales per Item",
  channelScope: "Source-defined scope; channel not encoded",
  scopeAuthority: "Human-approved export configuration",
  status: "Candidate authoritative weekly source",
});

export function parseWeeklySalesMatrix(matrix, options = {}) {
  if (!Array.isArray(matrix)) fail("PUL-030I-001", "Source matrix is missing.");
  if (matrix.length < 3) fail("PUL-030I-002", "Sales per Item report has no data rows.");

  const scopeContract = validateScopeContract(options.scopeContract ?? WEEKLY_SALES_SCOPE_CONTRACT);
  const period = parsePeriodCell(matrix[0]?.[0]);
  const headerRow = normalizeMatrixRow(matrix[1], WEEKLY_SALES_HEADERS.length);
  validateHeaders(headerRow);

  const lastDataIndex = findLastPopulatedRow(matrix, 2);
  if (lastDataIndex < 2) fail("PUL-030I-002", "Sales per Item report has no data rows.");

  const parsedRows = [];
  for (let matrixIndex = 2; matrixIndex <= lastDataIndex; matrixIndex += 1) {
    const sourceRowNumber = matrixIndex + 1;
    const source = Array.isArray(matrix[matrixIndex]) ? matrix[matrixIndex] : [];
    if (isBlankRow(source)) {
      fail("PUL-030I-010", `Blank row found inside the data region at row ${sourceRowNumber}.`);
    }
    if (source.slice(WEEKLY_SALES_HEADERS.length).some(value => !isBlank(value))) {
      fail("PUL-030I-011", `Unexpected populated column after Amount at row ${sourceRowNumber}.`);
    }

    const values = normalizeMatrixRow(source, WEEKLY_SALES_HEADERS.length);
    const [restaurant, mainCategory, subCategory, salesAccount, item] = values.slice(0, 5);
    const required = [restaurant, mainCategory, subCategory, salesAccount, item];
    if (required.some(value => typeof value !== "string" || value.length === 0)) {
      fail("PUL-030I-012", `Required source text is blank or non-text at row ${sourceRowNumber}.`);
    }

    const quantity = strictNumber(values[5], "Quantity", sourceRowNumber);
    const salesNok = strictNumber(values[6], "Amount", sourceRowNumber);
    parsedRows.push({
      sourceRowNumber,
      restaurant,
      mainCategory,
      subCategory,
      salesAccount,
      item,
      quantity,
      salesNok,
    });
  }

  const rowRecords = parsedRows.map(row => sourceRowRecord(row));
  const semanticRecords = [
    record("PERIOD", [period.periodStart, period.periodEnd]),
    record("HEADER", WEEKLY_SALES_HEADERS),
    ...rowRecords.slice().sort(),
  ];
  const semanticFingerprint = hashStrings(semanticRecords, "WSF-");
  const sourceFileId = `SRCFILE-${semanticFingerprint.slice(4)}`;
  const occurrenceByRecord = Object.create(null);
  const normalizedRows = parsedRows.map((row, index) => {
    const rowRecord = rowRecords[index];
    const occurrence = (occurrenceByRecord[rowRecord] ?? 0) + 1;
    occurrenceByRecord[rowRecord] = occurrence;
    return {
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
      salesNok: row.salesNok,
    };
  });

  const restaurants = sortedUnique(normalizedRows.map(row => row.restaurant));
  const mainCategories = sortedUnique(normalizedRows.map(row => row.mainCategory));
  const subCategories = sortedUnique(normalizedRows.map(row => row.subCategory));
  const salesAccounts = sortedUnique(normalizedRows.map(row => row.salesAccount));
  const items = sortedUnique(normalizedRows.map(row => row.item));
  const totalQuantity = roundNumber(sum(normalizedRows.map(row => row.quantity)), 6);
  const totalSalesNok = roundNumber(sum(normalizedRows.map(row => row.salesNok)), 2);
  const scopeFingerprint = fingerprintScopeContract(scopeContract);

  return {
    parserVersion: WEEKLY_SALES_PARSER_VERSION,
    schemaVersion: WEEKLY_SALES_SCHEMA_VERSION,
    scopeContract,
    manifest: {
      sourceFileId,
      sourceLocator: String(options.sourceLocator ?? ""),
      sourceBinaryFingerprint: String(options.sourceBinaryFingerprint ?? ""),
      semanticFingerprint,
      sourcePeriodKey: period.sourcePeriodKey,
      periodStart: period.periodStart,
      periodEnd: period.periodEnd,
      inclusiveDays: 7,
      isoYear: period.isoYear,
      isoWeek: period.isoWeek,
      sheetName: WEEKLY_SALES_SHEET,
      schemaVersion: WEEKLY_SALES_SCHEMA_VERSION,
      parserVersion: WEEKLY_SALES_PARSER_VERSION,
      scopeId: scopeContract.scopeId,
      scopeFingerprint,
      channelScope: scopeContract.channelScope,
      channelFieldEncoded: false,
      exportFilterMetadataEncoded: false,
      sourceRowCount: normalizedRows.length,
      totalSalesNok,
      totalQuantity,
      restaurantCount: restaurants.length,
      restaurants,
      mainCategoryCount: mainCategories.length,
      subCategoryCount: subCategories.length,
      salesAccountCount: salesAccounts.length,
      salesAccounts,
      itemCount: items.length,
      contentReconciliationStatus: "PASS",
    },
    rows: normalizedRows,
  };
}

export function buildWeeklyCorpusManifest(parsedReports) {
  if (!Array.isArray(parsedReports) || parsedReports.length === 0) {
    fail("PUL-030I-020", "At least one parsed weekly report is required.");
  }
  const manifests = parsedReports.map(value => value?.manifest ?? value);
  for (const manifest of manifests) {
    if (!manifest?.sourceFileId || !manifest?.sourcePeriodKey) {
      fail("PUL-030I-021", "A parsed report manifest is incomplete.");
    }
    if (manifest.contentReconciliationStatus !== "PASS") {
      fail("PUL-030I-022", `${manifest.sourceFileId} did not reconcile internally.`);
    }
  }

  const ordered = manifests.slice().sort((left, right) => {
    const byStart = String(left.periodStart).localeCompare(String(right.periodStart));
    return byStart || String(left.sourceFileId).localeCompare(String(right.sourceFileId));
  });
  const scopeFingerprints = sortedUnique(ordered.map(value => String(value.scopeFingerprint)));
  const schemaVersions = sortedUnique(ordered.map(value => String(value.schemaVersion)));
  const periodGroups = groupBy(ordered, value => String(value.sourcePeriodKey));
  const duplicatePeriods = Object.keys(periodGroups)
    .filter(key => periodGroups[key].length > 1)
    .map(key => ({
      sourcePeriodKey: key,
      sourceFileIds: periodGroups[key].map(value => value.sourceFileId).sort(),
      semanticFingerprints: sortedUnique(periodGroups[key].map(value => value.semanticFingerprint)),
    }));

  const overlaps = [];
  const gaps = [];
  const uniquePeriods = Object.keys(periodGroups)
    .map(key => periodGroups[key][0])
    .sort((left, right) => String(left.periodStart).localeCompare(String(right.periodStart)));
  for (let index = 1; index < uniquePeriods.length; index += 1) {
    const previous = uniquePeriods[index - 1];
    const current = uniquePeriods[index];
    const expectedStart = addDays(previous.periodEnd, 1);
    if (current.periodStart <= previous.periodEnd) {
      overlaps.push({ previous: previous.sourcePeriodKey, current: current.sourcePeriodKey });
    } else if (current.periodStart !== expectedStart) {
      gaps.push({
        after: previous.sourcePeriodKey,
        before: current.sourcePeriodKey,
        missingStart: expectedStart,
        missingEnd: addDays(current.periodStart, -1),
      });
    }
  }

  const sourceFileIds = ordered.map(value => String(value.sourceFileId));
  const duplicateFileIds = duplicateValues(sourceFileIds);
  const restaurantSetSignatures = groupSignatures(ordered, "restaurants");
  const salesAccountSetSignatures = groupSignatures(ordered, "salesAccounts");
  const coverage = buildIsoCoverage(uniquePeriods);
  const status = scopeFingerprints.length === 1 && schemaVersions.length === 1 &&
    duplicatePeriods.length === 0 && duplicateFileIds.length === 0 &&
    overlaps.length === 0 && gaps.length === 0
    ? "PASS"
    : "FAIL";
  const corpusRecords = ordered.map(value => record("FILE", [
    value.sourceFileId,
    value.semanticFingerprint,
    value.sourcePeriodKey,
    value.sourceRowCount,
    value.totalSalesNok,
    value.totalQuantity,
    value.scopeFingerprint,
  ]));

  return {
    status,
    parserVersion: WEEKLY_SALES_PARSER_VERSION,
    reportCount: ordered.length,
    uniquePeriodCount: uniquePeriods.length,
    firstPeriodStart: uniquePeriods[0]?.periodStart ?? "",
    lastPeriodEnd: uniquePeriods[uniquePeriods.length - 1]?.periodEnd ?? "",
    sourceRowCount: sum(ordered.map(value => value.sourceRowCount)),
    totalSalesNok: roundNumber(sum(ordered.map(value => value.totalSalesNok)), 2),
    totalQuantity: roundNumber(sum(ordered.map(value => value.totalQuantity)), 6),
    scopeFingerprints,
    schemaVersions,
    duplicatePeriods,
    duplicateFileIds,
    overlaps,
    gaps,
    coverage,
    restaurantSetSignatures,
    salesAccountSetSignatures,
    corpusFingerprint: hashStrings(corpusRecords, "WSC-"),
    manifests: ordered,
  };
}

export function fingerprintScopeContract(scopeContract) {
  const scope = validateScopeContract(scopeContract);
  return hashStrings([record("SCOPE", [
    scope.scopeId,
    scope.scopeName,
    scope.sourceSystemId,
    scope.reportType,
    scope.channelScope,
    scope.scopeAuthority,
    scope.status,
  ])], "SCP-");
}

function validateScopeContract(value) {
  const fields = [
    "scopeId", "scopeName", "sourceSystemId", "reportType",
    "channelScope", "scopeAuthority", "status",
  ];
  const result = {};
  for (const field of fields) {
    if (typeof value?.[field] !== "string" || value[field].length === 0) {
      fail("PUL-030I-030", `Scope contract field ${field} is required.`);
    }
    result[field] = value[field];
  }
  return result;
}

function parsePeriodCell(value) {
  if (typeof value !== "string") fail("PUL-030I-003", "A1 must contain the internal Period field.");
  const match = /^Period: (\d{4}-\d{2}-\d{2}) - (\d{4}-\d{2}-\d{2})$/.exec(value);
  if (!match) fail("PUL-030I-004", "Period must use 'Period: YYYY-MM-DD - YYYY-MM-DD'.");
  const start = parseIsoDate(match[1]);
  const end = parseIsoDate(match[2]);
  const inclusiveDays = daysBetween(start, end) + 1;
  if (inclusiveDays !== 7 || start.getUTCDay() !== 1 || end.getUTCDay() !== 0) {
    fail("PUL-030I-005", "Period must be exactly Monday through Sunday inclusive.");
  }
  const iso = isoWeek(start);
  const endIso = isoWeek(end);
  if (iso.year !== endIso.year || iso.week !== endIso.week) {
    fail("PUL-030I-006", "PeriodStart and PeriodEnd must belong to one ISO week.");
  }
  return {
    periodStart: match[1],
    periodEnd: match[2],
    isoYear: iso.year,
    isoWeek: iso.week,
    sourcePeriodKey: `PERIOD-${match[1]}-${match[2]}`,
  };
}

function validateHeaders(headers) {
  if (headers.length !== WEEKLY_SALES_HEADERS.length) {
    fail("PUL-030I-007", "Header width is invalid.");
  }
  for (let index = 0; index < WEEKLY_SALES_HEADERS.length; index += 1) {
    if (headers[index] !== WEEKLY_SALES_HEADERS[index]) {
      fail("PUL-030I-008", `Header ${index + 1} must be '${WEEKLY_SALES_HEADERS[index]}'.`);
    }
  }
}

function strictNumber(value, field, sourceRowNumber) {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    fail("PUL-030I-013", `${field} must be a finite numeric cell at row ${sourceRowNumber}.`);
  }
  return value;
}

function parseIsoDate(value) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) fail("PUL-030I-014", `Invalid ISO date ${value}.`);
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    fail("PUL-030I-014", `Invalid ISO date ${value}.`);
  }
  return date;
}

function isoWeek(date) {
  const value = new Date(date.getTime());
  const day = value.getUTCDay() || 7;
  value.setUTCDate(value.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(value.getUTCFullYear(), 0, 1));
  return {
    year: value.getUTCFullYear(),
    week: Math.ceil((((value.getTime() - yearStart.getTime()) / 86400000) + 1) / 7),
  };
}

function daysBetween(left, right) {
  return Math.round((right.getTime() - left.getTime()) / 86400000);
}

function addDays(value, days) {
  const date = parseIsoDate(value);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

function normalizeMatrixRow(row, length) {
  const values = Array.isArray(row) ? row.slice(0, length) : [];
  while (values.length < length) values.push("");
  return values;
}

function findLastPopulatedRow(matrix, startIndex) {
  for (let index = matrix.length - 1; index >= startIndex; index -= 1) {
    if (!isBlankRow(matrix[index])) return index;
  }
  return -1;
}

function isBlankRow(row) {
  return !Array.isArray(row) || row.every(isBlank);
}

function isBlank(value) {
  return value === "" || value === null || value === undefined;
}

function sourceRowRecord(row) {
  return record("ROW", [
    row.restaurant,
    row.mainCategory,
    row.subCategory,
    row.salesAccount,
    row.item,
    canonicalFingerprintNumber(row.quantity, 6),
    canonicalFingerprintNumber(row.salesNok, 2),
  ]);
}

function canonicalFingerprintNumber(value, decimals) {
  return String(roundNumber(value, decimals));
}

function buildIsoCoverage(manifests) {
  const byYear = groupBy(manifests, value => String(value.isoYear));
  return Object.keys(byYear).sort().map(year => ({
    isoYear: Number(year),
    weeks: sortedUnique(byYear[year].map(value => value.isoWeek)).sort((a, b) => a - b),
    reportCount: byYear[year].length,
  }));
}

function groupSignatures(values, field) {
  const groups = Object.create(null);
  for (const value of values) {
    const items = Array.isArray(value[field]) ? value[field] : [];
    const signature = hashStrings(items.map(item => record("VALUE", [item])), "SET-");
    if (!groups[signature]) groups[signature] = { signature, values: items.slice(), reportCount: 0 };
    groups[signature].reportCount += 1;
  }
  return Object.keys(groups).sort().map(key => groups[key]);
}

function groupBy(values, keySelector) {
  const groups = Object.create(null);
  for (const value of values) {
    const key = keySelector(value);
    if (!groups[key]) groups[key] = [];
    groups[key].push(value);
  }
  return groups;
}

function duplicateValues(values) {
  const counts = Object.create(null);
  for (const value of values) counts[value] = (counts[value] ?? 0) + 1;
  return Object.keys(counts).filter(key => counts[key] > 1).sort();
}

function sortedUnique(values) {
  return [...new Set(values)].sort(compareText);
}

function sum(values) {
  return values.reduce((total, value) => total + Number(value), 0);
}

function roundNumber(value, decimals) {
  const factor = 10 ** decimals;
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

function compareText(left, right) {
  const a = String(left);
  const b = String(right);
  return a < b ? -1 : a > b ? 1 : 0;
}

function record(kind, values) {
  return `${kind}|${values.map(value => {
    const normalized = normalizeFingerprintValue(value);
    return `${normalized.length}:${normalized}`;
  }).join("|")}`;
}

function normalizeFingerprintValue(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") return Number.isFinite(value) ? String(value) : "";
  return String(value);
}

function hashStrings(values, prefix) {
  let left = 0;
  let right = 0;
  for (const item of values) {
    const value = `${item}\n`;
    for (let index = 0; index < value.length; index += 1) {
      const code = value.charCodeAt(index);
      left = (left * 131 + code) % 2147483647;
      right = (right * 137 + code) % 2147483629;
    }
  }
  return `${prefix}${left.toString(16).padStart(8, "0")}${right.toString(16).padStart(8, "0")}`;
}

function fail(code, message) {
  throw new Error(`${code}: ${message}`);
}
