#!/usr/bin/env node

/**
 * Read-only development audit for the local weekly fixture corpus.
 *
 * Usage:
 *   node src/imports/audit-weekly-corpus.mjs /exact/corpus/path
 *
 * The exact path is mandatory. This command never searches Desktop, Downloads,
 * OneDrive, or the repository, and it never writes to the source workbooks.
 * Production intake remains OneDrive + Power Automate + Office Scripts.
 */

import { createHash } from "node:crypto";
import { readFile, stat } from "node:fs/promises";
import path from "node:path";
import {
  WEEKLY_SALES_SHEET,
  buildWeeklyCorpusManifest,
  parseWeeklySalesMatrix,
} from "./weekly-sales-parser.mjs";
import {
  listWeeklyXlsxFiles,
  readWeeklyWorkbookMatrix,
} from "./weekly-sales-ooxml.mjs";

const root = process.argv[2];
if (!root || root.startsWith("--")) {
  throw new Error("Provide one exact fixture corpus path. The audit will not search for source files.");
}
const expectedFlag = process.argv.indexOf("--expected");
const expectedPath = expectedFlag >= 0 ? process.argv[expectedFlag + 1] : "";
if (expectedFlag >= 0 && !expectedPath) throw new Error("--expected requires one exact JSON path.");

const files = await listWeeklyXlsxFiles(path.resolve(root));
if (files.length === 0) throw new Error(`No .xlsx reports found under ${root}.`);

const parsedReports = [];
const binaryBytes = [];
for (const file of files) {
  const fileBytes = await readFile(file);
  const workbook = readWeeklyWorkbookMatrix(file);
  if (workbook.sheetNames.length !== 1 || workbook.sheetNames[0] !== WEEKLY_SALES_SHEET) {
    throw new Error(`${file} must contain exactly one '${WEEKLY_SALES_SHEET}' sheet.`);
  }
  const binaryFingerprint = `SHA256-${createHash("sha256").update(fileBytes).digest("hex")}`;
  const parsed = parseWeeklySalesMatrix(workbook.matrix, {
    sourceLocator: path.relative(path.resolve(root), file),
    sourceBinaryFingerprint: binaryFingerprint,
  });
  parsedReports.push(parsed);
  binaryBytes.push((await stat(file)).size);
}

const corpus = buildWeeklyCorpusManifest(parsedReports);
const rows = parsedReports.flatMap(value => value.rows);
const result = {
  root: path.resolve(root),
  readOnly: true,
  reportCount: corpus.reportCount,
  uniquePeriodCount: corpus.uniquePeriodCount,
  firstPeriodStart: corpus.firstPeriodStart,
  lastPeriodEnd: corpus.lastPeriodEnd,
  sourceRowCount: corpus.sourceRowCount,
  totalSalesNok: corpus.totalSalesNok,
  totalQuantity: corpus.totalQuantity,
  totalSourceBytes: binaryBytes.reduce((total, value) => total + value, 0),
  uniqueRestaurantCount: new Set(rows.map(value => value.restaurant)).size,
  uniqueItemCount: new Set(rows.map(value => value.item)).size,
  scopeFingerprints: corpus.scopeFingerprints,
  schemaVersions: corpus.schemaVersions,
  coverage: corpus.coverage,
  duplicatePeriods: corpus.duplicatePeriods,
  duplicateFileIds: corpus.duplicateFileIds,
  overlaps: corpus.overlaps,
  gaps: corpus.gaps,
  restaurantSetSignatures: corpus.restaurantSetSignatures,
  salesAccountSetSignatureCount: corpus.salesAccountSetSignatures.length,
  corpusFingerprint: corpus.corpusFingerprint,
  status: corpus.status,
};

if (expectedPath) {
  const expected = JSON.parse(await readFile(path.resolve(expectedPath), "utf8"));
  const actual = {
    report_count: result.reportCount,
    unique_period_count: result.uniquePeriodCount,
    first_period_start: result.firstPeriodStart,
    last_period_end: result.lastPeriodEnd,
    source_row_count: result.sourceRowCount,
    total_sales_nok: result.totalSalesNok,
    total_quantity: result.totalQuantity,
    total_source_bytes: result.totalSourceBytes,
    unique_restaurant_source_values: result.uniqueRestaurantCount,
    unique_item_source_values: result.uniqueItemCount,
    scope_fingerprint: result.scopeFingerprints[0] ?? "",
    schema_version: result.schemaVersions[0] ?? "",
    corpus_fingerprint: result.corpusFingerprint,
    duplicate_period_count: result.duplicatePeriods.length,
    duplicate_file_id_count: result.duplicateFileIds.length,
    overlap_count: result.overlaps.length,
    gap_count: result.gaps.length,
    status: result.status,
  };
  for (const [key, value] of Object.entries(actual)) {
    if (expected[key] !== value) {
      throw new Error(`Expected corpus checkpoint mismatch for ${key}: expected ${expected[key]}, received ${value}.`);
    }
  }
  result.expectedCheckpointStatus = "PASS";
}

process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
