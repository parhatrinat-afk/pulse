#!/usr/bin/env node

/**
 * Read-only development preflight for parsed weekly source identities.
 *
 * Usage:
 *   node src/imports/audit-weekly-identity-preflight.mjs /exact/corpus/path \
 *     --catalog /exact/catalog.json [--expected /exact/expected.json]
 *
 * This command never searches for inputs and never writes to the corpus,
 * Pulse_Current.xlsx, or the supplied catalog.
 */

import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";
import {
  buildWeeklyIdentityPreflight,
  summarizeWeeklyIdentityPreflight,
} from "./weekly-identity-preflight.mjs";
import {
  listWeeklyXlsxFiles,
  readWeeklyWorkbookMatrix,
} from "./weekly-sales-ooxml.mjs";
import {
  WEEKLY_SALES_SHEET,
  parseWeeklySalesMatrix,
} from "./weekly-sales-parser.mjs";

const root = requiredPositionalPath(process.argv[2]);
const catalogPath = requiredFlagPath("--catalog");
const expectedPath = optionalFlagPath("--expected");
const absoluteRoot = path.resolve(root);
const files = await listWeeklyXlsxFiles(absoluteRoot);
if (!files.length) throw new Error(`No .xlsx reports found under ${root}.`);

const parsedReports = [];
for (const file of files) {
  const bytes = await readFile(file);
  const workbook = readWeeklyWorkbookMatrix(file);
  if (workbook.sheetNames.length !== 1 || workbook.sheetNames[0] !== WEEKLY_SALES_SHEET) {
    throw new Error(`${file} must contain exactly one '${WEEKLY_SALES_SHEET}' sheet.`);
  }
  parsedReports.push(parseWeeklySalesMatrix(workbook.matrix, {
    sourceLocator: path.relative(absoluteRoot, file),
    sourceBinaryFingerprint: `SHA256-${createHash("sha256").update(bytes).digest("hex")}`,
  }));
}

const catalogs = JSON.parse(await readFile(path.resolve(catalogPath), "utf8"));
const result = buildWeeklyIdentityPreflight({ parsedReports, catalogs });
const summary = summarizeWeeklyIdentityPreflight(result, 10);
const checkpoint = buildCheckpoint(summary);

if (expectedPath) {
  const expected = JSON.parse(await readFile(path.resolve(expectedPath), "utf8"));
  if (JSON.stringify(checkpoint) !== JSON.stringify(expected)) {
    throw new Error(
      `PUL-030I-108: Weekly identity checkpoint differs.\nExpected ${JSON.stringify(expected)}\nActual ${JSON.stringify(checkpoint)}`,
    );
  }
  summary.expectedCheckpointStatus = "PASS";
}

process.stdout.write(`${JSON.stringify({ readOnly: true, ...summary }, null, 2)}\n`);

function requiredPositionalPath(value) {
  if (!value || value.startsWith("--")) {
    throw new Error("Provide one exact fixture corpus path. The preflight will not search for source files.");
  }
  return value;
}

function requiredFlagPath(flag) {
  const value = optionalFlagPath(flag);
  if (!value) throw new Error(`${flag} requires one exact JSON path.`);
  return value;
}

function optionalFlagPath(flag) {
  const index = process.argv.indexOf(flag);
  if (index < 0) return "";
  const value = process.argv[index + 1];
  if (!value || value.startsWith("--")) throw new Error(`${flag} requires one exact JSON path.`);
  return value;
}

function buildCheckpoint(summary) {
  const tests = summary.newIdentityCandidates.restaurants.filter(row => row.testDepartment);
  return {
    contract_version: summary.contractVersion,
    status: summary.status,
    source_system_id: summary.sourceSystemId,
    mapping_content_fingerprint: summary.mappingContentFingerprint,
    report_count: summary.reportCount,
    source: snakeMetric(summary.sourceTotals),
    known_reused: {
      restaurants: summary.knownIdentitiesReused.restaurants.distinctIdentityCount,
      products: summary.knownIdentitiesReused.products.distinctIdentityCount,
      classifications: summary.knownIdentitiesReused.classifications.distinctIdentityCount,
    },
    new_candidates: {
      restaurants: summary.newIdentityCandidates.restaurants.length,
      products: summary.newIdentityCandidates.productCount,
      product_id_range: summary.newIdentityCandidates.productIdRange,
      classifications: summary.newIdentityCandidates.classificationCount,
      classification_id_range: summary.newIdentityCandidates.classificationIdRange,
    },
    test_departments: tests.map(row => ({
      restaurant_id: row.restaurantId,
      source_name: row.sourceRestaurantName,
      reporting_enabled: row.reportingEnabled,
      impact: snakeMetric(row.impact),
    })),
    identity_pending_count: summary.identityPendingItems.length,
    hierarchy_review_product_ids: summary.hierarchyReview.map(row => row.productId),
    mapping_states: Object.fromEntries(Object.entries(summary.mappingStateCoverage).map(([key, value]) => [key, snakeMetric(value)])),
    reconciliation_status: summary.reconciliation.status,
    duplicate_proposed_stable_ids: summary.reconciliation.duplicateProposedStableIds,
    duplicate_proposed_stable_keys: summary.reconciliation.duplicateProposedStableKeys,
    fingerprints: summary.fingerprints,
  };
}

function snakeMetric(metric) {
  return {
    fact_count: metric.factCount,
    sales_nok: metric.salesNok,
    quantity: metric.quantity,
  };
}
