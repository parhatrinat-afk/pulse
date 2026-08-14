#!/usr/bin/env node

/**
 * Read-only development build/validation for the candidate compact weekly
 * cache. It requires exact caller-supplied inputs and writes no source,
 * workbook, active cache, or repository artifact.
 */

import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";
import {
  WEEKLY_RPG_CACHE_GRAIN,
  WEEKLY_SCOPE_CACHE_GRAIN,
  buildCandidateWeeklyCache,
  compareCandidateCacheRanges,
  estimateCandidateWorkbookFootprint,
} from "./weekly-compact-cache.mjs";
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
const expectedMappingContentFingerprint = requiredFlagValue("--mapping-content-fingerprint");
const expectedPreflightFingerprint = requiredFlagValue("--preflight-fingerprint");
const expectedPath = optionalFlagValue("--expected");
const checkpointOnly = process.argv.includes("--checkpoint-only");
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
const cache = buildCandidateWeeklyCache({
  parsedReports,
  catalogs,
  expectedMappingContentFingerprint,
  expectedIdentityPreflightFingerprint: expectedPreflightFingerprint,
});
const comparisons = [
  compare("W31 2026 vs W31 2025", 2026, 31, 31, 2025, 31, 31),
  compare("W01-W32 2026 vs W01-W32 2025", 2026, 1, 32, 2025, 1, 32),
  compare("W20-W30 2026 vs W20-W30 2025", 2026, 20, 30, 2025, 20, 30),
];
const footprint = estimateCandidateWorkbookFootprint(cache);
const summary = {
  readOnly: true,
  status: cache.validation.status,
  schemas: {
    weeklyRpgCacheGrain: WEEKLY_RPG_CACHE_GRAIN,
    weeklyScopeCacheGrain: WEEKLY_SCOPE_CACHE_GRAIN,
  },
  versionManifest: cache.versionManifest,
  counts: {
    periodRows: cache.periodManifest.length,
    cacheVersionRows: 1,
    scopeCacheRows: cache.scopeCacheRows.length,
    denseRpgCacheRows: cache.weeklyRpgCacheRows.length,
    nonzeroRpgCacheRows: cache.versionManifest.nonzeroRpgCacheRowCount,
    analyticalRows: cache.scopeCacheRows.length + cache.weeklyRpgCacheRows.length,
    completeCandidateRows: 1 + cache.periodManifest.length +
      cache.scopeCacheRows.length + cache.weeklyRpgCacheRows.length,
  },
  reconciliation: {
    periodStatusCounts: countStatuses(cache.validation.periodReconciliation),
    years: cache.validation.yearReconciliation,
    corpus: cache.validation.corpusReconciliation,
    mappedRpg: cache.validation.mappedRpgReconciliation,
    performanceScope: cache.validation.performanceScopeReconciliation,
  },
  identityPending: cache.identityPreflight.identityStateCoverage["Identity Pending"],
  hierarchyReviewProductIds: cache.identityPreflight.hierarchyReview.map(row => row.productId),
  reportingScopeRestaurantIds: cache.reportingScopeRestaurantIds,
  excludedReportingScopeRestaurantIds: cache.excludedReportingScopeRestaurantIds,
  comparisons,
  footprint,
};
const checkpoint = buildCheckpoint(summary);
if (expectedPath) {
  const expected = JSON.parse(await readFile(path.resolve(expectedPath), "utf8"));
  if (JSON.stringify(checkpoint) !== JSON.stringify(expected)) {
    throw new Error(
      `PUL-030C-024: Weekly cache checkpoint differs.\nExpected ${JSON.stringify(expected)}\nActual ${JSON.stringify(checkpoint)}`,
    );
  }
  summary.expectedCheckpointStatus = "PASS";
}
process.stdout.write(`${JSON.stringify(checkpointOnly ? checkpoint : summary, null, 2)}\n`);

function compare(label, currentYear, currentStart, currentEnd, comparisonYear, comparisonStart, comparisonEnd) {
  const value = compareCandidateCacheRanges({
    cache,
    currentRange: { isoYear: currentYear, weekStart: currentStart, weekEnd: currentEnd },
    comparisonRange: {
      isoYear: comparisonYear,
      weekStart: comparisonStart,
      weekEnd: comparisonEnd,
    },
  });
  return { label, ...value };
}

function requiredPositionalPath(value) {
  if (!value || value.startsWith("--")) {
    throw new Error("Provide one exact fixture corpus path. The cache build will not search for source files.");
  }
  return value;
}

function requiredFlagPath(flag) {
  const value = requiredFlagValue(flag);
  return path.resolve(value);
}

function requiredFlagValue(flag) {
  const value = optionalFlagValue(flag);
  if (!value) throw new Error(`${flag} requires one exact value.`);
  return value;
}

function optionalFlagValue(flag) {
  const index = process.argv.indexOf(flag);
  if (index < 0) return "";
  const value = process.argv[index + 1];
  if (!value || value.startsWith("--")) throw new Error(`${flag} requires one exact value.`);
  return value;
}

function countStatuses(rows) {
  return rows.reduce((result, row) => {
    result[row.status] = (result[row.status] ?? 0) + 1;
    return result;
  }, {});
}

function buildCheckpoint(summary) {
  return {
    schema_version: summary.versionManifest.cacheSchemaVersion,
    status: summary.status,
    cache_version: summary.versionManifest.cacheVersion,
    cache_fingerprint: summary.versionManifest.cacheFingerprint,
    mapping_content_fingerprint: summary.versionManifest.mappingContentFingerprint,
    mapping_fingerprint: summary.versionManifest.mappingFingerprint,
    mapping_as_of_date: summary.versionManifest.mappingAsOfDate,
    source_corpus_fingerprint: summary.versionManifest.sourceCorpusFingerprint,
    identity_preflight_fingerprint: summary.versionManifest.identityPreflightFingerprint,
    performance_restaurant_scope_fingerprint:
      summary.versionManifest.performanceRestaurantScopeFingerprint,
    counts: summary.counts,
    period_status_counts: summary.reconciliation.periodStatusCounts,
    years: summary.reconciliation.years.map(row => ({
      label: row.label,
      source: snakeMetric(row.source),
      coverage: snakeCoverage(row.coverage),
      status: row.status,
    })),
    corpus: {
      source: snakeMetric(summary.reconciliation.corpus.source),
      coverage: snakeCoverage(summary.reconciliation.corpus.coverage),
      status: summary.reconciliation.corpus.status,
    },
    mapped_rpg_status: summary.reconciliation.mappedRpg.status,
    performance_scope: {
      enabled: snakeMetric(summary.reconciliation.performanceScope.enabled),
      excluded: snakeMetric(summary.reconciliation.performanceScope.excluded),
      status: summary.reconciliation.performanceScope.status,
    },
    identity_pending: snakeMetric(summary.identityPending),
    hierarchy_review_product_ids: summary.hierarchyReviewProductIds,
    reporting_scope_restaurant_ids: summary.reportingScopeRestaurantIds,
    excluded_reporting_scope_restaurant_ids: summary.excludedReportingScopeRestaurantIds,
    comparisons: summary.comparisons.map(value => ({
      label: value.label,
      current_denominator: snakeMetric(value.currentDenominator),
      comparison_denominator: snakeMetric(value.comparisonDenominator),
      results: value.results.filter(row => ["RPG-0001", "RPG-0005", "RPG-0009"]
        .includes(row.reportingGroupId)).map(row => ({
        reporting_group_id: row.reportingGroupId,
        current_share: row.currentShare,
        comparison_share: row.comparisonShare,
        pp_change: row.ppChange,
        current_sales_nok: row.currentSalesNok,
        nok_impact: row.nokImpact,
      })),
    })),
    footprint: summary.footprint,
  };
}

function snakeCoverage(coverage) {
  return Object.fromEntries(Object.entries(coverage).map(([state, metric]) => [state, snakeMetric(metric)]));
}

function snakeMetric(metric) {
  return {
    fact_count: metric.factCount,
    sales_nok: metric.salesNok,
    quantity: metric.quantity,
  };
}
