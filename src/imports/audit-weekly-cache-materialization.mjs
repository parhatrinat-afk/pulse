#!/usr/bin/env node

/**
 * Read-only payload generator and validator for the accepted inactive weekly
 * cache materialization. It requires explicit paths and accepted fingerprints;
 * it never searches for files or writes a workbook/artifact.
 */

import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";
import {
  buildCandidateWeeklyCache,
  fingerprintWeeklyCacheRows,
} from "./weekly-compact-cache.mjs";
import {
  buildWeeklyCacheMaterializationPlan,
  materializationChunk,
  materializedSnapshotFromPlan,
} from "./weekly-cache-materialization.mjs";
import { listWeeklyXlsxFiles, readWeeklyWorkbookMatrix } from "./weekly-sales-ooxml.mjs";
import { WEEKLY_SALES_SHEET, parseWeeklySalesMatrix } from "./weekly-sales-parser.mjs";

const root = requiredPositionalPath(process.argv[2]);
const catalogPath = requiredFlagPath("--catalog");
const expectedMappingContentFingerprint = requiredFlagValue("--mapping-content-fingerprint");
const expectedPreflightFingerprint = requiredFlagValue("--preflight-fingerprint");
const expectedCacheVersion = requiredFlagValue("--cache-version");
const expectedCacheFingerprint = requiredFlagValue("--cache-fingerprint");
const sectionName = optionalFlagValue("--section");
const offset = optionalIntegerFlag("--offset", 0);
const limit = optionalIntegerFlag("--limit", 500);

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
if (cache.versionManifest.cacheVersion !== expectedCacheVersion) {
  throw new Error(
    `PUL-030M-010: CacheVersion ${cache.versionManifest.cacheVersion} differs from accepted ${expectedCacheVersion}.`,
  );
}
if (cache.versionManifest.cacheFingerprint !== expectedCacheFingerprint) {
  throw new Error(
    `PUL-030M-011: CacheFingerprint ${cache.versionManifest.cacheFingerprint} differs from accepted ${expectedCacheFingerprint}.`,
  );
}
const plan = buildWeeklyCacheMaterializationPlan(cache);
const output = sectionName
  ? materializationChunk(plan, sectionName, offset, limit)
  : {
      readOnly: true,
      materialization: materializedSnapshotFromPlan(plan),
      source: {
        files: files.length,
        factCount: cache.validation.corpusReconciliation.source.factCount,
        salesNok: cache.validation.corpusReconciliation.source.salesNok,
        quantity: cache.validation.corpusReconciliation.source.quantity,
      },
      identityPending: cache.identityPreflight.identityStateCoverage["Identity Pending"],
      performanceScope: cache.validation.performanceScopeReconciliation,
      excelCoercedPeriodFingerprint: fingerprintWeeklyCacheRows({
        cacheVersion: cache.versionManifest.cacheVersion,
        periodManifest: cache.periodManifest.map(row => ({
          ...row,
          periodStart: isoDateToExcelSerial(row.periodStart),
          periodEnd: isoDateToExcelSerial(row.periodEnd),
        })),
        scopeCacheRows: cache.scopeCacheRows,
        weeklyRpgCacheRows: cache.weeklyRpgCacheRows,
      }),
    };
process.stdout.write(`${JSON.stringify(output)}\n`);

function requiredPositionalPath(value) {
  if (!value || value.startsWith("--")) {
    throw new Error("Provide one exact fixture corpus path. The materializer will not search for files.");
  }
  return value;
}

function requiredFlagPath(flag) {
  return path.resolve(requiredFlagValue(flag));
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

function optionalIntegerFlag(flag, fallback) {
  const value = optionalFlagValue(flag);
  if (!value) return fallback;
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 0) throw new Error(`${flag} requires a non-negative integer.`);
  return parsed;
}

function isoDateToExcelSerial(value) {
  const milliseconds = Date.parse(`${value}T00:00:00Z`);
  if (!Number.isFinite(milliseconds)) throw new Error(`Invalid ISO date ${value}.`);
  return Math.round(milliseconds / 86400000 + 25569);
}
