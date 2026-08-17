#!/usr/bin/env node

/**
 * Read-only end-to-end audit for one New weekly publication.
 *
 * The caller supplies the exact Active corpus, incoming workbook, and catalog.
 * The command never searches outside those paths and never writes a source,
 * workbook, cache, or repository artifact.
 */

import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";
import {
  buildCandidateWeeklyCache,
  compareCandidateCacheRanges,
} from "./weekly-compact-cache.mjs";
import {
  acceptedIdentityRegistryFromPreflight,
  buildWeeklyIdentityPreflight,
  fingerprintAcceptedIdentityRegistry,
} from "./weekly-identity-preflight.mjs";
import {
  planWeeklyIntakePublication,
} from "./weekly-intake-publisher.mjs";
import {
  planWeeklyIntakePublishedActivation,
} from "./weekly-intake-publication-runtime.mjs";
import {
  listWeeklyXlsxFiles,
  readWeeklyWorkbookMatrix,
} from "./weekly-sales-ooxml.mjs";
import {
  WEEKLY_SALES_SHEET,
  parseWeeklySalesMatrix,
} from "./weekly-sales-parser.mjs";

const corpusRoot = exactPath(process.argv[2], "Active corpus path");
const incomingPath = exactPath(flag("--incoming"), "Incoming report path");
const catalogPath = exactPath(flag("--catalog"), "Catalog path");
const incomingResolved = path.resolve(incomingPath);
const corpusFiles = (await listWeeklyXlsxFiles(corpusRoot))
  .filter(file => path.resolve(file) !== incomingResolved);
if (!corpusFiles.length) throw new Error("Active corpus contains no .xlsx reports.");
const activeReports = [];
for (const file of corpusFiles) activeReports.push(await parseFile(file, corpusRoot));
const incoming = await parseFile(incomingPath, path.dirname(incomingPath));
const catalogs = JSON.parse(await readFile(catalogPath, "utf8"));

const activePreflight = buildWeeklyIdentityPreflight({ parsedReports: activeReports, catalogs });
const activeIdentityRegistry = acceptedIdentityRegistryFromPreflight(activePreflight);
const active = buildCandidateWeeklyCache({
  parsedReports: activeReports,
  catalogs,
  expectedMappingContentFingerprint: activePreflight.mappingContentFingerprint,
  expectedIdentityPreflightFingerprint: activePreflight.fingerprints.preflightFingerprint,
});
active.identityRegistry = activeIdentityRegistry;
active.versionManifest = {
  ...active.versionManifest,
  cacheStatus: "Active",
  activationState: "Active",
};
const currentFreshness = freshness(active.versionManifest);
const publication = planWeeklyIntakePublication({
  parsedReport: incoming,
  activeCache: active,
  versionManifests: [active.versionManifest],
  catalogs,
  currentFreshness,
  processedAt: "READ-ONLY-AUDIT",
});
if (publication.outcome !== "New" || !publication.candidatePrepared) {
  throw new Error(`Incoming report classified ${publication.outcome}; expected New. ${publication.ledgerEntry?.statusMessage ?? ""}`);
}
const activation = planWeeklyIntakePublishedActivation({
  publicationResult: publication,
  versionManifests: [active.versionManifest],
  currentFreshness,
  processedAt: "READ-ONLY-AUDIT",
});
const fullPreflight = buildWeeklyIdentityPreflight({
  parsedReports: [...activeReports, incoming], catalogs,
});
const full = buildCandidateWeeklyCache({
  parsedReports: [...activeReports, incoming],
  catalogs,
  expectedMappingContentFingerprint: fullPreflight.mappingContentFingerprint,
  expectedIdentityPreflightFingerprint: fullPreflight.fingerprints.preflightFingerprint,
});
const candidate = publication.candidate;
const incomingPreflight = candidate.identityPreflight;
const w33Comparison = compareCandidateCacheRanges({
  cache: candidate,
  currentRange: { isoYear: incoming.manifest.isoYear, weekStart: incoming.manifest.isoWeek,
    weekEnd: incoming.manifest.isoWeek },
  comparisonRange: { isoYear: incoming.manifest.isoYear, weekStart: incoming.manifest.isoWeek - 1,
    weekEnd: incoming.manifest.isoWeek - 1 },
});
const summary = {
  readOnly: true,
  status: candidate.validation.status,
  incoming: {
    period: incoming.manifest.sourcePeriodKey,
    periodStart: incoming.manifest.periodStart,
    periodEnd: incoming.manifest.periodEnd,
    sourceRowCount: incoming.manifest.sourceRowCount,
    salesNok: incoming.manifest.totalSalesNok,
    quantity: incoming.manifest.totalQuantity,
    semanticFingerprint: incoming.manifest.semanticFingerprint,
    newIdentityCandidates: {
      restaurants: incomingPreflight.newIdentityCandidates.restaurants.length,
      classifications: incomingPreflight.newIdentityCandidates.classifications.length,
      products: incomingPreflight.newIdentityCandidates.products.length,
    },
    newProducts: incomingPreflight.newIdentityCandidates.products.map(row => ({
      productId: row.productId,
      sourceProductName: row.sourceProductName,
      salesAccount: row.salesAccount,
      sourceClassificationId: row.sourceClassificationId,
      hierarchyStatus: row.hierarchyStatus,
      impact: row.impact,
    })),
    identityPending: incomingPreflight.identityStateCoverage["Identity Pending"],
    mappingStateCoverage: incomingPreflight.mappingStateCoverage,
  },
  prior: {
    cacheVersion: active.versionManifest.cacheVersion,
    cacheFingerprint: active.versionManifest.cacheFingerprint,
    periodRows: active.periodManifest.length,
    identityRegistryFingerprint: fingerprintAcceptedIdentityRegistry(activeIdentityRegistry),
    identityRegistryCounts: registryCounts(activeIdentityRegistry),
  },
  candidate: {
    cacheVersion: candidate.versionManifest.cacheVersion,
    cacheFingerprint: candidate.versionManifest.cacheFingerprint,
    sourceCorpusFingerprint: candidate.versionManifest.sourceCorpusFingerprint,
    identityPreflightFingerprint: candidate.versionManifest.identityPreflightFingerprint,
    periodRows: candidate.periodManifest.length,
    scopeRows: candidate.scopeCacheRows.length,
    rpgRows: candidate.weeklyRpgCacheRows.length,
    sourceTotals: candidate.validation.sourceTotals,
    mappingStateCoverage: candidate.validation.mappingStateCoverage,
    identityRegistryFingerprint: fingerprintAcceptedIdentityRegistry(candidate.identityRegistry),
    identityRegistryCounts: registryCounts(candidate.identityRegistry),
  },
  activation: activation.result,
  w33VsPriorWeek: {
    currentDenominator: w33Comparison.currentDenominator,
    comparisonDenominator: w33Comparison.comparisonDenominator,
    results: w33Comparison.results.filter(row =>
      ["RPG-0001", "RPG-0005", "RPG-0009"].includes(row.reportingGroupId)),
  },
  independentFullRebuild: {
    status: full.validation.status,
    businessRowsEqual: businessRowsEqual(candidate, full),
    sourceTotalsEqual: metricsEqual(candidate.validation.sourceTotals,
      full.validation.corpusReconciliation.source),
  },
};
process.stdout.write(`${JSON.stringify(summary, null, 2)}\n`);

async function parseFile(file, locatorRoot) {
  const bytes = await readFile(file);
  const workbook = readWeeklyWorkbookMatrix(file);
  if (workbook.sheetNames.length !== 1 || workbook.sheetNames[0] !== WEEKLY_SALES_SHEET) {
    throw new Error(`${file} must contain exactly one '${WEEKLY_SALES_SHEET}' sheet.`);
  }
  return parseWeeklySalesMatrix(workbook.matrix, {
    sourceLocator: path.relative(locatorRoot, file),
    sourceBinaryFingerprint: `SHA256-${createHash("sha256").update(bytes).digest("hex")}`,
  });
}

function freshness(version) {
  return {
    mappingContentFingerprint: version.mappingContentFingerprint,
    catalogContentFingerprint: version.catalogContentFingerprint,
    identityPreflightFingerprint: version.identityPreflightFingerprint,
    performanceRestaurantScopeFingerprint: version.performanceRestaurantScopeFingerprint,
  };
}

function businessRowsEqual(left, right) {
  const normalize = (rows, idField) => rows.map(row => {
    const { [idField]: ignoredId, cacheVersion: ignoredVersion, ...business } = row;
    return business;
  });
  return JSON.stringify(normalize(left.periodManifest, "weeklyPeriodManifestRowId")) ===
      JSON.stringify(normalize(right.periodManifest, "weeklyPeriodManifestRowId")) &&
    JSON.stringify(normalize(left.scopeCacheRows, "weeklyScopeCacheRowId")) ===
      JSON.stringify(normalize(right.scopeCacheRows, "weeklyScopeCacheRowId")) &&
    JSON.stringify(normalize(left.weeklyRpgCacheRows, "weeklyRpgCacheRowId")) ===
      JSON.stringify(normalize(right.weeklyRpgCacheRows, "weeklyRpgCacheRowId"));
}

function metricsEqual(left, right) {
  return Number(left.factCount) === Number(right.factCount) &&
    Math.abs(Number(left.salesNok) - Number(right.salesNok)) <= 0.005 &&
    Math.abs(Number(left.quantity) - Number(right.quantity)) <= 0.0000005;
}

function registryCounts(registry) {
  return {
    restaurants: registry.restaurants.length,
    classifications: registry.classifications.length,
    products: registry.products.length,
  };
}

function flag(name) {
  const index = process.argv.indexOf(name);
  const value = index < 0 ? "" : process.argv[index + 1];
  if (!value || value.startsWith("--")) throw new Error(`${name} requires one exact path.`);
  return value;
}

function exactPath(value, label) {
  if (!value || value.startsWith("--")) throw new Error(`${label} is required.`);
  return path.resolve(value);
}
