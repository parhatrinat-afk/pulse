/**
 * Deterministic Excel materialization contract for the inactive compact weekly
 * cache candidate. This module creates table-shaped values only; it does not
 * open or mutate a workbook.
 */

export const WEEKLY_CACHE_SHEET = "_Weekly_Cache";
export const WEEKLY_CACHE_STAGING_SHEET = "_Weekly_Cache_Staging";

export const WEEKLY_CACHE_TABLES = Object.freeze({
  version: "tblWeeklyCacheVersions",
  period: "tblWeeklyPeriodManifest",
  scope: "tblWeeklyScopeCache",
  rpg: "tblWeeklyRPGCache",
});

export const WEEKLY_CACHE_LAYOUT = Object.freeze({
  version: Object.freeze({ startRow: 1, startColumn: 1 }),
  period: Object.freeze({ startRow: 5, startColumn: 1 }),
  scope: Object.freeze({ startRow: 92, startColumn: 1 }),
  rpg: Object.freeze({ startRow: 1516, startColumn: 1 }),
});

export const WEEKLY_CACHE_COLUMNS = Object.freeze({
  version: Object.freeze([
    ["CacheVersion", "cacheVersion"],
    ["CacheSchemaVersion", "cacheSchemaVersion"],
    ["CacheStatus", "cacheStatus"],
    ["ActivationState", "activationState"],
    ["ValidationStatus", "validationStatus"],
    ["SourceSystemID", "sourceSystemId"],
    ["ParserVersion", "parserVersion"],
    ["IdentityContractVersion", "identityContractVersion"],
    ["SourceCorpusFingerprint", "sourceCorpusFingerprint"],
    ["IdentityPreflightFingerprint", "identityPreflightFingerprint"],
    ["CatalogFingerprint", "catalogFingerprint"],
    ["CatalogContentFingerprint", "catalogContentFingerprint"],
    ["MappingContentFingerprint", "mappingContentFingerprint"],
    ["Phase2AMappingFingerprint", "mappingFingerprint"],
    ["MappingAsOfDate", "mappingAsOfDate"],
    ["ActiveReportingGroupFingerprint", "activeReportingGroupFingerprint"],
    ["PerformanceRestaurantScopeFingerprint", "performanceRestaurantScopeFingerprint"],
    ["PeriodRowCount", "periodRowCount"],
    ["ScopeCacheRowCount", "scopeCacheRowCount"],
    ["DenseRPGCacheRowCount", "denseRpgCacheRowCount"],
    ["NonzeroRPGCacheRowCount", "nonzeroRpgCacheRowCount"],
    ["CacheFingerprint", "cacheFingerprint"],
  ]),
  period: Object.freeze([
    ["WeeklyPeriodManifestRowID", "weeklyPeriodManifestRowId"],
    ["CacheVersion", "cacheVersion"],
    ["SourcePeriodKey", "sourcePeriodKey"],
    ["PeriodStart", "periodStart"],
    ["PeriodEnd", "periodEnd"],
    ["ISOYear", "isoYear"],
    ["ISOWeek", "isoWeek"],
    ["SourceFileID", "sourceFileId"],
    ["SourceSemanticFingerprint", "sourceSemanticFingerprint"],
    ["SourceBinaryFingerprint", "sourceBinaryFingerprint"],
    ["ScopeID", "scopeId"],
    ["ScopeFingerprint", "scopeFingerprint"],
    ["SourceFactCount", "sourceFactCount"],
    ["SourceSalesNOK", "sourceSalesNok"],
    ["SourceQuantity", "sourceQuantity"],
    ["SourceRestaurantCount", "sourceRestaurantCount"],
  ]),
  scope: Object.freeze([
    ["WeeklyScopeCacheRowID", "weeklyScopeCacheRowId"],
    ["CacheVersion", "cacheVersion"],
    ["SourcePeriodKey", "sourcePeriodKey"],
    ["RestaurantID", "restaurantId"],
    ["PerformanceEligible", "performanceEligible"],
    ["SourceFactCount", "sourceFactCount"],
    ["SourceSalesNOK", "sourceSalesNok"],
    ["SourceQuantity", "sourceQuantity"],
    ["MappedFactCount", "mappedFactCount"],
    ["MappedSalesNOK", "mappedSalesNok"],
    ["MappedQuantity", "mappedQuantity"],
    ["UnmappedFactCount", "unmappedFactCount"],
    ["UnmappedSalesNOK", "unmappedSalesNok"],
    ["UnmappedQuantity", "unmappedQuantity"],
    ["IdentityPendingFactCount", "identityPendingFactCount"],
    ["IdentityPendingSalesNOK", "identityPendingSalesNok"],
    ["IdentityPendingQuantity", "identityPendingQuantity"],
    ["ConflictFactCount", "conflictFactCount"],
    ["ConflictSalesNOK", "conflictSalesNok"],
    ["ConflictQuantity", "conflictQuantity"],
    ["InactiveTargetFactCount", "inactiveTargetFactCount"],
    ["InactiveTargetSalesNOK", "inactiveTargetSalesNok"],
    ["InactiveTargetQuantity", "inactiveTargetQuantity"],
  ]),
  rpg: Object.freeze([
    ["WeeklyRPGCacheRowID", "weeklyRpgCacheRowId"],
    ["CacheVersion", "cacheVersion"],
    ["SourcePeriodKey", "sourcePeriodKey"],
    ["RestaurantID", "restaurantId"],
    ["ReportingGroupID", "reportingGroupId"],
    ["MappedFactCount", "mappedFactCount"],
    ["MappedSalesNOK", "mappedSalesNok"],
    ["MappedQuantity", "mappedQuantity"],
  ]),
});

export function buildWeeklyCacheMaterializationPlan(cache) {
  validateInactiveCandidate(cache);
  const sources = {
    version: [cache.versionManifest],
    period: cache.periodManifest,
    scope: cache.scopeCacheRows,
    rpg: cache.weeklyRpgCacheRows,
  };
  const sections = {};
  for (const name of Object.keys(WEEKLY_CACHE_COLUMNS)) {
    const columns = WEEKLY_CACHE_COLUMNS[name];
    const values = sources[name].map(row => columns.map(([, key]) => row[key] ?? ""));
    const layout = WEEKLY_CACHE_LAYOUT[name];
    sections[name] = {
      name,
      tableName: WEEKLY_CACHE_TABLES[name],
      startRow: layout.startRow,
      startColumn: layout.startColumn,
      headers: columns.map(([header]) => header),
      values,
      rowCount: values.length,
      columnCount: columns.length,
      address: rangeAddress(layout.startRow, layout.startColumn, values.length + 1, columns.length),
    };
  }
  validateNoOverlap(sections);
  return {
    sheetName: WEEKLY_CACHE_SHEET,
    stagingSheetName: WEEKLY_CACHE_STAGING_SHEET,
    cacheVersion: cache.versionManifest.cacheVersion,
    cacheFingerprint: cache.versionManifest.cacheFingerprint,
    mappingContentFingerprint: cache.versionManifest.mappingContentFingerprint,
    status: cache.versionManifest.cacheStatus,
    activationState: cache.versionManifest.activationState,
    sections,
    completeCandidateRows: Object.values(sections).reduce((sum, section) => sum + section.rowCount, 0),
  };
}

export function materializationChunk(plan, sectionName, offset = 0, limit = 500) {
  const section = plan.sections[sectionName];
  if (!section) throw new Error(`PUL-030M-005: Unknown materialization section ${sectionName}.`);
  if (!Number.isInteger(offset) || offset < 0 || offset > section.rowCount) {
    throw new Error(`PUL-030M-006: Invalid ${sectionName} row offset ${offset}.`);
  }
  if (!Number.isInteger(limit) || limit < 1) {
    throw new Error(`PUL-030M-007: Invalid ${sectionName} chunk size ${limit}.`);
  }
  const values = section.values.slice(offset, offset + limit);
  return {
    sheetName: plan.stagingSheetName,
    cacheVersion: plan.cacheVersion,
    cacheFingerprint: plan.cacheFingerprint,
    section: sectionName,
    tableName: section.tableName,
    headers: offset === 0 ? section.headers : [],
    offset,
    rowCount: values.length,
    totalRowCount: section.rowCount,
    startRow: section.startRow + 1 + offset,
    startColumn: section.startColumn,
    values,
  };
}

export function materializedSnapshotFromPlan(plan) {
  return {
    sheetName: plan.sheetName,
    cacheVersion: plan.cacheVersion,
    cacheFingerprint: plan.cacheFingerprint,
    status: plan.status,
    activationState: plan.activationState,
    completeCandidateRows: plan.completeCandidateRows,
    tables: Object.fromEntries(Object.entries(plan.sections).map(([name, section]) => [name, {
      tableName: section.tableName,
      address: section.address,
      headers: section.headers,
      rowCount: section.rowCount,
      columnCount: section.columnCount,
    }])),
  };
}

export function decideWeeklyCacheMaterialization(existingVersions, candidateVersion, candidateFingerprint) {
  const matching = existingVersions.filter(row => row.cacheVersion === candidateVersion);
  if (matching.some(row => row.activationState === "Active" || row.cacheStatus === "Active")) {
    throw new Error(`PUL-030M-008: Active cache version ${candidateVersion} cannot be overwritten.`);
  }
  if (!matching.length) return "Prepare";
  if (matching.length === 1 && matching[0].cacheFingerprint === candidateFingerprint &&
      matching[0].cacheStatus === "Candidate" && matching[0].activationState === "Not Active") {
    return "Already Materialized";
  }
  throw new Error(`PUL-030M-009: Existing cache version ${candidateVersion} differs from the accepted candidate.`);
}

function validateInactiveCandidate(cache) {
  const manifest = cache?.versionManifest;
  if (!manifest || manifest.cacheStatus !== "Candidate" || manifest.activationState !== "Not Active") {
    throw new Error("PUL-030M-001: Only a Candidate / Not Active cache can be materialized.");
  }
  if (manifest.validationStatus !== "PASS" || cache.validation?.status !== "PASS") {
    throw new Error("PUL-030M-002: Candidate validation must pass before materialization.");
  }
  const expected = {
    period: Number(manifest.periodRowCount),
    scope: Number(manifest.scopeCacheRowCount),
    rpg: Number(manifest.denseRpgCacheRowCount),
  };
  const actual = {
    period: cache.periodManifest?.length ?? -1,
    scope: cache.scopeCacheRows?.length ?? -1,
    rpg: cache.weeklyRpgCacheRows?.length ?? -1,
  };
  for (const name of Object.keys(expected)) {
    if (expected[name] !== actual[name]) {
      throw new Error(`PUL-030M-003: Candidate ${name} rows ${actual[name]}; expected ${expected[name]}.`);
    }
  }
}

function validateNoOverlap(sections) {
  const values = Object.values(sections).sort((left, right) => left.startRow - right.startRow);
  for (let index = 1; index < values.length; index += 1) {
    const previousEnd = values[index - 1].startRow + values[index - 1].rowCount;
    if (values[index].startRow <= previousEnd) {
      throw new Error(`PUL-030M-004: ${values[index - 1].name} overlaps ${values[index].name}.`);
    }
  }
}

function rangeAddress(startRow, startColumn, rowCount, columnCount) {
  const endRow = startRow + rowCount - 1;
  const endColumn = startColumn + columnCount - 1;
  return `${columnName(startColumn)}${startRow}:${columnName(endColumn)}${endRow}`;
}

function columnName(oneBasedColumn) {
  let value = oneBasedColumn;
  let result = "";
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + value % 26) + result;
    value = Math.floor(value / 26);
  }
  return result;
}
