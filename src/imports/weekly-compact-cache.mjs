import {
  IDENTITY_PENDING,
  WEEKLY_IDENTITY_PREFLIGHT_VERSION,
  buildWeeklyIdentityPreflight,
  deriveWeeklyMappingContentFingerprint,
} from "./weekly-identity-preflight.mjs";
import {
  WEEKLY_SALES_PARSER_VERSION,
  buildWeeklyCorpusManifest,
} from "./weekly-sales-parser.mjs";
import { deterministicRestaurantScopeFingerprint } from "../metrics/reporting-group-sales-share.mjs";
import { matrixDisplayValue } from "../reporting/interactive-performance.mjs";

export const WEEKLY_COMPACT_CACHE_SCHEMA_VERSION = "0.3.0-weekly-compact-cache-v1";
export const WEEKLY_RPG_CACHE_GRAIN = Object.freeze([
  "CacheVersion",
  "SourcePeriodKey",
  "RestaurantID",
  "ReportingGroupID",
]);
export const WEEKLY_SCOPE_CACHE_GRAIN = Object.freeze([
  "CacheVersion",
  "SourcePeriodKey",
  "RestaurantID",
]);
export const WEEKLY_CACHE_STATES = Object.freeze([
  "Mapped",
  "Unmapped",
  IDENTITY_PENDING,
  "Conflict",
  "Inactive Target",
]);

const CACHE_STATE_FIELDS = Object.freeze({
  Mapped: "mapped",
  Unmapped: "unmapped",
  [IDENTITY_PENDING]: "identityPending",
  Conflict: "conflict",
  "Inactive Target": "inactiveTarget",
});

/**
 * Build one complete candidate cache in memory. No workbook, source file,
 * active cache, catalog, mapping, or fact table is mutated. Validation failure
 * throws before any candidate can be returned or activated.
 */
export function buildCandidateWeeklyCache({
  parsedReports,
  catalogs,
  acceptedIdentityRegistry,
  expectedMappingContentFingerprint,
  expectedIdentityPreflightFingerprint,
  existingVersionManifests = [],
}) {
  validateMappingInput({
    catalogs,
    expectedMappingContentFingerprint,
  });
  const corpus = buildWeeklyCorpusManifest(parsedReports);
  if (corpus.status !== "PASS") {
    fail("PUL-030C-001", "Weekly corpus manifest must pass before cache construction.");
  }
  const preflight = buildWeeklyIdentityPreflight({
    parsedReports,
    catalogs,
    acceptedIdentityRegistry,
  });
  if (preflight.reconciliation.status !== "PASS") {
    fail("PUL-030C-002", "Weekly Identity Preflight must reconcile before cache construction.");
  }
  if (preflight.mappingContentFingerprint !== expectedMappingContentFingerprint) {
    fail("PUL-030C-003", "Weekly Identity Preflight mapping content is stale.");
  }
  if (preflight.fingerprints.preflightFingerprint !== expectedIdentityPreflightFingerprint) {
    fail(
      "PUL-030C-004",
      `Weekly Identity Preflight fingerprint ${preflight.fingerprints.preflightFingerprint} differs from accepted ${expectedIdentityPreflightFingerprint}.`,
    );
  }

  const activeGroups = activeReportingGroups(catalogs.reportingGroups);
  const restaurants = buildRestaurantRegistry(catalogs, preflight);
  const reportingScopeRestaurantIds = restaurants
    .filter(row => row.performanceEligible === "Yes")
    .map(row => row.restaurantId)
    .sort(compareText);
  const excludedReportingScopeRestaurantIds = restaurants
    .filter(row => row.performanceEligible !== "Yes")
    .map(row => row.restaurantId)
    .sort(compareText);
  const cacheVersion = buildCacheVersion({
    corpusFingerprint: corpus.corpusFingerprint,
    preflightFingerprint: preflight.fingerprints.preflightFingerprint,
    mappingContentFingerprint: expectedMappingContentFingerprint,
    activeGroups,
    restaurants,
  });
  assertNoActiveVersionOverwrite(existingVersionManifests, cacheVersion);

  const periodManifest = buildPeriodManifest(parsedReports, cacheVersion);
  const assignmentsBySourceRow = uniqueIndex(
    preflight.rowAssignments,
    row => row.sourceRowId,
    "Identity Preflight SourceRowID",
  );
  const restaurantById = uniqueIndex(restaurants, row => row.restaurantId, "RestaurantID");
  const activeGroupById = uniqueIndex(activeGroups, row => row.reportingGroupId, "active ReportingGroupID");
  const scopeAccumulators = new Map();
  const rpgAccumulators = new Map();
  let sourceRowsSeen = 0;

  for (const report of parsedReports) {
    for (const sourceRow of report.rows) {
      sourceRowsSeen += 1;
      const assignment = assignmentsBySourceRow.get(sourceRow.sourceRowId);
      if (!assignment) {
        fail("PUL-030C-005", `Identity Preflight is missing source row ${sourceRow.sourceRowId}.`);
      }
      const restaurant = restaurantById.get(assignment.restaurantId);
      if (!restaurant) {
        fail(
          "PUL-030C-006",
          `Source row ${sourceRow.sourceRowId} has no stable RestaurantID for cache grain.`,
        );
      }
      if (!WEEKLY_CACHE_STATES.includes(assignment.mappingStatus)) {
        fail(
          "PUL-030C-007",
          `Source row ${sourceRow.sourceRowId} has unsupported mapping state ${assignment.mappingStatus}.`,
        );
      }

      const scopeKey = grainKey([sourceRow.sourcePeriodKey, assignment.restaurantId]);
      const scope = scopeAccumulators.get(scopeKey) ?? newScopeAccumulator({
        sourcePeriodKey: sourceRow.sourcePeriodKey,
        restaurantId: assignment.restaurantId,
        performanceEligible: restaurant.performanceEligible,
      });
      addMetric(scope.source, sourceRow);
      addMetric(scope.states[assignment.mappingStatus], sourceRow);
      scopeAccumulators.set(scopeKey, scope);

      if (assignment.mappingStatus !== "Mapped") continue;
      const reportingGroup = activeGroupById.get(assignment.effectiveReportingGroupId);
      if (!reportingGroup) {
        fail(
          "PUL-030C-008",
          `Mapped source row ${sourceRow.sourceRowId} targets unavailable ReportingGroupID ${assignment.effectiveReportingGroupId}.`,
        );
      }
      const rpgKey = grainKey([
        sourceRow.sourcePeriodKey,
        assignment.restaurantId,
        assignment.effectiveReportingGroupId,
      ]);
      const rpg = rpgAccumulators.get(rpgKey) ?? emptyMetric();
      addMetric(rpg, sourceRow);
      rpgAccumulators.set(rpgKey, rpg);
    }
  }

  if (sourceRowsSeen !== preflight.sourceTotals.factCount ||
      assignmentsBySourceRow.size !== sourceRowsSeen) {
    fail("PUL-030C-009", "Source rows and Identity Preflight assignments are not one-to-one.");
  }

  const scopeCacheRows = [...scopeAccumulators.values()]
    .map(value => finalizeScopeRow(value, cacheVersion))
    .sort(compareScopeRows);
  const weeklyRpgCacheRows = [];
  for (const scopeRow of scopeCacheRows) {
    for (const group of activeGroups) {
      const key = grainKey([
        scopeRow.sourcePeriodKey,
        scopeRow.restaurantId,
        group.reportingGroupId,
      ]);
      const metric = finalizeMetric(rpgAccumulators.get(key) ?? emptyMetric());
      weeklyRpgCacheRows.push({
        weeklyRpgCacheRowId: stableId("WRPG-", [
          cacheVersion,
          scopeRow.sourcePeriodKey,
          scopeRow.restaurantId,
          group.reportingGroupId,
        ]),
        cacheVersion,
        sourcePeriodKey: scopeRow.sourcePeriodKey,
        restaurantId: scopeRow.restaurantId,
        reportingGroupId: group.reportingGroupId,
        mappedFactCount: metric.factCount,
        mappedSalesNok: metric.salesNok,
        mappedQuantity: metric.quantity,
      });
    }
  }
  weeklyRpgCacheRows.sort(compareRpgRows);

  const validation = validateCandidateCacheRows({
    periodManifest,
    scopeCacheRows,
    weeklyRpgCacheRows,
    activeGroups,
    sourceTotals: preflight.sourceTotals,
    reportingScopeRestaurantIds,
    excludedReportingScopeRestaurantIds,
  });
  if (validation.status !== "PASS") {
    fail("PUL-030C-010", `Candidate weekly cache validation failed: ${validation.errors.join("; ")}`);
  }

  const cacheFingerprint = fingerprintWeeklyCacheRows({
    cacheVersion,
    periodManifest,
    scopeCacheRows,
    weeklyRpgCacheRows,
  });
  const versionManifest = {
    cacheVersion,
    cacheSchemaVersion: WEEKLY_COMPACT_CACHE_SCHEMA_VERSION,
    cacheStatus: "Candidate",
    activationState: "Not Active",
    validationStatus: validation.status,
    sourceSystemId: preflight.sourceSystemId,
    parserVersion: WEEKLY_SALES_PARSER_VERSION,
    identityContractVersion: WEEKLY_IDENTITY_PREFLIGHT_VERSION,
    sourceCorpusFingerprint: corpus.corpusFingerprint,
    identityPreflightFingerprint: preflight.fingerprints.preflightFingerprint,
    catalogFingerprint: preflight.fingerprints.catalogFingerprint,
    catalogContentFingerprint: preflight.fingerprints.catalogContentFingerprint,
    mappingContentFingerprint: expectedMappingContentFingerprint,
    mappingFingerprint: preflight.mappingFingerprint,
    mappingAsOfDate: preflight.catalogAsOfDate,
    activeReportingGroupFingerprint: fingerprintActiveGroups(activeGroups),
    performanceRestaurantScopeFingerprint: deterministicRestaurantScopeFingerprint(
      reportingScopeRestaurantIds.map(restaurantId => ({ restaurantId })),
    ),
    periodRowCount: periodManifest.length,
    scopeCacheRowCount: scopeCacheRows.length,
    denseRpgCacheRowCount: weeklyRpgCacheRows.length,
    nonzeroRpgCacheRowCount: weeklyRpgCacheRows.filter(isNonzeroRpgRow).length,
    cacheFingerprint,
  };

  return {
    schemaVersion: WEEKLY_COMPACT_CACHE_SCHEMA_VERSION,
    versionManifest,
    periodManifest,
    scopeCacheRows,
    weeklyRpgCacheRows,
    activeReportingGroups: activeGroups,
    reportingScopeRestaurantIds,
    excludedReportingScopeRestaurantIds,
    validation,
    identityPreflight: preflight,
  };
}

export function validateWeeklyCacheFreshness({
  versionManifest,
  currentMappingContentFingerprint,
}) {
  const errors = [];
  if (versionManifest.mappingContentFingerprint !== currentMappingContentFingerprint) {
    errors.push(
      `Candidate cache MappingContentFingerprint ${versionManifest.mappingContentFingerprint} differs from current ${currentMappingContentFingerprint}.`,
    );
  }
  return errors;
}

export function compareCandidateCacheRanges({
  cache,
  currentRange,
  comparisonRange,
  reportingGroupIds,
  restaurantIds = cache.reportingScopeRestaurantIds,
}) {
  const currentPeriodKeys = resolveRangePeriodKeys(cache.periodManifest, currentRange);
  const comparisonPeriodKeys = resolveRangePeriodKeys(cache.periodManifest, comparisonRange);
  const groupIds = reportingGroupIds?.length
    ? uniqueStrings(reportingGroupIds, "ReportingGroupID")
    : cache.activeReportingGroups.map(row => row.reportingGroupId);
  const selectedRestaurantIds = uniqueStrings(restaurantIds, "RestaurantID");
  const allowedRestaurants = new Set(cache.reportingScopeRestaurantIds);
  for (const restaurantId of selectedRestaurantIds) {
    if (!allowedRestaurants.has(restaurantId)) {
      fail("PUL-030C-011", `RestaurantID ${restaurantId} is outside normal Performance scope.`);
    }
  }
  const current = aggregateCacheRange({
    cache,
    periodKeys: currentPeriodKeys,
    restaurantIds: selectedRestaurantIds,
    reportingGroupIds: groupIds,
  });
  const comparison = aggregateCacheRange({
    cache,
    periodKeys: comparisonPeriodKeys,
    restaurantIds: selectedRestaurantIds,
    reportingGroupIds: groupIds,
  });
  const results = groupIds.map(reportingGroupId => {
    const currentGroup = current.byReportingGroup.get(reportingGroupId) ?? emptyMetric();
    const comparisonGroup = comparison.byReportingGroup.get(reportingGroupId) ?? emptyMetric();
    const cell = {
      current: canonicalComponents(currentGroup.salesNok, current.denominator.salesNok),
      comparison: canonicalComponents(comparisonGroup.salesNok, comparison.denominator.salesNok),
    };
    return {
      reportingGroupId,
      currentShare: matrixDisplayValue(cell, "Current Share"),
      comparisonShare: matrixDisplayValue(cell, "Comparison Share"),
      ppChange: matrixDisplayValue(cell, "PP Change"),
      currentSalesNok: matrixDisplayValue(cell, "Current Sales NOK"),
      nokImpact: matrixDisplayValue(cell, "NOK Impact"),
      currentMappedFactCount: currentGroup.factCount,
      currentMappedQuantity: currentGroup.quantity,
      comparisonMappedFactCount: comparisonGroup.factCount,
      comparisonMappedQuantity: comparisonGroup.quantity,
    };
  });
  return {
    currentRange: summarizeRange(currentRange, currentPeriodKeys),
    comparisonRange: summarizeRange(comparisonRange, comparisonPeriodKeys),
    restaurantCount: selectedRestaurantIds.length,
    currentDenominator: current.denominator,
    comparisonDenominator: comparison.denominator,
    results,
  };
}

export function estimateCandidateWorkbookFootprint(cache) {
  const periodColumns = columnCount(cache.periodManifest);
  const scopeColumns = columnCount(cache.scopeCacheRows);
  const rpgColumns = columnCount(cache.weeklyRpgCacheRows);
  const versionColumns = Object.keys(cache.versionManifest).length;
  const totalCells = periodColumns * cache.periodManifest.length +
    scopeColumns * cache.scopeCacheRows.length +
    rpgColumns * cache.weeklyRpgCacheRows.length +
    versionColumns;
  const serializedBytes = new TextEncoder().encode(JSON.stringify({
    versionManifest: cache.versionManifest,
    periodManifest: cache.periodManifest,
    scopeCacheRows: cache.scopeCacheRows,
    weeklyRpgCacheRows: cache.weeklyRpgCacheRows,
  })).length;
  return {
    periodColumns,
    scopeColumns,
    rpgColumns,
    versionColumns,
    totalCells,
    inMemoryJsonBytes: serializedBytes,
    estimatedAdditionalXlsxMegabytes: {
      low: roundNumber(serializedBytes * 0.18 / 1_000_000, 2),
      high: roundNumber(serializedBytes * 0.55 / 1_000_000, 2),
    },
  };
}

function validateMappingInput({ catalogs, expectedMappingContentFingerprint }) {
  const expectedFingerprint = requiredText(
    expectedMappingContentFingerprint,
    "expected MappingContentFingerprint",
  );
  const currentFingerprint = deriveWeeklyMappingContentFingerprint(catalogs);
  if (currentFingerprint !== expectedFingerprint) {
    fail(
      "PUL-030C-012",
      `Catalog MappingContentFingerprint ${currentFingerprint} differs from accepted ${expectedFingerprint}.`,
    );
  }
}

function activeReportingGroups(groups) {
  if (!Array.isArray(groups)) fail("PUL-030C-014", "Reporting Group catalog is missing.");
  const active = groups.filter(row => String(row.active) === "Yes").map(row => ({
    reportingGroupId: requiredText(row.reportingGroupId, "ReportingGroupID"),
    reportingGroupName: String(row.reportingGroupName ?? ""),
    sortOrder: Number(row.sortOrder),
  }));
  uniqueIndex(active, row => row.reportingGroupId, "active ReportingGroupID");
  if (active.length !== 9) {
    fail("PUL-030C-015", `Candidate weekly cache requires the accepted nine active Reporting Groups; found ${active.length}.`);
  }
  return active.sort((left, right) => left.sortOrder - right.sortOrder ||
    compareText(left.reportingGroupId, right.reportingGroupId));
}

function buildRestaurantRegistry(catalogs, preflight) {
  const current = catalogs.restaurants.map(row => ({
    restaurantId: row.restaurantId,
    sourceRestaurantName: row.sourceRestaurantName,
    status: row.status,
    reportingEnabled: row.reportingEnabled,
  }));
  const accepted = (preflight.acceptedIdentityRegistry?.restaurants ?? []).map(row => ({
    restaurantId: row.restaurantId,
    sourceRestaurantName: row.sourceRestaurantName,
    status: row.status,
    reportingEnabled: row.reportingEnabled,
  }));
  const candidates = preflight.newIdentityCandidates.restaurants.map(row => ({
    restaurantId: row.restaurantId,
    sourceRestaurantName: row.sourceRestaurantName,
    status: row.status,
    reportingEnabled: row.reportingEnabled,
  }));
  const result = [...current, ...accepted, ...candidates].map(row => ({
    ...row,
    performanceEligible: row.status === "Active" && row.reportingEnabled === "Yes" ? "Yes" : "No",
  }));
  uniqueIndex(result, row => row.restaurantId, "RestaurantID");
  uniqueIndex(result, row => `${catalogs.sourceSystemId} || ${row.sourceRestaurantName}`, "Restaurant stable key");
  for (const row of result.filter(value =>
    value.sourceRestaurantName === "Test Department (Not for User)" ||
    value.sourceRestaurantName === "Test Department (Not for Users)")) {
    if (row.reportingEnabled !== "No" || row.performanceEligible !== "No") {
      fail("PUL-030C-016", `${row.sourceRestaurantName} must remain ReportingEnabled=No.`);
    }
  }
  return result.sort((left, right) => compareText(left.restaurantId, right.restaurantId));
}

function buildCacheVersion({
  corpusFingerprint,
  preflightFingerprint,
  mappingContentFingerprint,
  activeGroups,
  restaurants,
}) {
  const records = [
    record("SCHEMA", [WEEKLY_COMPACT_CACHE_SCHEMA_VERSION]),
    record("CORPUS", [corpusFingerprint]),
    record("IDENTITY", [preflightFingerprint]),
    record("MAPPING_CONTENT", [mappingContentFingerprint]),
    ...activeGroups.map(row => record("RPG", [row.reportingGroupId])),
    ...restaurants.map(row => record("RESTAURANT", [
      row.restaurantId,
      row.sourceRestaurantName,
      row.status,
      row.reportingEnabled,
    ])),
  ].sort(compareText);
  return hashStrings(records, "WCV-");
}

function assertNoActiveVersionOverwrite(existing, cacheVersion) {
  if (!Array.isArray(existing)) fail("PUL-030C-017", "Existing cache-version manifest must be an array.");
  const active = existing.find(row => row.cacheVersion === cacheVersion &&
    (row.activationState === "Active" || row.cacheStatus === "Active"));
  if (active) {
    fail("PUL-030C-018", `Candidate build would overwrite active cache version ${cacheVersion}.`);
  }
}

function buildPeriodManifest(parsedReports, cacheVersion) {
  return parsedReports.map(report => ({
    weeklyPeriodManifestRowId: stableId("WPER-", [cacheVersion, report.manifest.sourcePeriodKey]),
    cacheVersion,
    sourcePeriodKey: report.manifest.sourcePeriodKey,
    periodStart: report.manifest.periodStart,
    periodEnd: report.manifest.periodEnd,
    isoYear: report.manifest.isoYear,
    isoWeek: report.manifest.isoWeek,
    sourceFileId: report.manifest.sourceFileId,
    sourceSemanticFingerprint: report.manifest.semanticFingerprint,
    sourceBinaryFingerprint: report.manifest.sourceBinaryFingerprint,
    scopeId: report.manifest.scopeId,
    scopeFingerprint: report.manifest.scopeFingerprint,
    sourceFactCount: report.manifest.sourceRowCount,
    sourceSalesNok: report.manifest.totalSalesNok,
    sourceQuantity: report.manifest.totalQuantity,
    sourceRestaurantCount: report.manifest.restaurantCount,
  })).sort((left, right) => compareText(left.periodStart, right.periodStart));
}

function newScopeAccumulator({ sourcePeriodKey, restaurantId, performanceEligible }) {
  return {
    sourcePeriodKey,
    restaurantId,
    performanceEligible,
    source: emptyMetric(),
    states: Object.fromEntries(WEEKLY_CACHE_STATES.map(state => [state, emptyMetric()])),
  };
}

function finalizeScopeRow(value, cacheVersion) {
  const source = finalizeMetric(value.source);
  const row = {
    weeklyScopeCacheRowId: stableId("WSCP-", [
      cacheVersion,
      value.sourcePeriodKey,
      value.restaurantId,
    ]),
    cacheVersion,
    sourcePeriodKey: value.sourcePeriodKey,
    restaurantId: value.restaurantId,
    performanceEligible: value.performanceEligible,
    sourceFactCount: source.factCount,
    sourceSalesNok: source.salesNok,
    sourceQuantity: source.quantity,
  };
  for (const state of WEEKLY_CACHE_STATES) {
    const prefix = CACHE_STATE_FIELDS[state];
    const metric = finalizeMetric(value.states[state]);
    row[`${prefix}FactCount`] = metric.factCount;
    row[`${prefix}SalesNok`] = metric.salesNok;
    row[`${prefix}Quantity`] = metric.quantity;
  }
  return row;
}

function validateCandidateCacheRows({
  periodManifest,
  scopeCacheRows,
  weeklyRpgCacheRows,
  activeGroups,
  sourceTotals,
  reportingScopeRestaurantIds,
  excludedReportingScopeRestaurantIds,
}) {
  const errors = [];
  collectDuplicateGrainErrors(scopeCacheRows, row => grainKey([
    row.cacheVersion,
    row.sourcePeriodKey,
    row.restaurantId,
  ]), "weekly scope cache", errors);
  collectDuplicateGrainErrors(weeklyRpgCacheRows, row => grainKey([
    row.cacheVersion,
    row.sourcePeriodKey,
    row.restaurantId,
    row.reportingGroupId,
  ]), "weekly RPG cache", errors);
  const expectedDenseRows = scopeCacheRows.length * activeGroups.length;
  if (weeklyRpgCacheRows.length !== expectedDenseRows) {
    errors.push(`Dense RPG row count is ${weeklyRpgCacheRows.length}; expected ${expectedDenseRows}.`);
  }
  for (const row of scopeCacheRows) {
    const states = metricFromScopeStates(row);
    const source = metricFromScopeSource(row);
    if (!metricsEqual(source, states)) {
      errors.push(`${row.sourcePeriodKey}/${row.restaurantId} state coverage does not reconcile.`);
    }
  }

  const periodReconciliation = periodManifest.map(period => reconcileScopeRows(
    scopeCacheRows.filter(row => row.sourcePeriodKey === period.sourcePeriodKey),
    {
      label: period.sourcePeriodKey,
      expected: {
        factCount: period.sourceFactCount,
        salesNok: period.sourceSalesNok,
        quantity: period.sourceQuantity,
      },
    },
  ));
  const yearReconciliation = [2025, 2026].map(isoYear => {
    const periodKeys = new Set(periodManifest.filter(row => row.isoYear === isoYear)
      .map(row => row.sourcePeriodKey));
    return reconcileScopeRows(scopeCacheRows.filter(row => periodKeys.has(row.sourcePeriodKey)), {
      label: String(isoYear),
      expected: sumPeriodManifest(periodManifest.filter(row => row.isoYear === isoYear)),
    });
  });
  const corpusReconciliation = reconcileScopeRows(scopeCacheRows, {
    label: "Complete 84-week corpus",
    expected: sourceTotals,
  });
  for (const result of [...periodReconciliation, ...yearReconciliation, corpusReconciliation]) {
    if (result.status !== "PASS") errors.push(`${result.label} reconciliation failed.`);
  }

  const mappedScope = sumScopeState(scopeCacheRows, "mapped");
  const mappedRpg = sumMetrics(weeklyRpgCacheRows.map(row => ({
    factCount: row.mappedFactCount,
    salesNok: row.mappedSalesNok,
    quantity: row.mappedQuantity,
  })));
  if (!metricsEqual(mappedScope, mappedRpg)) {
    errors.push("Dense RPG cache does not reconcile to Mapped scope coverage.");
  }

  const enabledSet = new Set(reportingScopeRestaurantIds);
  const excludedSet = new Set(excludedReportingScopeRestaurantIds);
  const reportingScope = sumMetrics(scopeCacheRows
    .filter(row => enabledSet.has(row.restaurantId)).map(metricFromScopeSource));
  const excludedReportingScope = sumMetrics(scopeCacheRows
    .filter(row => excludedSet.has(row.restaurantId)).map(metricFromScopeSource));
  if (!metricsEqual(sumMetrics([reportingScope, excludedReportingScope]), sourceTotals)) {
    errors.push("Performance-enabled plus excluded restaurant scope does not reconcile to source.");
  }

  return {
    status: errors.length ? "FAIL" : "PASS",
    errors,
    periodReconciliation,
    yearReconciliation,
    corpusReconciliation,
    mappedRpgReconciliation: {
      scopeMapped: mappedScope,
      rpgMapped: mappedRpg,
      status: metricsEqual(mappedScope, mappedRpg) ? "PASS" : "FAIL",
    },
    performanceScopeReconciliation: {
      enabled: reportingScope,
      excluded: excludedReportingScope,
      complete: sumMetrics([reportingScope, excludedReportingScope]),
      status: metricsEqual(sumMetrics([reportingScope, excludedReportingScope]), sourceTotals)
        ? "PASS"
        : "FAIL",
    },
  };
}

function reconcileScopeRows(rows, { label, expected }) {
  const source = sumMetrics(rows.map(metricFromScopeSource));
  const coverage = Object.fromEntries(WEEKLY_CACHE_STATES.map(state => [
    state,
    sumScopeState(rows, CACHE_STATE_FIELDS[state]),
  ]));
  const coverageTotal = sumMetrics(Object.values(coverage));
  return {
    label,
    source,
    coverage,
    coverageTotal,
    expected: finalizeMetric(expected),
    status: metricsEqual(source, expected) && metricsEqual(source, coverageTotal) ? "PASS" : "FAIL",
  };
}

function aggregateCacheRange({ cache, periodKeys, restaurantIds, reportingGroupIds }) {
  const periodSet = new Set(periodKeys);
  const restaurantSet = new Set(restaurantIds);
  const groupSet = new Set(reportingGroupIds);
  const scopeRows = cache.scopeCacheRows.filter(row =>
    periodSet.has(row.sourcePeriodKey) && restaurantSet.has(row.restaurantId));
  const denominator = sumMetrics(scopeRows.map(metricFromScopeSource));
  const byReportingGroup = new Map(reportingGroupIds.map(id => [id, emptyMetric()]));
  for (const row of cache.weeklyRpgCacheRows) {
    if (!periodSet.has(row.sourcePeriodKey) || !restaurantSet.has(row.restaurantId) ||
        !groupSet.has(row.reportingGroupId)) continue;
    const metric = byReportingGroup.get(row.reportingGroupId);
    metric.factCount += row.mappedFactCount;
    metric.salesNok += row.mappedSalesNok;
    metric.quantity += row.mappedQuantity;
  }
  byReportingGroup.forEach((value, key) => byReportingGroup.set(key, finalizeMetric(value)));
  return { denominator, byReportingGroup };
}

function resolveRangePeriodKeys(periodManifest, range) {
  const isoYear = Number(range?.isoYear);
  const weekStart = Number(range?.weekStart);
  const weekEnd = Number(range?.weekEnd);
  if (!Number.isInteger(isoYear) || !Number.isInteger(weekStart) ||
      !Number.isInteger(weekEnd) || weekStart < 1 || weekEnd < weekStart || weekEnd > 53) {
    fail("PUL-030C-019", "Range must contain a valid ISO year and inclusive weekStart/weekEnd.");
  }
  const rows = periodManifest.filter(row => row.isoYear === isoYear &&
    row.isoWeek >= weekStart && row.isoWeek <= weekEnd);
  const expectedCount = weekEnd - weekStart + 1;
  if (rows.length !== expectedCount) {
    fail(
      "PUL-030C-020",
      `Range ${isoYear} W${weekStart}-W${weekEnd} contains ${rows.length} periods; expected ${expectedCount}.`,
    );
  }
  return rows.sort((left, right) => left.isoWeek - right.isoWeek)
    .map(row => row.sourcePeriodKey);
}

function summarizeRange(range, periodKeys) {
  return {
    rangeId: `RNG-${range.isoYear}-W${String(range.weekStart).padStart(2, "0")}-W${String(range.weekEnd).padStart(2, "0")}`,
    isoYear: Number(range.isoYear),
    weekStart: Number(range.weekStart),
    weekEnd: Number(range.weekEnd),
    periodCount: periodKeys.length,
    sourcePeriodKeys: periodKeys,
  };
}

function canonicalComponents(numeratorSalesNok, denominatorSalesNok) {
  return {
    numeratorSalesNok,
    denominatorSalesNok,
    metricValue: denominatorSalesNok === 0 ? 0 : numeratorSalesNok / denominatorSalesNok,
  };
}

export function fingerprintWeeklyCacheRows({
  cacheVersion,
  periodManifest,
  scopeCacheRows,
  weeklyRpgCacheRows,
}) {
  const records = [record("CACHE_VERSION", [cacheVersion])];
  for (const row of periodManifest) {
    records.push(record("PERIOD", [
      row.sourcePeriodKey,
      row.periodStart,
      row.periodEnd,
      row.sourceFileId,
      row.sourceSemanticFingerprint,
      row.sourceFactCount,
      canonicalSales(row.sourceSalesNok),
      canonicalQuantity(row.sourceQuantity),
    ]));
  }
  for (const row of scopeCacheRows) {
    records.push(record("SCOPE", [
      row.sourcePeriodKey,
      row.restaurantId,
      row.performanceEligible,
      ...metricFields(metricFromScopeSource(row)),
      ...WEEKLY_CACHE_STATES.flatMap(state => metricFields(metricFromScopeState(
        row,
        CACHE_STATE_FIELDS[state],
      ))),
    ]));
  }
  for (const row of weeklyRpgCacheRows) {
    records.push(record("RPG", [
      row.sourcePeriodKey,
      row.restaurantId,
      row.reportingGroupId,
      row.mappedFactCount,
      canonicalSales(row.mappedSalesNok),
      canonicalQuantity(row.mappedQuantity),
    ]));
  }
  return hashStrings(records.sort(compareText), "WCC-");
}

function fingerprintActiveGroups(groups) {
  return hashStrings(groups.map(row => record("RPG", [row.reportingGroupId]))
    .sort(compareText), "RGS-");
}

function collectDuplicateGrainErrors(rows, keySelector, label, errors) {
  const seen = new Set();
  for (const row of rows) {
    const key = keySelector(row);
    if (seen.has(key)) errors.push(`${label} repeats grain ${key}.`);
    seen.add(key);
  }
}

function metricFromScopeSource(row) {
  return {
    factCount: row.sourceFactCount,
    salesNok: row.sourceSalesNok,
    quantity: row.sourceQuantity,
  };
}

function metricFromScopeState(row, prefix) {
  return {
    factCount: row[`${prefix}FactCount`],
    salesNok: row[`${prefix}SalesNok`],
    quantity: row[`${prefix}Quantity`],
  };
}

function metricFromScopeStates(row) {
  return sumMetrics(WEEKLY_CACHE_STATES.map(state =>
    metricFromScopeState(row, CACHE_STATE_FIELDS[state])));
}

function sumScopeState(rows, prefix) {
  return sumMetrics(rows.map(row => metricFromScopeState(row, prefix)));
}

function sumPeriodManifest(rows) {
  return sumMetrics(rows.map(row => ({
    factCount: row.sourceFactCount,
    salesNok: row.sourceSalesNok,
    quantity: row.sourceQuantity,
  })));
}

function newMetric() {
  return { factCount: 0, salesNok: 0, quantity: 0 };
}

function emptyMetric() {
  return newMetric();
}

function addMetric(metric, row) {
  metric.factCount += Number(row.factCount ?? row.sourceFactCount ?? 1);
  metric.salesNok += Number(row.salesNok ?? row.salesAmount ?? 0);
  metric.quantity += Number(row.quantity ?? 0);
}

function sumMetrics(metrics) {
  const result = emptyMetric();
  for (const metric of metrics) addMetric(result, metric);
  return finalizeMetric(result);
}

function finalizeMetric(metric) {
  return {
    factCount: Number(metric.factCount),
    salesNok: roundNumber(metric.salesNok, 2),
    quantity: roundNumber(metric.quantity, 6),
  };
}

function metricsEqual(left, right) {
  const a = finalizeMetric(left);
  const b = finalizeMetric(right);
  return a.factCount === b.factCount && almostEqual(a.salesNok, b.salesNok) &&
    almostEqual(a.quantity, b.quantity);
}

function metricFields(metric) {
  const value = finalizeMetric(metric);
  return [value.factCount, canonicalSales(value.salesNok), canonicalQuantity(value.quantity)];
}

function isNonzeroRpgRow(row) {
  return row.mappedFactCount !== 0 || row.mappedSalesNok !== 0 || row.mappedQuantity !== 0;
}

function uniqueIndex(rows, keySelector, label) {
  const result = new Map();
  for (const row of rows) {
    const key = requiredText(keySelector(row), label);
    if (result.has(key)) fail("PUL-030C-021", `${label} repeats ${key}.`);
    result.set(key, row);
  }
  return result;
}

function uniqueStrings(values, label) {
  const seen = new Set();
  const result = [];
  for (const raw of values) {
    const value = requiredText(raw, label);
    if (seen.has(value)) fail("PUL-030C-022", `${label} repeats ${value}.`);
    seen.add(value);
    result.push(value);
  }
  return result;
}

function columnCount(rows) {
  return rows.length ? Object.keys(rows[0]).length : 0;
}

function compareScopeRows(left, right) {
  return compareText(left.sourcePeriodKey, right.sourcePeriodKey) ||
    compareText(left.restaurantId, right.restaurantId);
}

function compareRpgRows(left, right) {
  return compareText(left.sourcePeriodKey, right.sourcePeriodKey) ||
    compareText(left.restaurantId, right.restaurantId) ||
    compareText(left.reportingGroupId, right.reportingGroupId);
}

function grainKey(values) {
  return values.join("\u001f");
}

function stableId(prefix, values) {
  return hashStrings([record("ID", values)], prefix);
}

function canonicalSales(value) {
  return Number(value).toFixed(2);
}

function canonicalQuantity(value) {
  return Number(value).toFixed(6);
}

function requiredText(value, label) {
  const result = String(value ?? "");
  if (!result) fail("PUL-030C-023", `${label} is required.`);
  return result;
}

function roundNumber(value, decimals) {
  const factor = 10 ** decimals;
  return Math.round((Number(value) + Number.EPSILON) * factor) / factor;
}

function almostEqual(left, right) {
  return Math.abs(Number(left) - Number(right)) <= 0.000001;
}

function record(kind, values) {
  return `${kind}|${values.map(value => {
    const normalized = value === null || value === undefined ? "" : String(value);
    return `${normalized.length}:${normalized}`;
  }).join("|")}`;
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

function compareText(left, right) {
  const a = String(left);
  const b = String(right);
  return a < b ? -1 : a > b ? 1 : 0;
}

function fail(code, message) {
  throw new Error(`${code}: ${message}`);
}
