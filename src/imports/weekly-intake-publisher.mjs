import {
  WEEKLY_CACHE_ACTIVE_STATUS,
  validateActiveWeeklyCacheFreshness,
} from "./weekly-cache-activation.mjs";
import {
  WEEKLY_CACHE_STATES,
  WEEKLY_COMPACT_CACHE_SCHEMA_VERSION,
  buildCandidateWeeklyCache,
  fingerprintWeeklyCacheRows,
} from "./weekly-compact-cache.mjs";
import {
  WEEKLY_SALES_PARSER_VERSION,
  WEEKLY_SALES_SCHEMA_VERSION,
} from "./weekly-sales-parser.mjs";
import { buildWeeklyIdentityPreflight } from "./weekly-identity-preflight.mjs";

export const WEEKLY_INTAKE_PUBLISHER_VERSION = "0.3.0-weekly-intake-publisher-v1";

export const WEEKLY_INTAKE_OUTCOMES = Object.freeze([
  "New",
  "Duplicate",
  "Correction Review",
  "Rejected",
  "Cache Stale",
]);

export const WEEKLY_INTAKE_LOG_COLUMNS = Object.freeze([
  ["IntakeEventID", "intakeEventId"],
  ["SourceLocator", "sourceLocator"],
  ["SourceFileID", "sourceFileId"],
  ["SourcePeriodKey", "sourcePeriodKey"],
  ["SourceSemanticFingerprint", "sourceSemanticFingerprint"],
  ["IdentityPreflightFingerprint", "identityPreflightFingerprint"],
  ["IntakeStatus", "intakeStatus"],
  ["StatusMessage", "statusMessage"],
  ["SourceRowCount", "sourceRowCount"],
  ["SourceSalesNOK", "sourceSalesNok"],
  ["ProcessedAt", "processedAt"],
  ["PriorCacheVersion", "priorCacheVersion"],
  ["ResultingCacheVersion", "resultingCacheVersion"],
  ["SupersededCacheVersion", "supersededCacheVersion"],
]);

export const WEEKLY_FULL_VERSION_RETENTION = Object.freeze({
  maximumMaterializedFullVersions: 2,
  manifestRetention: "All version, period, source and intake manifests",
  analyticalRetention: "Active version plus one previous full rollback version",
});

const CACHE_STATE_FIELDS = Object.freeze({
  Mapped: "mapped",
  Unmapped: "unmapped",
  "Identity Pending": "identityPending",
  Conflict: "conflict",
  "Inactive Target": "inactiveTarget",
});

/**
 * Classify one already-parsed weekly report and, only for a fresh new period,
 * construct one complete inactive Candidate cache version. The active cache is
 * never mutated. Power Automate is intentionally outside this contract.
 */
export function planWeeklyIntakePublication({
  parsedReport,
  parseError = "",
  activeCache,
  versionManifests,
  catalogs,
  currentFreshness,
  processedAt = "",
  sourceLocator: explicitSourceLocator = "",
}) {
  const sourceLocator = String(parsedReport?.manifest?.sourceLocator ?? explicitSourceLocator);
  if (parseError) {
    return noChangeResult("Rejected", String(parseError), {
      sourceLocator,
      processedAt,
    });
  }
  const parserErrors = validateParsedReport(parsedReport);
  if (parserErrors.length) {
    return noChangeResult("Rejected", parserErrors.join(" "), {
      parsedReport,
      sourceLocator,
      processedAt,
    });
  }

  const freshness = validateActiveWeeklyCacheFreshness({
    versionManifests,
    current: currentFreshness,
  });
  if (freshness.status !== "Available") {
    return noChangeResult("Cache Stale", freshness.errors.join(" "), {
      parsedReport,
      sourceLocator,
      processedAt,
      freshness,
    });
  }
  let active;
  try {
    active = requireActiveCacheSnapshot(activeCache, freshness.activeVersion);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return noChangeResult("Cache Stale", message, {
      parsedReport,
      sourceLocator,
      processedAt,
      freshness,
    });
  }
  const matchingPeriod = active.periodManifest.find(row =>
    row.sourcePeriodKey === parsedReport.manifest.sourcePeriodKey);
  if (matchingPeriod) {
    if (matchingPeriod.sourceSemanticFingerprint === parsedReport.manifest.semanticFingerprint) {
      return noChangeResult("Duplicate", "Same period and semantic fingerprint; no cache change.", {
        parsedReport,
        sourceLocator,
        processedAt,
        freshness,
        priorCacheVersion: active.versionManifest.cacheVersion,
        identityPreflightFingerprint: active.versionManifest.identityPreflightFingerprint,
      });
    }
    return noChangeResult(
      "Correction Review",
      `Period ${parsedReport.manifest.sourcePeriodKey} already exists with different content; explicit supersession is required.`,
      {
        parsedReport,
        sourceLocator,
        processedAt,
        freshness,
        priorCacheVersion: active.versionManifest.cacheVersion,
        identityPreflightFingerprint: active.versionManifest.identityPreflightFingerprint,
      },
    );
  }

  let week;
  try {
    week = buildOneWeekSlice(parsedReport, catalogs, currentFreshness);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    const outcome = /stale|Fingerprint .* differs|MappingContentFingerprint/i.test(message)
      ? "Cache Stale"
      : "Rejected";
    return noChangeResult(outcome, message, {
      parsedReport,
      sourceLocator,
      processedAt,
      freshness,
      priorCacheVersion: active.versionManifest.cacheVersion,
    });
  }
  if (week.versionManifest.performanceRestaurantScopeFingerprint !==
      active.versionManifest.performanceRestaurantScopeFingerprint) {
    return noChangeResult(
      "Cache Stale",
      "Incoming identity state changes the ReportingEnabled restaurant scope; rebuild is required.",
      {
        parsedReport,
        sourceLocator,
        processedAt,
        freshness,
        priorCacheVersion: active.versionManifest.cacheVersion,
      },
    );
  }

  const candidate = combineActiveCacheAndNewWeek({
    active,
    week,
    versionManifests,
  });
  const ledgerEntry = buildIntakeLogEntry({
    outcome: "New",
    message: `Validated Candidate ${candidate.versionManifest.cacheVersion} prepared; not active.`,
    parsedReport,
    processedAt,
    priorCacheVersion: active.versionManifest.cacheVersion,
    resultingCacheVersion: candidate.versionManifest.cacheVersion,
    identityPreflightFingerprint: candidate.versionManifest.identityPreflightFingerprint,
  });
  return {
    outcome: "New",
    cacheChanged: false,
    candidatePrepared: true,
    activeCacheVersion: active.versionManifest.cacheVersion,
    candidate,
    freshness,
    identityPendingAccepted: candidate.identityPreflight.identityStateCoverage["Identity Pending"].factCount > 0,
    ledgerEntry,
  };
}

/**
 * Return the bounded full-row retention decision without deleting anything.
 * Manifests and source evidence are never cleanup candidates.
 */
export function planWeeklyFullVersionRetention({ versionManifests, candidateVersion }) {
  const versions = Array.isArray(versionManifests) ? versionManifests : [];
  const active = versions.filter(row => row.cacheStatus === WEEKLY_CACHE_ACTIVE_STATUS &&
    row.activationState === WEEKLY_CACHE_ACTIVE_STATUS);
  if (active.length !== 1) {
    fail("PUL-030WIP-030", `Retention requires exactly one Active / Active cache; found ${active.length}.`);
  }
  const candidate = requiredText(candidateVersion, "candidate CacheVersion");
  const retainFullVersions = [active[0].cacheVersion, candidate];
  const eligibleForAnalyticalRowCleanup = versions
    .map(row => row.cacheVersion)
    .filter(version => version && !retainFullVersions.includes(version))
    .sort(compareText);
  return {
    maximumMaterializedFullVersions: WEEKLY_FULL_VERSION_RETENTION.maximumMaterializedFullVersions,
    retainFullVersions,
    eligibleForAnalyticalRowCleanup,
    cleanupAction: "None in this slice",
    retainAllManifests: true,
  };
}

function buildOneWeekSlice(parsedReport, catalogs, currentFreshness) {
  const preliminary = buildCandidateWeeklyCache({
    parsedReports: [parsedReport],
    catalogs,
    expectedMappingContentFingerprint: currentFreshness.mappingContentFingerprint,
    expectedIdentityPreflightFingerprint: buildExpectedOneWeekIdentity(parsedReport, catalogs),
    existingVersionManifests: [],
  });
  if (preliminary.versionManifest.catalogContentFingerprint !==
      currentFreshness.catalogContentFingerprint) {
    fail(
      "PUL-030WIP-011",
      `CatalogContentFingerprint ${preliminary.versionManifest.catalogContentFingerprint} differs from current ${currentFreshness.catalogContentFingerprint}.`,
    );
  }
  return preliminary;
}

function buildExpectedOneWeekIdentity(parsedReport, catalogs) {
  // buildCandidateWeeklyCache performs the full preflight again. The first
  // pass is deliberately local and read-only so the accepted fingerprint is
  // an explicit input to its stale gate.
  const active = buildWeeklyIdentityPreflight({
    parsedReports: [parsedReport],
    catalogs,
  });
  return active.fingerprints.preflightFingerprint;
}

function combineActiveCacheAndNewWeek({ active, week, versionManifests }) {
  const combinedPeriodBase = [...active.periodManifest, ...week.periodManifest]
    .sort(comparePeriodRows);
  const sourceCorpusFingerprint = fingerprintCorpusManifest(combinedPeriodBase);
  const identityEvidence = normalizedIdentityEvidence(active);
  identityEvidence.push({
    evidenceKey: week.periodManifest[0].sourcePeriodKey,
    identityPreflightFingerprint: week.versionManifest.identityPreflightFingerprint,
  });
  identityEvidence.sort((left, right) => compareText(left.evidenceKey, right.evidenceKey));
  const identityPreflightFingerprint = hashStrings([
    record("CONTRACT", [WEEKLY_INTAKE_PUBLISHER_VERSION]),
    record("CORPUS", [sourceCorpusFingerprint]),
    ...identityEvidence.map(row => record("IDENTITY_EVIDENCE", [
      row.evidenceKey,
      row.identityPreflightFingerprint,
    ])),
  ], "IDP-");
  const reportingGroupIds = week.activeReportingGroups
    .map(row => row.reportingGroupId).sort(compareText);
  const cacheVersion = hashStrings([
    record("CONTRACT", [WEEKLY_INTAKE_PUBLISHER_VERSION]),
    record("SCHEMA", [WEEKLY_COMPACT_CACHE_SCHEMA_VERSION]),
    record("CORPUS", [sourceCorpusFingerprint]),
    record("IDENTITY", [identityPreflightFingerprint]),
    record("CATALOG", [week.versionManifest.catalogContentFingerprint]),
    record("MAPPING_CONTENT", [week.versionManifest.mappingContentFingerprint]),
    record("RESTAURANT_SCOPE", [week.versionManifest.performanceRestaurantScopeFingerprint]),
    ...reportingGroupIds.map(id => record("RPG", [id])),
  ], "WCV-");
  if ((versionManifests ?? []).some(row => row.cacheVersion === cacheVersion &&
      (row.cacheStatus === WEEKLY_CACHE_ACTIVE_STATUS ||
       row.activationState === WEEKLY_CACHE_ACTIVE_STATUS))) {
    fail("PUL-030WIP-014", `Candidate would overwrite Active cache ${cacheVersion}.`);
  }

  const periodManifest = combinedPeriodBase.map(row => rekeyPeriod(row, cacheVersion));
  const scopeCacheRows = [...active.scopeCacheRows, ...week.scopeCacheRows]
    .map(row => rekeyScope(row, cacheVersion)).sort(compareScopeRows);
  const weeklyRpgCacheRows = [...active.weeklyRpgCacheRows, ...week.weeklyRpgCacheRows]
    .map(row => rekeyRpg(row, cacheVersion)).sort(compareRpgRows);
  const validation = validateCompleteCacheRows({
    periodManifest,
    scopeCacheRows,
    weeklyRpgCacheRows,
    activeReportingGroups: week.activeReportingGroups,
  });
  if (validation.status !== "PASS") {
    fail("PUL-030WIP-015", `Complete Candidate validation failed: ${validation.errors.join("; ")}`);
  }
  const cacheFingerprint = fingerprintWeeklyCacheRows({
    cacheVersion,
    periodManifest,
    scopeCacheRows,
    weeklyRpgCacheRows,
  });
  const versionManifest = {
    ...week.versionManifest,
    cacheVersion,
    cacheStatus: "Candidate",
    activationState: "Not Active",
    validationStatus: "PASS",
    sourceCorpusFingerprint,
    identityPreflightFingerprint,
    periodRowCount: periodManifest.length,
    scopeCacheRowCount: scopeCacheRows.length,
    denseRpgCacheRowCount: weeklyRpgCacheRows.length,
    nonzeroRpgCacheRowCount: weeklyRpgCacheRows.filter(isNonzeroRpg).length,
    cacheFingerprint,
  };
  return {
    schemaVersion: WEEKLY_COMPACT_CACHE_SCHEMA_VERSION,
    versionManifest,
    periodManifest,
    scopeCacheRows,
    weeklyRpgCacheRows,
    activeReportingGroups: week.activeReportingGroups,
    reportingScopeRestaurantIds: week.reportingScopeRestaurantIds,
    excludedReportingScopeRestaurantIds: week.excludedReportingScopeRestaurantIds,
    validation,
    identityPreflight: week.identityPreflight,
    identityEvidence,
  };
}

function requireActiveCacheSnapshot(cache, activeVersion) {
  if (!cache?.versionManifest || cache.versionManifest.cacheVersion !== activeVersion) {
    fail("PUL-030WIP-016", "Active cache snapshot does not match the resolved authority.");
  }
  if (cache.versionManifest.cacheStatus !== "Active" ||
      cache.versionManifest.activationState !== "Active" ||
      cache.versionManifest.validationStatus !== "PASS") {
    fail("PUL-030WIP-017", "Active cache snapshot is not Active / Active / PASS.");
  }
  const validation = validateCompleteCacheRows(cache);
  if (validation.status !== "PASS") {
    fail("PUL-030WIP-018", `Active cache snapshot failed reconciliation: ${validation.errors.join("; ")}`);
  }
  const fingerprint = fingerprintWeeklyCacheRows({
    cacheVersion: activeVersion,
    periodManifest: cache.periodManifest,
    scopeCacheRows: cache.scopeCacheRows,
    weeklyRpgCacheRows: cache.weeklyRpgCacheRows,
  });
  if (fingerprint !== cache.versionManifest.cacheFingerprint) {
    fail("PUL-030WIP-019", `Active CacheFingerprint ${fingerprint} differs from manifest ${cache.versionManifest.cacheFingerprint}.`);
  }
  return cache;
}

function validateCompleteCacheRows({
  periodManifest,
  scopeCacheRows,
  weeklyRpgCacheRows,
  activeReportingGroups,
}) {
  const errors = [];
  const periods = Array.isArray(periodManifest) ? periodManifest : [];
  const scopes = Array.isArray(scopeCacheRows) ? scopeCacheRows : [];
  const rpgs = Array.isArray(weeklyRpgCacheRows) ? weeklyRpgCacheRows : [];
  const groups = Array.isArray(activeReportingGroups) ? activeReportingGroups : [];
  const periodKeys = new Set();
  for (const row of periods) {
    if (periodKeys.has(row.sourcePeriodKey)) errors.push(`Duplicate period ${row.sourcePeriodKey}.`);
    periodKeys.add(row.sourcePeriodKey);
  }
  const scopeGrains = new Set();
  for (const row of scopes) {
    const grain = `${row.cacheVersion}|${row.sourcePeriodKey}|${row.restaurantId}`;
    if (scopeGrains.has(grain)) errors.push(`Duplicate scope grain ${grain}.`);
    scopeGrains.add(grain);
    const source = sourceMetric(row);
    const coverage = sumMetrics(WEEKLY_CACHE_STATES.map(state => stateMetric(row, CACHE_STATE_FIELDS[state])));
    if (!metricsEqual(source, coverage)) errors.push(`${row.sourcePeriodKey}/${row.restaurantId} coverage does not reconcile.`);
  }
  const rpgGrains = new Set();
  for (const row of rpgs) {
    const grain = `${row.cacheVersion}|${row.sourcePeriodKey}|${row.restaurantId}|${row.reportingGroupId}`;
    if (rpgGrains.has(grain)) errors.push(`Duplicate RPG grain ${grain}.`);
    rpgGrains.add(grain);
  }
  if (groups.length !== 9) errors.push(`Active Reporting Groups are ${groups.length}; expected 9.`);
  if (rpgs.length !== scopes.length * groups.length) {
    errors.push(`Dense RPG rows are ${rpgs.length}; expected ${scopes.length * groups.length}.`);
  }
  for (const period of periods) {
    const expected = {
      factCount: period.sourceFactCount,
      salesNok: period.sourceSalesNok,
      quantity: period.sourceQuantity,
    };
    const actual = sumMetrics(scopes.filter(row => row.sourcePeriodKey === period.sourcePeriodKey)
      .map(sourceMetric));
    if (!metricsEqual(expected, actual)) errors.push(`${period.sourcePeriodKey} source reconciliation failed.`);
  }
  const mappedScope = sumMetrics(scopes.map(row => stateMetric(row, "mapped")));
  const mappedRpg = sumMetrics(rpgs.map(row => ({
    factCount: row.mappedFactCount,
    salesNok: row.mappedSalesNok,
    quantity: row.mappedQuantity,
  })));
  if (!metricsEqual(mappedScope, mappedRpg)) errors.push("Mapped scope does not reconcile to dense RPG rows.");
  return {
    status: errors.length ? "FAIL" : "PASS",
    errors,
    sourceTotals: sumMetrics(scopes.map(sourceMetric)),
    mappingStateCoverage: Object.fromEntries(WEEKLY_CACHE_STATES.map(state => [
      state,
      sumMetrics(scopes.map(row => stateMetric(row, CACHE_STATE_FIELDS[state]))),
    ])),
    mappedRpgTotals: mappedRpg,
  };
}

function validateParsedReport(report) {
  const errors = [];
  if (!report?.manifest || !Array.isArray(report.rows)) return ["Parsed weekly report is missing manifest or rows."];
  const manifest = report.manifest;
  if (report.parserVersion !== WEEKLY_SALES_PARSER_VERSION ||
      manifest.parserVersion !== WEEKLY_SALES_PARSER_VERSION) {
    errors.push(`ParserVersion must be ${WEEKLY_SALES_PARSER_VERSION}.`);
  }
  if (report.schemaVersion !== WEEKLY_SALES_SCHEMA_VERSION ||
      manifest.schemaVersion !== WEEKLY_SALES_SCHEMA_VERSION) {
    errors.push(`SchemaVersion must be ${WEEKLY_SALES_SCHEMA_VERSION}.`);
  }
  if (manifest.contentReconciliationStatus !== "PASS") errors.push("Source report reconciliation is not PASS.");
  if (Number(manifest.inclusiveDays) !== 7) errors.push("Source period is not seven days.");
  if (report.rows.length !== Number(manifest.sourceRowCount)) errors.push("Source row count differs from parsed rows.");
  const totals = sumMetrics(report.rows.map(row => ({
    factCount: 1,
    salesNok: row.salesNok,
    quantity: row.quantity,
  })));
  if (!metricsEqual(totals, {
    factCount: manifest.sourceRowCount,
    salesNok: manifest.totalSalesNok,
    quantity: manifest.totalQuantity,
  })) errors.push("Parsed row totals do not reconcile to the source manifest.");
  for (const row of report.rows) {
    if (row.sourcePeriodKey !== manifest.sourcePeriodKey || row.sourceFileId !== manifest.sourceFileId) {
      errors.push("Parsed row lineage differs from the source manifest.");
      break;
    }
  }
  return errors;
}

function noChangeResult(outcome, message, values) {
  if (!WEEKLY_INTAKE_OUTCOMES.includes(outcome)) fail("PUL-030WIP-020", `Unsupported outcome ${outcome}.`);
  const ledgerEntry = buildIntakeLogEntry({
    outcome,
    message,
    parsedReport: values.parsedReport,
    sourceLocator: values.sourceLocator,
    processedAt: values.processedAt,
    priorCacheVersion: values.priorCacheVersion ?? values.freshness?.activeVersion ?? "",
    identityPreflightFingerprint: values.identityPreflightFingerprint ?? "",
  });
  return {
    outcome,
    cacheChanged: false,
    candidatePrepared: false,
    activeCacheVersion: values.freshness?.activeVersion ?? values.priorCacheVersion ?? "",
    candidate: null,
    freshness: values.freshness ?? null,
    identityPendingAccepted: outcome !== "Rejected",
    ledgerEntry,
  };
}

function buildIntakeLogEntry({
  outcome,
  message,
  parsedReport,
  sourceLocator = "",
  processedAt = "",
  priorCacheVersion = "",
  resultingCacheVersion = "",
  supersededCacheVersion = "",
  identityPreflightFingerprint = "",
}) {
  const manifest = parsedReport?.manifest ?? {};
  const locator = String(manifest.sourceLocator ?? sourceLocator ?? "");
  const eventKey = manifest.semanticFingerprint
    ? [manifest.sourcePeriodKey, manifest.semanticFingerprint, outcome, resultingCacheVersion]
    : [locator, outcome, message];
  return {
    intakeEventId: hashStrings([record("INTAKE", eventKey)], "WINT-"),
    sourceLocator: locator,
    sourceFileId: String(manifest.sourceFileId ?? ""),
    sourcePeriodKey: String(manifest.sourcePeriodKey ?? ""),
    sourceSemanticFingerprint: String(manifest.semanticFingerprint ?? ""),
    identityPreflightFingerprint: String(identityPreflightFingerprint),
    intakeStatus: outcome,
    statusMessage: String(message),
    sourceRowCount: Number(manifest.sourceRowCount ?? 0),
    sourceSalesNok: round(Number(manifest.totalSalesNok ?? 0), 2),
    processedAt: String(processedAt),
    priorCacheVersion: String(priorCacheVersion),
    resultingCacheVersion: String(resultingCacheVersion),
    supersededCacheVersion: String(supersededCacheVersion),
  };
}

function normalizedIdentityEvidence(active) {
  const evidence = Array.isArray(active.identityEvidence) && active.identityEvidence.length
    ? active.identityEvidence
    : [{
      evidenceKey: `BASE:${active.versionManifest.cacheVersion}`,
      identityPreflightFingerprint: active.versionManifest.identityPreflightFingerprint,
    }];
  const seen = new Set();
  const result = [];
  for (const row of evidence) {
    const key = requiredText(row.evidenceKey, "identity evidence key");
    const fingerprint = requiredText(row.identityPreflightFingerprint, "identity evidence fingerprint");
    if (seen.has(key)) fail("PUL-030WIP-021", `Duplicate identity evidence ${key}.`);
    seen.add(key);
    result.push({ evidenceKey: key, identityPreflightFingerprint: fingerprint });
  }
  return result;
}

function fingerprintCorpusManifest(rows) {
  const records = rows.slice().sort(comparePeriodRows).map(row => record("FILE", [
    row.sourceFileId,
    row.sourceSemanticFingerprint,
    row.sourcePeriodKey,
    row.sourceFactCount,
    row.sourceSalesNok,
    row.sourceQuantity,
    row.scopeFingerprint,
  ]));
  return hashStrings(records, "WSC-");
}

function rekeyPeriod(row, cacheVersion) {
  return {
    ...row,
    weeklyPeriodManifestRowId: stableId("WPER-", [cacheVersion, row.sourcePeriodKey]),
    cacheVersion,
  };
}

function rekeyScope(row, cacheVersion) {
  return {
    ...row,
    weeklyScopeCacheRowId: stableId("WSCP-", [cacheVersion, row.sourcePeriodKey, row.restaurantId]),
    cacheVersion,
  };
}

function rekeyRpg(row, cacheVersion) {
  return {
    ...row,
    weeklyRpgCacheRowId: stableId("WRPG-", [
      cacheVersion,
      row.sourcePeriodKey,
      row.restaurantId,
      row.reportingGroupId,
    ]),
    cacheVersion,
  };
}

function sourceMetric(row) {
  return { factCount: row.sourceFactCount, salesNok: row.sourceSalesNok, quantity: row.sourceQuantity };
}

function stateMetric(row, prefix) {
  return {
    factCount: row[`${prefix}FactCount`],
    salesNok: row[`${prefix}SalesNok`],
    quantity: row[`${prefix}Quantity`],
  };
}

function sumMetrics(metrics) {
  const result = { factCount: 0, salesNok: 0, quantity: 0 };
  for (const metric of metrics) {
    result.factCount += Number(metric?.factCount ?? 0);
    result.salesNok += Number(metric?.salesNok ?? 0);
    result.quantity += Number(metric?.quantity ?? 0);
  }
  return {
    factCount: result.factCount,
    salesNok: round(result.salesNok, 2),
    quantity: round(result.quantity, 6),
  };
}

function metricsEqual(left, right) {
  return Number(left?.factCount) === Number(right?.factCount) &&
    Math.abs(Number(left?.salesNok) - Number(right?.salesNok)) <= 0.005 &&
    Math.abs(Number(left?.quantity) - Number(right?.quantity)) <= 0.0000005;
}

function isNonzeroRpg(row) {
  return Number(row.mappedFactCount) !== 0 || Number(row.mappedSalesNok) !== 0 ||
    Number(row.mappedQuantity) !== 0;
}

function comparePeriodRows(left, right) {
  return compareText(left.periodStart, right.periodStart) ||
    compareText(left.sourcePeriodKey, right.sourcePeriodKey);
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

function stableId(prefix, values) {
  return hashStrings([record("ID", values)], prefix);
}

function round(value, decimals) {
  const factor = 10 ** decimals;
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

function requiredText(value, label) {
  const result = String(value ?? "").trim();
  if (!result) fail("PUL-030WIP-022", `${label} is required.`);
  return result;
}

function compareText(left, right) {
  const a = String(left); const b = String(right);
  return a < b ? -1 : a > b ? 1 : 0;
}

function record(kind, values) {
  return `${kind}|${values.map(value => {
    const normalized = value === null || value === undefined ? "" : String(value);
    return `${normalized.length}:${normalized}`;
  }).join("|")}`;
}

function hashStrings(values, prefix) {
  let left = 0; let right = 0;
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
