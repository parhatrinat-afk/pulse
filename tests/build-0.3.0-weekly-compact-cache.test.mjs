import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  WEEKLY_RPG_CACHE_GRAIN,
  WEEKLY_SCOPE_CACHE_GRAIN,
  buildCandidateWeeklyCache,
  compareCandidateCacheRanges,
  validateWeeklyCacheFreshness,
} from "../src/imports/weekly-compact-cache.mjs";
import {
  buildWeeklyIdentityPreflight,
  deriveWeeklyMappingContentFingerprint,
} from "../src/imports/weekly-identity-preflight.mjs";
import {
  WEEKLY_SALES_HEADERS,
  parseWeeklySalesMatrix,
} from "../src/imports/weekly-sales-parser.mjs";

const expectedPath = new URL("./expected-build-0.3.0-weekly-compact-cache.json", import.meta.url);
const modulePath = new URL("../src/imports/weekly-compact-cache.mjs", import.meta.url);
const auditPath = new URL("../src/imports/audit-weekly-compact-cache.mjs", import.meta.url);

test("candidate cache uses the approved compact grains and deterministic row IDs", () => {
  const cache = buildCache([report(2026, 1, [
    sourceRow("Known Restaurant", "Main A", "Sub A", "Known Product", 2, 200),
    sourceRow("Known Restaurant", "No Map", "No Map", "Unmapped Product", 1, 50),
  ])]);

  assert.deepEqual(WEEKLY_SCOPE_CACHE_GRAIN, ["CacheVersion", "SourcePeriodKey", "RestaurantID"]);
  assert.deepEqual(WEEKLY_RPG_CACHE_GRAIN, [
    "CacheVersion",
    "SourcePeriodKey",
    "RestaurantID",
    "ReportingGroupID",
  ]);
  assert.equal(cache.scopeCacheRows.length, 1);
  assert.equal(cache.weeklyRpgCacheRows.length, 9);
  assert.equal(new Set(cache.scopeCacheRows.map(scopeGrain)).size, 1);
  assert.equal(new Set(cache.weeklyRpgCacheRows.map(rpgGrain)).size, 9);
  assert.match(cache.scopeCacheRows[0].weeklyScopeCacheRowId, /^WSCP-/);
  assert.ok(cache.weeklyRpgCacheRows.every(row => /^WRPG-/.test(row.weeklyRpgCacheRowId)));
  assert.equal(cache.versionManifest.cacheStatus, "Candidate");
  assert.equal(cache.versionManifest.activationState, "Not Active");
  assert.equal(
    cache.versionManifest.performanceRestaurantScopeFingerprint,
    "RSC-6180f43b08db71db",
  );
  assert.equal(cache.validation.status, "PASS");
});

test("scope denominator exists once per restaurant/week and is absent from RPG rows", () => {
  const cache = buildCache([report(2026, 1, [
    sourceRow("Known Restaurant", "Main A", "Sub A", "Known Product", 2, 200),
    sourceRow("Known Restaurant", "No Map", "No Map", "Unmapped Product", 1, 50),
  ])]);
  const scope = cache.scopeCacheRows[0];

  assert.equal(scope.sourceFactCount, 2);
  assert.equal(scope.sourceSalesNok, 250);
  assert.equal(scope.sourceQuantity, 3);
  assert.ok(cache.weeklyRpgCacheRows.every(row => !("sourceSalesNok" in row)));
  assert.ok(cache.weeklyRpgCacheRows.every(row => !("denominatorSalesNok" in row)));
});

test("Identity Pending remains in denominator and never contributes to an RPG numerator", () => {
  const cache = buildCache([report(2026, 1, [
    sourceRow("Known Restaurant", "Main A", "Sub A", "New Multi", 1, 100),
    sourceRow("Known Restaurant", "Main B", "Sub B", "New Multi", 2, 250),
  ])]);
  const scope = cache.scopeCacheRows[0];

  assert.equal(scope.sourceFactCount, 2);
  assert.equal(scope.sourceSalesNok, 350);
  assert.equal(scope.identityPendingFactCount, 2);
  assert.equal(scope.identityPendingSalesNok, 350);
  assert.equal(scope.mappedFactCount, 0);
  assert.ok(cache.weeklyRpgCacheRows.every(row => row.mappedFactCount === 0));
  assert.equal(cache.validation.corpusReconciliation.status, "PASS");
});

test("current Pulse Product hierarchy remains authority over alternate weekly paths", () => {
  const cache = buildCache([report(2026, 1, [
    sourceRow("Known Restaurant", "Main B", "Sub B", "Known Product", 2, 300),
  ])]);
  const groupA = cache.weeklyRpgCacheRows.find(row => row.reportingGroupId === "RPG-0001");
  const groupB = cache.weeklyRpgCacheRows.find(row => row.reportingGroupId === "RPG-0009");

  assert.equal(groupA.mappedSalesNok, 300);
  assert.equal(groupB.mappedSalesNok, 0);
  assert.equal(cache.identityPreflight.hierarchyReview[0].currentSourceClassificationId, "SCL-00001");
  assert.equal(cache.identityPreflight.hierarchyReview[0].alternatePaths[0]
    .alternateOutcome.effectiveReportingGroupId, "RPG-0009");
});

test("Test Department facts reconcile but stay outside normal Performance scope", () => {
  const cache = buildCache([report(2026, 1, [
    sourceRow("Known Restaurant", "Main A", "Sub A", "Known Product", 1, 100),
    sourceRow("Test Department (Not for User)", "Main A", "Sub A", "Known Product", 1, 20),
    sourceRow("Test Department (Not for Users)", "Main A", "Sub A", "Known Product", 1, 30),
  ])]);

  assert.deepEqual(cache.excludedReportingScopeRestaurantIds, ["RST-0002", "RST-0003"]);
  assert.equal(cache.validation.performanceScopeReconciliation.enabled.salesNok, 100);
  assert.equal(cache.validation.performanceScopeReconciliation.excluded.salesNok, 50);
  assert.equal(cache.validation.performanceScopeReconciliation.complete.salesNok, 150);
  assert.equal(cache.validation.performanceScopeReconciliation.status, "PASS");
});

test("range metrics aggregate additive components before calculating shares and impact", () => {
  const cache = buildCache([
    report(2026, 1, [
      sourceRow("Known Restaurant", "Main A", "Sub A", "Known Product", 1, 10),
      sourceRow("Known Restaurant", "No Map", "No Map", "Unmapped Product", 1, 90),
    ]),
    report(2026, 2, [
      sourceRow("Known Restaurant", "Main A", "Sub A", "Known Product", 1, 30),
      sourceRow("Known Restaurant", "No Map", "No Map", "Unmapped Product", 1, 70),
    ]),
  ]);
  const comparison = compareCandidateCacheRanges({
    cache,
    currentRange: { isoYear: 2026, weekStart: 2, weekEnd: 2 },
    comparisonRange: { isoYear: 2026, weekStart: 1, weekEnd: 1 },
    reportingGroupIds: ["RPG-0001"],
  });
  const result = comparison.results[0];

  assert.equal(comparison.currentDenominator.salesNok, 100);
  assert.equal(comparison.comparisonDenominator.salesNok, 100);
  assert.equal(result.currentShare, 0.3);
  assert.equal(result.comparisonShare, 0.1);
  assert.ok(Math.abs(result.ppChange - 20) < 1e-12);
  assert.equal(result.currentSalesNok, 30);
  assert.equal(result.nokImpact, 20);
});

test("cache version, fingerprint, row IDs and outputs ignore locator and source row order", () => {
  const rows = [
    sourceRow("Known Restaurant", "Main A", "Sub A", "Known Product", 2, 200),
    sourceRow("Known Restaurant", "No Map", "No Map", "Unmapped Product", 1, 50),
  ];
  const first = buildCache([report(2026, 1, rows, "normal.xlsx")]);
  const second = buildCache([report(2026, 1, [...rows].reverse(), "banana.xlsx")]);

  assert.equal(first.versionManifest.cacheVersion, second.versionManifest.cacheVersion);
  assert.equal(first.versionManifest.cacheFingerprint, second.versionManifest.cacheFingerprint);
  assert.deepEqual(first.scopeCacheRows, second.scopeCacheRows);
  assert.deepEqual(first.weeklyRpgCacheRows, second.weeklyRpgCacheRows);
});

test("editable Reporting Group labels do not change stable-ID cache membership or numerators", () => {
  const reports = [report(2026, 1, [
    sourceRow("Known Restaurant", "Main A", "Sub A", "Known Product", 2, 200),
  ])];
  const originalCatalog = fixtureCatalog();
  const relabeledCatalog = fixtureCatalog();
  relabeledCatalog.reportingGroups = relabeledCatalog.reportingGroups.map((row, index) => ({
    ...row,
    reportingGroupName: `Editable label ${index + 1}`,
    sortOrder: 100 - index,
  }));
  const original = buildCache(reports, originalCatalog);
  const relabeled = buildCache(reports, relabeledCatalog);

  assert.equal(
    original.versionManifest.activeReportingGroupFingerprint,
    relabeled.versionManifest.activeReportingGroupFingerprint,
  );
  assert.notEqual(original.versionManifest.cacheVersion, relabeled.versionManifest.cacheVersion);
  assert.deepEqual(
    original.scopeCacheRows.map(({ weeklyScopeCacheRowId, cacheVersion, ...row }) => row),
    relabeled.scopeCacheRows.map(({ weeklyScopeCacheRowId, cacheVersion, ...row }) => row),
  );
  assert.deepEqual(
    original.weeklyRpgCacheRows.map(({ weeklyRpgCacheRowId, cacheVersion, ...row }) => row),
    relabeled.weeklyRpgCacheRows.map(({ weeklyRpgCacheRowId, cacheVersion, ...row }) => row),
  );
});

test("stale mapping and stale accepted preflight fail before candidate construction", () => {
  const reports = [report(2026, 1, [
    sourceRow("Known Restaurant", "Main A", "Sub A", "Known Product", 1, 100),
  ])];
  const catalog = fixtureCatalog();
  const preflight = buildWeeklyIdentityPreflight({ parsedReports: reports, catalogs: catalog });

  assert.throws(() => buildCandidateWeeklyCache({
    parsedReports: reports,
    catalogs: catalog,
    expectedMappingContentFingerprint: "MCF-stale",
    expectedIdentityPreflightFingerprint: preflight.fingerprints.preflightFingerprint,
  }), /PUL-030C-012/);
  assert.throws(() => buildCandidateWeeklyCache({
    parsedReports: reports,
    catalogs: catalog,
    expectedMappingContentFingerprint: preflight.mappingContentFingerprint,
    expectedIdentityPreflightFingerprint: "IDP-stale",
  }), /PUL-030C-004/);
});

test("candidate version cannot overwrite an active cache version", () => {
  const reports = [report(2026, 1, [
    sourceRow("Known Restaurant", "Main A", "Sub A", "Known Product", 1, 100),
  ])];
  const first = buildCache(reports);
  const catalog = fixtureCatalog();
  const preflight = buildWeeklyIdentityPreflight({ parsedReports: reports, catalogs: catalog });

  assert.throws(() => buildCandidateWeeklyCache({
    parsedReports: reports,
    catalogs: catalog,
    expectedMappingContentFingerprint: preflight.mappingContentFingerprint,
    expectedIdentityPreflightFingerprint: preflight.fingerprints.preflightFingerprint,
    existingVersionManifests: [{
      ...first.versionManifest,
      cacheStatus: "Active",
      activationState: "Active",
    }],
  }), /PUL-030C-018/);
});

test("freshness validation is deterministic and user-visible", () => {
  const cache = buildCache([report(2026, 1, [
    sourceRow("Known Restaurant", "Main A", "Sub A", "Known Product", 1, 100),
  ])]);

  assert.deepEqual(validateWeeklyCacheFreshness({
    versionManifest: cache.versionManifest,
    currentMappingContentFingerprint: cache.versionManifest.mappingContentFingerprint,
  }), []);
  assert.deepEqual(validateWeeklyCacheFreshness({
    versionManifest: cache.versionManifest,
    currentMappingContentFingerprint: "MCF-changed",
  }), [
    `Candidate cache MappingContentFingerprint ${cache.versionManifest.mappingContentFingerprint} differs from current MCF-changed.`,
  ]);
});

test("date-only audit advancement does not stale or re-version weekly cache content", () => {
  const reports = [report(2026, 1, [
    sourceRow("Known Restaurant", "Main A", "Sub A", "Known Product", 1, 100),
  ])];
  const august11 = fixtureCatalog();
  august11.catalogAsOfDate = "2026-08-11";
  august11.catalogAsOfExcelSerial = 46245;
  august11.mappingFingerprint = "MAP-date-sensitive-11";
  const august12 = fixtureCatalog();
  august12.mappingFingerprint = "MAP-date-sensitive-12";
  const first = buildCache(reports, august11);
  const second = buildCache(reports, august12);

  assert.equal(first.versionManifest.mappingContentFingerprint,
    second.versionManifest.mappingContentFingerprint);
  assert.equal(first.versionManifest.cacheVersion, second.versionManifest.cacheVersion);
  assert.equal(first.versionManifest.cacheFingerprint, second.versionManifest.cacheFingerprint);
  assert.deepEqual(first.scopeCacheRows, second.scopeCacheRows);
  assert.deepEqual(first.weeklyRpgCacheRows, second.weeklyRpgCacheRows);
  assert.notEqual(first.versionManifest.mappingFingerprint, second.versionManifest.mappingFingerprint);
  assert.notEqual(first.versionManifest.mappingAsOfDate, second.versionManifest.mappingAsOfDate);
  assert.deepEqual(validateWeeklyCacheFreshness({
    versionManifest: first.versionManifest,
    currentMappingContentFingerprint: second.versionManifest.mappingContentFingerprint,
  }), []);
});

test("actual mapping content changes are rejected by the weekly stale gate", () => {
  const reports = [report(2026, 1, [
    sourceRow("Known Restaurant", "Main A", "Sub A", "Known Product", 1, 100),
  ])];
  const accepted = fixtureCatalog();
  const expectedMappingContentFingerprint = deriveWeeklyMappingContentFingerprint(accepted);
  const changed = fixtureCatalog();
  changed.mappingRules[0].targetReportingGroupId = "RPG-0002";

  assert.notEqual(
    deriveWeeklyMappingContentFingerprint(changed),
    expectedMappingContentFingerprint,
  );
  assert.throws(() => buildCandidateWeeklyCache({
    parsedReports: reports,
    catalogs: changed,
    expectedMappingContentFingerprint,
    expectedIdentityPreflightFingerprint: "IDP-not-reached",
  }), /PUL-030C-012.*MappingContentFingerprint/);
});

test("frozen 84-week cache evidence covers all periods, states and required ranges", async () => {
  const expected = JSON.parse(await readFile(expectedPath, "utf8"));

  assert.equal(expected.status, "PASS");
  assert.equal(expected.mapping_content_fingerprint, "MCF-759cc92c4304a913");
  assert.equal(expected.identity_preflight_fingerprint, "IDP-062c182f23905ae8");
  assert.equal(expected.cache_version, "WCV-1a34ad1f46763d9b");
  assert.equal(expected.cache_fingerprint, "WCC-508dd608166cdb6e");
  assert.equal(expected.counts.periodRows, 84);
  assert.equal(expected.counts.scopeCacheRows, 1421);
  assert.equal(expected.counts.denseRpgCacheRows, 1421 * 9);
  assert.equal(expected.period_status_counts.PASS, 84);
  assert.equal(
    expected.performance_restaurant_scope_fingerprint,
    "RSC-08df626f217dd94b",
  );
  assert.equal(expected.corpus.status, "PASS");
  assert.deepEqual(expected.identity_pending, {
    fact_count: 120,
    sales_nok: 114876,
    quantity: 951,
  });
  assert.deepEqual(expected.excluded_reporting_scope_restaurant_ids, ["RST-0017", "RST-0018"]);
  assert.deepEqual(expected.comparisons.map(value => value.label), [
    "W31 2026 vs W31 2025",
    "W01-W32 2026 vs W01-W32 2025",
    "W20-W30 2026 vs W20-W30 2025",
  ]);
  const w31Addons = expected.comparisons[0].results[0];
  assert.equal(w31Addons.reporting_group_id, "RPG-0001");
  assert.equal(w31Addons.current_sales_nok, 80263);
  assert.ok(Math.abs(w31Addons.pp_change - 0.5181916760332484) < 1e-12);
  assert.ok(Math.abs(w31Addons.nok_impact - 35001.58321842167) < 1e-9);
});

test("candidate cache implementation is repository-only and does not cut over Performance", async () => {
  const [source, audit] = await Promise.all([
    readFile(modulePath, "utf8"),
    readFile(auditPath, "utf8"),
  ]);

  assert.match(audit, /Provide one exact fixture corpus path/);
  assert.match(audit, /readOnly: true/);
  assert.doesNotMatch(audit, /\/Users\/|process\.env\.(HOME|CODEX_HOME)|homedir\(/);
  assert.doesNotMatch(`${source}\n${audit}`, /writeFile|copyFile|rename\(|unlink\(|rm\(/);
  assert.doesNotMatch(`${source}\n${audit}`, /tblMetricRPGResults|Build_0_3_0_Phase2C|Pulse_Current\.xlsx/);
});

function buildCache(parsedReports, catalogs = fixtureCatalog()) {
  const preflight = buildWeeklyIdentityPreflight({ parsedReports, catalogs });
  return buildCandidateWeeklyCache({
    parsedReports,
    catalogs,
    expectedMappingContentFingerprint: preflight.mappingContentFingerprint,
    expectedIdentityPreflightFingerprint: preflight.fingerprints.preflightFingerprint,
  });
}

function fixtureCatalog() {
  return {
    sourceSystemId: "SRC-TEST-SALES",
    catalogAsOfDate: "2026-08-12",
    catalogAsOfExcelSerial: 46246,
    mappingFingerprint: "MAP-fixture",
    restaurants: [{
      restaurantId: "RST-0001",
      sourceSystemId: "SRC-TEST-SALES",
      sourceRestaurantName: "Known Restaurant",
      displayName: "Known Restaurant",
      status: "Active",
      reportingEnabled: "Yes",
    }],
    classifications: [
      classification("SCL-00001", "Main A", "Sub A"),
      classification("SCL-00002", "Main B", "Sub B"),
    ],
    products: [{
      productId: "PRD-000001",
      sourceSystemId: "SRC-TEST-SALES",
      sourceProductName: "Known Product",
      salesAccount: "3000 - Sales 25%",
      sourceClassificationId: "SCL-00001",
      productKey: "SRC-TEST-SALES || Known Product || 3000 - Sales 25%",
      productStatus: "Active",
    }],
    reportingGroups: Array.from({ length: 9 }, (_, index) => ({
      reportingGroupId: `RPG-${String(index + 1).padStart(4, "0")}`,
      reportingGroupName: `Group ${index + 1}`,
      active: "Yes",
      sortOrder: (index + 1) * 10,
    })),
    mappingRules: [
      rule("MAP-000001", "SRC-TEST-SALES || Main || Main A", "RPG-0001"),
      rule("MAP-000002", "SRC-TEST-SALES || Main || Main B", "RPG-0009"),
    ],
  };
}

function classification(id, main, sub) {
  return {
    sourceClassificationId: id,
    sourceSystemId: "SRC-TEST-SALES",
    sourceMainCategory: main,
    sourceSubCategory: sub,
    sourceClassificationKey: `SRC-TEST-SALES || ${main} || ${sub}`,
    status: "Active",
  };
}

function rule(id, nodeId, reportingGroupId) {
  return {
    mappingRuleId: id,
    sourceSystemId: "SRC-TEST-SALES",
    scopeType: "SourceMainCategory",
    nodeId,
    targetReportingGroupId: reportingGroupId,
    effectiveFrom: 45000,
    effectiveTo: null,
    status: "Active",
    ruleAction: "Map",
  };
}

function report(isoYear, isoWeek, rows, sourceLocator = "fixture.xlsx") {
  const [periodStart, periodEnd] = periodDates(isoYear, isoWeek);
  return parseWeeklySalesMatrix([
    [`Period: ${periodStart} - ${periodEnd}`],
    [...WEEKLY_SALES_HEADERS],
    ...rows,
  ], { sourceLocator });
}

function sourceRow(restaurant, main, sub, item, quantity, amount) {
  return [restaurant, main, sub, "3000 - Sales 25%", item, quantity, amount];
}

function periodDates(isoYear, isoWeek) {
  const january4 = new Date(Date.UTC(isoYear, 0, 4));
  const weekday = january4.getUTCDay() || 7;
  const monday = new Date(january4);
  monday.setUTCDate(january4.getUTCDate() - weekday + 1 + (isoWeek - 1) * 7);
  const sunday = new Date(monday);
  sunday.setUTCDate(monday.getUTCDate() + 6);
  return [monday.toISOString().slice(0, 10), sunday.toISOString().slice(0, 10)];
}

function scopeGrain(row) {
  return [row.cacheVersion, row.sourcePeriodKey, row.restaurantId].join("|");
}

function rpgGrain(row) {
  return [row.cacheVersion, row.sourcePeriodKey, row.restaurantId, row.reportingGroupId].join("|");
}
