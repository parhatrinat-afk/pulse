import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import {
  WEEKLY_FULL_VERSION_RETENTION,
  WEEKLY_INTAKE_LOG_COLUMNS,
  WEEKLY_INTAKE_OUTCOMES,
  WEEKLY_INTAKE_PUBLISHER_VERSION,
  planWeeklyFullVersionRetention,
  planWeeklyIntakePublication,
} from "../src/imports/weekly-intake-publisher.mjs";
import { buildCandidateWeeklyCache } from "../src/imports/weekly-compact-cache.mjs";
import { buildWeeklyIdentityPreflight } from "../src/imports/weekly-identity-preflight.mjs";
import {
  WEEKLY_SALES_HEADERS,
  parseWeeklySalesMatrix,
} from "../src/imports/weekly-sales-parser.mjs";

const moduleUrl = new URL("../src/imports/weekly-intake-publisher.mjs", import.meta.url);
const expectedUrl = new URL("./expected-build-0.3.0-weekly-compact-cache.json", import.meta.url);

test("dynamic intake contract prepares a complete deterministic Candidate for one new week", () => {
  const catalog = fixtureCatalog();
  const history = [
    report(2026, 1, [sourceRow("Known Product", 1, 100)]),
    report(2026, 2, [sourceRow("Known Product", 2, 200)]),
  ];
  const incoming = report(2026, 3, [sourceRow("Known Product", 3, 300)]);
  const active = buildActiveCache(history, catalog);
  const activeBefore = structuredClone(active);
  const result = publish({
    active,
    catalog,
    parsedReport: incoming,
  });
  const fullRebuild = buildInactiveCache([...history, incoming], catalog);

  assert.equal(WEEKLY_INTAKE_PUBLISHER_VERSION, "0.3.0-weekly-intake-publisher-v1");
  assert.equal(result.outcome, "New");
  assert.equal(result.cacheChanged, false);
  assert.equal(result.candidatePrepared, true);
  assert.equal(result.candidate.versionManifest.cacheStatus, "Candidate");
  assert.equal(result.candidate.versionManifest.activationState, "Not Active");
  assert.equal(result.candidate.versionManifest.validationStatus, "PASS");
  assert.equal(result.candidate.periodManifest.length, 3);
  assert.equal(result.candidate.scopeCacheRows.length, 3);
  assert.equal(result.candidate.weeklyRpgCacheRows.length, 27);
  assert.deepEqual(result.candidate.validation.sourceTotals, {
    factCount: 3, salesNok: 600, quantity: 6,
  });
  assert.notEqual(result.candidate.versionManifest.cacheVersion, active.versionManifest.cacheVersion);
  assert.match(result.candidate.versionManifest.cacheVersion, /^WCV-[0-9a-f]{16}$/);
  assert.match(result.candidate.versionManifest.cacheFingerprint, /^WCC-[0-9a-f]{16}$/);
  assert.ok(result.candidate.periodManifest.every(row =>
    row.cacheVersion === result.candidate.versionManifest.cacheVersion));
  assert.ok(result.candidate.scopeCacheRows.every(row =>
    row.cacheVersion === result.candidate.versionManifest.cacheVersion));
  assert.ok(result.candidate.weeklyRpgCacheRows.every(row =>
    row.cacheVersion === result.candidate.versionManifest.cacheVersion));
  assert.deepEqual(withoutVersionedIds(result.candidate.scopeCacheRows, "weeklyScopeCacheRowId"),
    withoutVersionedIds(fullRebuild.scopeCacheRows, "weeklyScopeCacheRowId"));
  assert.deepEqual(withoutVersionedIds(result.candidate.weeklyRpgCacheRows, "weeklyRpgCacheRowId"),
    withoutVersionedIds(fullRebuild.weeklyRpgCacheRows, "weeklyRpgCacheRowId"));
  assert.deepEqual(active, activeBefore, "publisher must not mutate the Active cache");
});

test("same period/content is Duplicate and changed content is Correction Review", () => {
  const catalog = fixtureCatalog();
  const original = report(2026, 1, [sourceRow("Known Product", 1, 100)], "original.xlsx");
  const active = buildActiveCache([original], catalog);
  const duplicate = publish({
    active,
    catalog,
    parsedReport: report(2026, 1, [sourceRow("Known Product", 1, 100)], "banana.xlsx"),
  });
  const correction = publish({
    active,
    catalog,
    parsedReport: report(2026, 1, [sourceRow("Known Product", 1, 101)], "correction.xlsx"),
  });

  assert.equal(duplicate.outcome, "Duplicate");
  assert.equal(duplicate.candidate, null);
  assert.equal(duplicate.cacheChanged, false);
  assert.match(duplicate.ledgerEntry.statusMessage, /no cache change/i);
  assert.equal(correction.outcome, "Correction Review");
  assert.equal(correction.candidate, null);
  assert.equal(correction.cacheChanged, false);
  assert.match(correction.ledgerEntry.statusMessage, /explicit supersession/i);
});

test("parser/reconciliation failure is Rejected without a cache candidate", () => {
  const result = planWeeklyIntakePublication({
    parseError: "PUL-030I-008: Header 2 is invalid.",
    sourceLocator: "MANUAL-QA-invalid",
    versionManifests: [],
    processedAt: "2026-08-16T12:00:00Z",
  });
  assert.equal(result.outcome, "Rejected");
  assert.equal(result.candidate, null);
  assert.equal(result.cacheChanged, false);
  assert.match(result.ledgerEntry.statusMessage, /Header 2/);
});

test("mapping, catalog, identity, scope, zero-active and multiple-active states block as Cache Stale", () => {
  const catalog = fixtureCatalog();
  const active = buildActiveCache([report(2026, 1, [sourceRow("Known Product", 1, 100)])], catalog);
  const incoming = report(2026, 2, [sourceRow("Known Product", 1, 100)]);
  for (const [field, value] of [
    ["mappingContentFingerprint", "MCF-changed"],
    ["catalogContentFingerprint", "ICC-changed"],
    ["identityPreflightFingerprint", "IDP-changed"],
    ["performanceRestaurantScopeFingerprint", "RSC-changed"],
  ]) {
    const result = planWeeklyIntakePublication({
      parsedReport: incoming,
      activeCache: active,
      versionManifests: [active.versionManifest],
      catalogs: catalog,
      currentFreshness: { ...freshness(active), [field]: value },
    });
    assert.equal(result.outcome, "Cache Stale", field);
    assert.equal(result.candidate, null);
  }
  assert.equal(planWeeklyIntakePublication({
    parsedReport: incoming, activeCache: active, versionManifests: [], catalogs: catalog,
    currentFreshness: freshness(active),
  }).outcome, "Cache Stale");
  assert.equal(planWeeklyIntakePublication({
    parsedReport: incoming,
    activeCache: active,
    versionManifests: [
      active.versionManifest,
      { ...active.versionManifest, cacheVersion: "WCV-other" },
    ],
    catalogs: catalog,
    currentFreshness: freshness(active),
  }).outcome, "Cache Stale");
});

test("Identity Pending alone is accepted, reconciled and excluded from RPG numerators", () => {
  const catalog = fixtureCatalog();
  const active = buildActiveCache([report(2026, 1, [sourceRow("Known Product", 1, 100)])], catalog);
  const pendingRows = [
    sourceRow("New Multi", 1, 40, "Main A", "Sub A"),
    sourceRow("New Multi", 2, 60, "Main B", "Sub B"),
  ];
  const result = publish({ active, catalog, parsedReport: report(2026, 2, pendingRows) });
  const newScope = result.candidate.scopeCacheRows.find(row =>
    row.sourcePeriodKey === result.candidate.periodManifest.at(-1).sourcePeriodKey);
  const newRpgRows = result.candidate.weeklyRpgCacheRows.filter(row =>
    row.sourcePeriodKey === newScope.sourcePeriodKey);

  assert.equal(result.outcome, "New");
  assert.equal(result.identityPendingAccepted, true);
  assert.equal(newScope.sourceFactCount, 2);
  assert.equal(newScope.sourceSalesNok, 100);
  assert.equal(newScope.identityPendingFactCount, 2);
  assert.equal(newScope.identityPendingSalesNok, 100);
  assert.equal(newScope.mappedFactCount, 0);
  assert.ok(newRpgRows.every(row => row.mappedFactCount === 0 && row.mappedSalesNok === 0));
  assert.equal(result.candidate.validation.status, "PASS");
});

test("candidate version/fingerprint ignore filename and source row order", () => {
  const catalog = fixtureCatalog();
  const active = buildActiveCache([report(2026, 1, [sourceRow("Known Product", 1, 100)])], catalog);
  const rows = [
    sourceRow("Known Product", 1, 100),
    sourceRow("Unmapped Product", 2, 50, "No Map", "No Map"),
  ];
  const first = publish({
    active, catalog, parsedReport: report(2026, 2, rows, "weekly.xlsx"),
  });
  const second = publish({
    active, catalog, parsedReport: report(2026, 2, [...rows].reverse(), "banana.xlsx"),
  });
  assert.equal(first.candidate.versionManifest.cacheVersion, second.candidate.versionManifest.cacheVersion);
  assert.equal(first.candidate.versionManifest.cacheFingerprint, second.candidate.versionManifest.cacheFingerprint);
  assert.deepEqual(first.candidate.scopeCacheRows, second.candidate.scopeCacheRows);
  assert.deepEqual(first.candidate.weeklyRpgCacheRows, second.candidate.weeklyRpgCacheRows);
});

test("incremental publication carries accepted exact identities across weeks", () => {
  const catalog = fixtureCatalog();
  const history = [report(2026, 1, [
    sourceRow("Registry Product", 1, 100, "Main A", "Sub A"),
  ])];
  const incoming = report(2026, 2, [
    sourceRow("Registry Product", 2, 200, "Main A", "Sub A"),
  ]);
  const active = buildActiveCache(history, catalog);
  const acceptedProduct = active.identityPreflight.newIdentityCandidates.products[0];
  const result = publish({ active, catalog, parsedReport: incoming });
  const full = buildInactiveCache([...history, incoming], catalog);

  assert.equal(result.outcome, "New");
  assert.equal(result.candidate.identityPreflight.newIdentityCandidates.products.length, 0);
  assert.equal(result.candidate.identityPreflight.rowAssignments[0].productId,
    acceptedProduct.productId);
  assert.equal(result.candidate.identityRegistry.products.length, 1);
  assert.equal(result.candidate.identityRegistry.products[0].productId,
    acceptedProduct.productId);
  assert.deepEqual(withoutVersionedIds(result.candidate.scopeCacheRows,
    "weeklyScopeCacheRowId"), withoutVersionedIds(full.scopeCacheRows,
    "weeklyScopeCacheRowId"));
  assert.deepEqual(withoutVersionedIds(result.candidate.weeklyRpgCacheRows,
    "weeklyRpgCacheRowId"), withoutVersionedIds(full.weeklyRpgCacheRows,
    "weeklyRpgCacheRowId"));
});

test("intake ledger stays minimal, operational and deterministic", () => {
  assert.deepEqual(WEEKLY_INTAKE_OUTCOMES, [
    "New", "Duplicate", "Correction Review", "Rejected", "Cache Stale",
  ]);
  assert.deepEqual(WEEKLY_INTAKE_LOG_COLUMNS.map(([header]) => header), [
    "IntakeEventID", "SourceLocator", "SourceFileID", "SourcePeriodKey",
    "SourceSemanticFingerprint", "IdentityPreflightFingerprint", "IntakeStatus",
    "StatusMessage", "SourceRowCount", "SourceSalesNOK", "ProcessedAt",
    "PriorCacheVersion", "ResultingCacheVersion", "SupersededCacheVersion",
  ]);
  const catalog = fixtureCatalog();
  const active = buildActiveCache([report(2026, 1, [sourceRow("Known Product", 1, 100)])], catalog);
  const input = report(2026, 2, [sourceRow("Known Product", 1, 100)], "banana.xlsx");
  const first = publish({ active, catalog, parsedReport: input });
  const second = publish({ active, catalog, parsedReport: input });
  assert.equal(first.ledgerEntry.intakeEventId, second.ledgerEntry.intakeEventId);
  assert.equal(first.ledgerEntry.sourceLocator, "banana.xlsx");
  assert.equal(first.ledgerEntry.intakeStatus, "New");
  assert.equal(first.ledgerEntry.priorCacheVersion, active.versionManifest.cacheVersion);
  assert.equal(first.ledgerEntry.resultingCacheVersion, first.candidate.versionManifest.cacheVersion);
});

test("retention contract keeps Candidate plus Active full rows and retains all manifests", () => {
  assert.equal(WEEKLY_FULL_VERSION_RETENTION.maximumMaterializedFullVersions, 2);
  const plan = planWeeklyFullVersionRetention({
    versionManifests: [
      { cacheVersion: "WCV-active", cacheStatus: "Active", activationState: "Active" },
      { cacheVersion: "WCV-old", cacheStatus: "Inactive", activationState: "Not Active" },
    ],
    candidateVersion: "WCV-candidate",
  });
  assert.deepEqual(plan.retainFullVersions, ["WCV-active", "WCV-candidate"]);
  assert.deepEqual(plan.eligibleForAnalyticalRowCleanup, ["WCV-old"]);
  assert.equal(plan.cleanupAction, "None in this slice");
  assert.equal(plan.retainAllManifests, true);
});

test("accepted 84-week checkpoint remains unchanged and publisher contains no workbook/flow mutation", async () => {
  const [expected, source] = await Promise.all([
    readFile(expectedUrl, "utf8").then(JSON.parse),
    readFile(moduleUrl, "utf8"),
  ]);
  assert.equal(expected.cache_version, "WCV-1a34ad1f46763d9b");
  assert.equal(expected.cache_fingerprint, "WCC-508dd608166cdb6e");
  assert.equal(expected.counts.periodRows, 84);
  assert.equal(expected.counts.scopeCacheRows, 1421);
  assert.equal(expected.counts.denseRpgCacheRows, 12789);
  assert.equal(expected.status, "PASS");
  assert.doesNotMatch(source, /ExcelScript|Pulse_Current\.xlsx|moveFile|rename\(|writeFile|unlink|rm\(/);
});

function publish({ active, catalog, parsedReport }) {
  return planWeeklyIntakePublication({
    parsedReport,
    activeCache: active,
    versionManifests: [active.versionManifest],
    catalogs: catalog,
    currentFreshness: freshness(active),
    processedAt: "2026-08-16T12:00:00Z",
  });
}

function buildActiveCache(parsedReports, catalogs) {
  const candidate = buildInactiveCache(parsedReports, catalogs);
  candidate.versionManifest = {
    ...candidate.versionManifest,
    cacheStatus: "Active",
    activationState: "Active",
  };
  return candidate;
}

function buildInactiveCache(parsedReports, catalogs) {
  const preflight = buildWeeklyIdentityPreflight({ parsedReports, catalogs });
  const candidate = buildCandidateWeeklyCache({
    parsedReports,
    catalogs,
    expectedMappingContentFingerprint: preflight.mappingContentFingerprint,
    expectedIdentityPreflightFingerprint: preflight.fingerprints.preflightFingerprint,
  });
  return candidate;
}

function withoutVersionedIds(rows, idField) {
  return rows.map(row => {
    const { [idField]: ignoredId, cacheVersion: ignoredVersion, ...values } = row;
    return values;
  });
}

function freshness(active) {
  return {
    mappingContentFingerprint: active.versionManifest.mappingContentFingerprint,
    catalogContentFingerprint: active.versionManifest.catalogContentFingerprint,
    identityPreflightFingerprint: active.versionManifest.identityPreflightFingerprint,
    performanceRestaurantScopeFingerprint: active.versionManifest.performanceRestaurantScopeFingerprint,
  };
}

function fixtureCatalog() {
  return {
    sourceSystemId: "SRC-TEST-SALES",
    catalogAsOfDate: "2026-08-12",
    catalogAsOfExcelSerial: 46246,
    mappingFingerprint: "MAP-fixture",
    restaurants: [{
      restaurantId: "RST-0001", sourceSystemId: "SRC-TEST-SALES",
      sourceRestaurantName: "Known Restaurant", displayName: "Known Restaurant",
      status: "Active", reportingEnabled: "Yes",
    }],
    classifications: [
      classification("SCL-00001", "Main A", "Sub A"),
      classification("SCL-00002", "Main B", "Sub B"),
      classification("SCL-00003", "No Map", "No Map"),
    ],
    products: [
      product("PRD-000001", "Known Product", "SCL-00001"),
      product("PRD-000002", "Unmapped Product", "SCL-00003"),
    ],
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

function product(id, name, classificationId) {
  return {
    productId: id,
    sourceSystemId: "SRC-TEST-SALES",
    sourceProductName: name,
    salesAccount: "3000 - Sales 25%",
    sourceClassificationId: classificationId,
    productKey: `SRC-TEST-SALES || ${name} || 3000 - Sales 25%`,
    productStatus: "Active",
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

function sourceRow(item, quantity, amount, main = "Main A", sub = "Sub A") {
  return ["Known Restaurant", main, sub, "3000 - Sales 25%", item, quantity, amount];
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
