import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  WEEKLY_PERFORMANCE_CONTRACT_VERSION,
  buildWeeklyPerformanceComparison,
  validateWeeklyPeriodSelection,
} from "../src/reporting/weekly-performance.mjs";

const scriptUrl = new URL("../office-scripts/Build_0_3_0_Weekly_Performance.ts", import.meta.url);
const phase2CUrl = new URL("../office-scripts/Build_0_3_0_Phase2C.ts", import.meta.url);
const expectedUrl = new URL("./expected-build-0.3.0-weekly-compact-cache.json", import.meta.url);
const script = await readFile(scriptUrl, "utf8");
const phase2C = await readFile(phase2CUrl, "utf8");
const expected = JSON.parse(await readFile(expectedUrl, "utf8"));

test("weekly Performance contract validates complete ISO-week selections", () => {
  const result = validateWeeklyPeriodSelection(manifest(), { isoYear: 2026, weekStart: "W01", weekEnd: "W03" });
  assert.equal(WEEKLY_PERFORMANCE_CONTRACT_VERSION, "0.3.0-weekly-performance-v1");
  assert.deepEqual(result, {
    status: "Valid", available: true, summary: "2026 W01–W03", isoYear: 2026,
    weekStart: 1, weekEnd: 3, availableWeeks: 3, expectedWeeks: 3,
  });
});

test("incomplete, unavailable, reversed, and duplicate periods block calculation", () => {
  const incomplete = validateWeeklyPeriodSelection(manifest().filter(row => row.isoWeek !== 2), {
    isoYear: 2026, weekStart: 1, weekEnd: 3,
  });
  assert.equal(incomplete.status, "Incomplete");
  assert.equal(incomplete.summary, "2026 W01–W03 — incomplete (2/3 weeks)");
  assert.equal(validateWeeklyPeriodSelection(manifest(), {
    isoYear: 2024, weekStart: 1, weekEnd: 2,
  }).status, "Invalid");
  assert.equal(validateWeeklyPeriodSelection(manifest(), {
    isoYear: 2026, weekStart: 3, weekEnd: 1,
  }).status, "Invalid");
  assert.equal(validateWeeklyPeriodSelection([...manifest(), manifest().find(row => row.isoYear === 2026)], {
    isoYear: 2026, weekStart: 1, weekEnd: 3,
  }).status, "Invalid");
});

test("active fresh cache supplies independent additive range components", () => {
  const result = buildWeeklyPerformanceComparison({
    cache: cache(), versionManifests: [activeVersion()], currentFreshness: freshness(),
    currentSelection: { isoYear: 2026, weekStart: 1, weekEnd: 2 },
    comparisonSelection: { isoYear: 2025, weekStart: 1, weekEnd: 2 },
    reportingGroupIds: ["RPG-0001"], restaurantIds: ["RST-0001"],
  });
  assert.equal(result.status, "Available");
  assert.equal(result.warning, "");
  assert.equal(result.result.currentDenominator.salesNok, 4000);
  assert.equal(result.result.comparisonDenominator.salesNok, 3000);
  assert.equal(result.result.results[0].currentSalesNok, 400);
  assert.equal(result.result.results[0].currentShare, 0.1);
  assert.equal(result.result.results[0].comparisonShare, 0.05);
  assert.equal(result.result.results[0].ppChange, 5);
  assert.equal(result.result.results[0].nokImpact, 200);
});

test("same period is allowed and produces zero PP/NOK Impact", () => {
  const result = buildWeeklyPerformanceComparison({
    cache: cache(), versionManifests: [activeVersion()], currentFreshness: freshness(),
    currentSelection: { isoYear: 2026, weekStart: 1, weekEnd: 2 },
    comparisonSelection: { isoYear: 2026, weekStart: 1, weekEnd: 2 },
    reportingGroupIds: ["RPG-0001"], restaurantIds: ["RST-0001"],
  });
  assert.equal(result.warning, "Same Current and Compare period.");
  assert.equal(result.result.results[0].ppChange, 0);
  assert.equal(result.result.results[0].nokImpact, 0);
});

test("different complete lengths remain available with a warning", () => {
  const result = buildWeeklyPerformanceComparison({
    cache: cache(), versionManifests: [activeVersion()], currentFreshness: freshness(),
    currentSelection: { isoYear: 2026, weekStart: 1, weekEnd: 3 },
    comparisonSelection: { isoYear: 2025, weekStart: 1, weekEnd: 2 },
    reportingGroupIds: ["RPG-0001"], restaurantIds: ["RST-0001"],
  });
  assert.equal(result.status, "Available");
  assert.equal(result.warning, "Different complete period lengths (3 vs 2 weeks) — comparison allowed.");
});

test("stale cache and incomplete ranges return unavailable without calculation", () => {
  const stale = buildWeeklyPerformanceComparison({
    cache: cache(), versionManifests: [activeVersion()], currentFreshness: freshness({ mappingContentFingerprint: "MCF-changed" }),
    currentSelection: { isoYear: 2026, weekStart: 1, weekEnd: 2 },
    comparisonSelection: { isoYear: 2025, weekStart: 1, weekEnd: 2 },
    reportingGroupIds: ["RPG-0001"], restaurantIds: ["RST-0001"],
  });
  assert.equal(stale.status, "Stale");
  assert.equal(stale.result, undefined);
  const incompleteCache = cache();
  incompleteCache.periodManifest = incompleteCache.periodManifest.filter(row => !(row.isoYear === 2026 && row.isoWeek === 2));
  const incomplete = buildWeeklyPerformanceComparison({
    cache: incompleteCache, versionManifests: [activeVersion()], currentFreshness: freshness(),
    currentSelection: { isoYear: 2026, weekStart: 1, weekEnd: 2 },
    comparisonSelection: { isoYear: 2025, weekStart: 1, weekEnd: 2 },
    reportingGroupIds: ["RPG-0001"], restaurantIds: ["RST-0001"],
  });
  assert.equal(incomplete.status, "Unavailable");
  assert.equal(incomplete.current.status, "Incomplete");
});

test("accepted Add-ons weekly range fixtures remain the cutover targets", () => {
  const fixtures = expected.comparisons.map(value => ({
    label: value.label,
    addOns: value.results.find(row => row.reporting_group_id === "RPG-0001"),
  }));
  assert.deepEqual(fixtures.map(value => value.label), [
    "W31 2026 vs W31 2025",
    "W01-W32 2026 vs W01-W32 2025",
    "W20-W30 2026 vs W20-W30 2025",
  ]);
  assert.deepEqual(fixtures.map(value => value.addOns.current_sales_nok), [80263, 2383679, 835122]);
  assert.ok(Math.abs(fixtures[0].addOns.nok_impact - 35001.58321842167) < 1e-9);
  assert.ok(Math.abs(fixtures[1].addOns.nok_impact - 1046091.4622320954) < 1e-9);
  assert.ok(Math.abs(fixtures[2].addOns.nok_impact - 378929.35570764536) < 1e-9);
});

test("Office Script installs the exact weekly controls and accepted defaults", () => {
  assert.match(script, /B10:B12[\s\S]*\[\[2026\], \["W01"\], \["W32"\]\]/);
  assert.match(script, /G10:G12[\s\S]*\[\[2025\], \["W01"\], \["W32"\]\]/);
  assert.match(script, /Performance!\$B\$10/);
  assert.match(script, /Performance!\$G\$10/);
  assert.match(script, /tblWeeklyPeriodManifest/);
  assert.match(script, /incomplete \("&\$?\{?count\}?&"\/"&\$?\{?expected\}?&" weeks\)/);
  assert.match(script, /Different complete period lengths/);
  assert.match(script, /Same Current and Compare period/);
});

test("Office Script resolves one dynamic Active authority without frozen WCV/WCC dependencies", () => {
  assert.doesNotMatch(script, /EXPECTED_CACHE_VERSION|EXPECTED_CACHE_FINGERPRINT/);
  assert.doesNotMatch(script, /WCV-1a34ad1f46763d9b|WCC-508dd608166cdb6e/);
  assert.match(script, /COUNTIFS\(tblWeeklyCacheVersions\[CacheStatus\],"Active",tblWeeklyCacheVersions\[ActivationState\],"Active"\)/);
  assert.match(script, /Unavailable — no active cache/);
  assert.match(script, /Unavailable — multiple active caches/);
  assert.match(script, /XLOOKUP\(v,tblWeeklyCacheVersions\[CacheVersion\],tblWeeklyCacheVersions\[MappingContentFingerprint\]/);
  assert.match(script, /XLOOKUP\(v,tblWeeklyCacheVersions\[CacheVersion\],tblWeeklyCacheVersions\[CatalogContentFingerprint\]/);
  assert.match(script, /XLOOKUP\(v,tblWeeklyCacheVersions\[CacheVersion\],tblWeeklyCacheVersions\[IdentityPreflightFingerprint\]/);
  assert.match(script, /PerformanceRestaurantScopeFingerprint/);
  assert.match(script, /PeriodRowCount/);
  assert.match(script, /ScopeCacheRowCount/);
  assert.match(script, /DenseRPGCacheRowCount/);
});

test("weekly cache replaces only Phase 2C additive component inputs", () => {
  assert.match(script, /writeWeeklyComponentBlock\(calc, "AN", true, true\)/);
  assert.match(script, /writeWeeklyComponentBlock\(calc, "AX", true, false\)/);
  assert.match(script, /writeWeeklyComponentBlock\(calc, "BR", false, true\)/);
  assert.match(script, /writeWeeklyComponentBlock\(calc, "CB", false, false\)/);
  assert.match(script, /numerator \? "tblWeeklyRPGCache" : "tblWeeklyScopeCache"/);
  assert.match(script, /numerator \? "MappedSalesNOK" : "SourceSalesNOK"/);
  assert.match(script, /SUM\(SUMIFS/);
  assert.doesNotMatch(script, /AVERAGEIFS|AVERAGE\(/);
  assert.doesNotMatch(script, /tblSalesFacts\[ReportingCategoryID\]/);
  assert.match(script, /"tblMetricRPGResults", rangeFingerprint\(requiredTable\(workbook, "tblMetricRPGResults"\)\.getRange\(\)\)/);
  assert.match(script, /"tblImports", rangeFingerprint\(requiredTable\(workbook, "tblImports"\)\.getRange\(\)\)/);
});

test("Phase 2C numeric, Total, sorting, facade and selection formulas are not rewritten", () => {
  assert.doesNotMatch(script, /writeNumericDisplayBlock|writeTotalAndSortHelpers|matrixFacadeFormula|SORTBY\(/);
  assert.doesNotMatch(script, /tblPerformanceRestaurantSelection\[[^\]]+\].*set|tblPerformanceRPGSelection\[[^\]]+\].*set/);
  assert.match(phase2C, /NOK Impact/);
  assert.match(phase2C, /Grand Total/);
  assert.match(phase2C, /FIXED\(/);
  assert.match(phase2C, /SORTBY\(/);
});

test("Reports follows generated period summaries while legacy results remain rollback-only", () => {
  assert.match(script, /reports\.getRange\("B8"\)\.setFormula\("=Performance!B13"\)/);
  assert.match(script, /reports\.getRange\("B9"\)\.setFormula\("=Performance!G13"\)/);
  assert.match(script, /"tblMetricRPGResults", rangeFingerprint\(requiredTable\(workbook, "tblMetricRPGResults"\)\.getRange\(\)\)/);
  assert.match(script, /"tblImports", rangeFingerprint\(requiredTable\(workbook, "tblImports"\)\.getRange\(\)\)/);
  assert.doesNotMatch(script, /delete\(\).*tblMetricRPGResults|tblMetricRPGResults.*delete/);
  assert.doesNotMatch(script, /supersed|Period selector.*Month|Quarter|Power Automate|Phase 3/i);
});

test("weekly Performance Office Script remains Office Scripts compatible", () => {
  assert.match(script, /^function main\(/m);
  assert.doesNotMatch(script, /^(export|async|public|private|protected)\s+function\s+main/m);
  assert.doesNotMatch(script, /new\s+Map\s*</);
  assert.doesNotMatch(script, /new\s+Set\s*</);
  assert.doesNotMatch(script, /\.(entries|keys|values)\s*\(\)/);
  assert.doesNotMatch(script, /Array\.from\s*\(/);
  assert.doesNotMatch(script, /\.\.\./);
  const loopBodies = Array.from(script.matchAll(/for\s*\([^)]*\)\s*\{([\s\S]*?)\n  \}/g), match => match[1]);
  for (const body of loopBodies) {
    assert.doesNotMatch(body, /workbook\.get|sheet\.get|table\.get|\.getValues\(|\.getTexts\(/);
  }
  assertBalanced(script);
});

function manifest() {
  return [2025, 2026].flatMap(isoYear => [1, 2, 3].map(isoWeek => ({
    cacheVersion: "WCV-test", sourcePeriodKey: `${isoYear}-W${isoWeek}`, isoYear, isoWeek,
  })));
}

function cache() {
  const periodManifest = manifest();
  const scopeCacheRows = periodManifest.map(row => ({
    cacheVersion: "WCV-test", sourcePeriodKey: row.sourcePeriodKey, restaurantId: "RST-0001",
    sourceFactCount: 1, sourceSalesNok: row.isoYear === 2026 ? 2000 : 1500, sourceQuantity: 1,
  }));
  const weeklyRpgCacheRows = periodManifest.map(row => ({
    cacheVersion: "WCV-test", sourcePeriodKey: row.sourcePeriodKey, restaurantId: "RST-0001",
    reportingGroupId: "RPG-0001", mappedFactCount: 1,
    mappedSalesNok: row.isoYear === 2026 ? 200 : 75, mappedQuantity: 1,
  }));
  return {
    periodManifest, scopeCacheRows, weeklyRpgCacheRows,
    reportingScopeRestaurantIds: ["RST-0001"],
    activeReportingGroups: [{ reportingGroupId: "RPG-0001" }],
  };
}

function activeVersion() {
  return {
    cacheVersion: "WCV-test", cacheStatus: "Active", activationState: "Active", validationStatus: "PASS",
    mappingContentFingerprint: "MCF-test", catalogContentFingerprint: "ICC-test",
    identityPreflightFingerprint: "IDP-test", performanceRestaurantScopeFingerprint: "RSC-test",
  };
}

function freshness(overrides = {}) {
  return {
    mappingContentFingerprint: "MCF-test", catalogContentFingerprint: "ICC-test",
    identityPreflightFingerprint: "IDP-test", performanceRestaurantScopeFingerprint: "RSC-test",
    ...overrides,
  };
}

function assertBalanced(source) {
  const pairs = [["{", "}"], ["(", ")"], ["[", "]"]];
  for (const [open, close] of pairs) {
    assert.equal([...source].filter(value => value === open).length,
      [...source].filter(value => value === close).length, `${open}${close} balance`);
  }
}
