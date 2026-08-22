import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { activeReportingGroups } from "../src/imports/weekly-compact-cache.mjs";
import {
  activeWeeklyPerformanceReportingGroups,
  buildWeeklyPerformanceLayout,
  planWeeklyPerformanceRpgSelection,
} from "../src/reporting/weekly-performance.mjs";

const [performanceScript, publisherScript, cacheSource, publisherSource] = await Promise.all([
  readFile(new URL("../office-scripts/Build_0_3_0_Weekly_Performance.ts", import.meta.url), "utf8"),
  readFile(new URL("../office-scripts/Publish_Weekly_Intake.ts", import.meta.url), "utf8"),
  readFile(new URL("../src/imports/weekly-compact-cache.mjs", import.meta.url), "utf8"),
  readFile(new URL("../src/imports/weekly-intake-publisher.mjs", import.meta.url), "utf8"),
]);

test("current nine-group runtime keeps the accepted helper geometry", () => {
  const layout = buildWeeklyPerformanceLayout({
    reportingGroups: groups(9), restaurantCapacity: 16,
  });
  assert.equal(layout.groupCapacity, 9);
  assert.deepEqual(layout.componentBlocks, [39, 49, 59, 69, 79, 89]);
  assert.equal(layout.numericDisplayStartColumn, 99);
  assert.equal(layout.totalComponentStartColumn, 109);
  assert.equal(layout.totalDisplayColumn, 113);
  assert.equal(layout.sortKeyColumn, 114);
  assert.equal(layout.sortedRestaurantIdColumn, 115);
  assert.equal(layout.periodKeyStartColumn, 117);
  assert.equal(layout.matrixEndColumn, 10);
});

test("ten-group runtime is deterministic and keeps the empty group opt-in", () => {
  const catalog = groups(10);
  const active = activeWeeklyPerformanceReportingGroups([...catalog].reverse());
  const cacheGroups = activeReportingGroups([...catalog].reverse());
  const selection = planWeeklyPerformanceRpgSelection({
    reportingGroups: catalog,
    priorCatalogExists: true,
    priorRows: groups(9).map(group => ({ id: group.reportingGroupId, include: "Yes" })),
  });
  const layout = buildWeeklyPerformanceLayout({ reportingGroups: catalog, restaurantCapacity: 16 });

  assert.deepEqual(active.map(row => row.reportingGroupId), catalog.map(row => row.reportingGroupId));
  assert.deepEqual(cacheGroups.map(row => row.reportingGroupId), catalog.map(row => row.reportingGroupId));
  assert.equal(selection.length, 10);
  assert.ok(selection.slice(0, 9).every(row => row.include === "Yes"));
  assert.equal(selection[9].include, "No");
  assert.equal(layout.groupCapacity, 10);
  assert.equal(layout.matrixEndColumn, 11);
});

test("cache, publisher, and current Performance runtime contain no fixed-nine capacity guard", () => {
  const currentRuntime = [performanceScript, publisherScript, cacheSource, publisherSource].join("\n");
  assert.doesNotMatch(currentRuntime, /activeGroups\.length\s*(?:!==|===)\s*9/);
  assert.doesNotMatch(currentRuntime, /groups\.length\s*(?:!==|===)\s*9/);
  assert.doesNotMatch(currentRuntime, /GROUP_CAPACITY\s*=\s*9/);
  assert.match(currentRuntime, /At least one active Reporting Group/);
  assert.match(currentRuntime, /repeat(?:s)? SortOrder/);
  assert.match(currentRuntime, /scopeCacheRows\.length \* activeGroups\.length|scopes\.length \* activeGroups\.length/);
});

test("current Performance installer resizes by active IDs and keeps legacy result authority protected", () => {
  assert.match(performanceScript, /table\.resize\(target\)/);
  assert.match(performanceScript, /prior\.includeById\[group\.id\]/);
  assert.match(performanceScript, /prior\.exists \? "No" : "Yes"/);
  assert.match(performanceScript, /makeLayout\(activeGroups\.length/);
  assert.match(performanceScript, /All \$\{layout\.groupCapacity\} Reporting Groups selected/);
  assert.match(performanceScript, /"tblMetricRPGResults",\s*rangeFingerprint/);
  assert.match(performanceScript, /reports\.getRange\("B8"\)\.setFormula\("=Performance!B13"\)/);
  assert.match(performanceScript, /reports\.getRange\("B9"\)\.setFormula\("=Performance!G13"\)/);
});

test("duplicate IDs, duplicate active SortOrder, and zero active groups fail deterministically", () => {
  assert.throws(() => activeReportingGroups([]), /at least one active Reporting Group/i);
  assert.throws(() => activeWeeklyPerformanceReportingGroups([
    ...groups(1), { ...groups(1)[0] },
  ]), /repeats ReportingGroupID/);
  assert.throws(() => activeReportingGroups([
    ...groups(1), { ...groups(1)[0], reportingGroupId: "RPG-0002" },
  ]), /repeat SortOrder/);
});

function groups(count) {
  return Array.from({ length: count }, (_, index) => ({
    reportingGroupId: `RPG-${String(index + 1).padStart(4, "0")}`,
    reportingGroupName: index === 9 ? "Kids Menu" : `Group ${index + 1}`,
    active: "Yes",
    sortOrder: (index + 1) * 10,
  }));
}
