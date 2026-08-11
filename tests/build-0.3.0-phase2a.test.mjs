import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";
import {
  aggregateReportingGroups,
  buildLegacyRpgCrosswalk,
  buildMetricBridge,
  compareLegacyDefinitions,
  computeMappingFingerprint,
  reconcileFactsAndBridge,
  validateEffectiveMappingFreshness,
  validateEquivalenceDefinitions,
} from "../src/metrics/reporting-group-metrics.mjs";

const expected = JSON.parse(fs.readFileSync(
  new URL("./expected-build-0.3.0-phase2a.json", import.meta.url),
  "utf8",
));

const asOfDate = expected.mapping_as_of_date;
const groups = [
  { reportingGroupId: "RPG-0001", reportingGroupName: "Add-ons", active: "Yes", sortOrder: 10 },
  { reportingGroupId: "RPG-0002", reportingGroupName: "Drinks", active: "Yes", sortOrder: 20 },
  { reportingGroupId: "RPG-0003", reportingGroupName: "Retired", active: "No", sortOrder: 30 },
];
const rules = [
  { mappingRuleId: "MAP-1", sourceSystemId: "SRC-1", scopeType: "SourceMainCategory", nodeId: "MAIN-A", targetReportingGroupId: "RPG-0001", effectiveFrom: 1, effectiveTo: null, status: "Active" },
  { mappingRuleId: "MAP-2", sourceSystemId: "SRC-1", scopeType: "Product", nodeId: "PRD-2", targetReportingGroupId: "RPG-0002", effectiveFrom: 1, effectiveTo: null, status: "Active" },
  { mappingRuleId: "MAP-3", sourceSystemId: "SRC-1", scopeType: "Product", nodeId: "PRD-4", targetReportingGroupId: "RPG-0001", effectiveFrom: 1, effectiveTo: null, status: "Active" },
  { mappingRuleId: "MAP-4", sourceSystemId: "SRC-1", scopeType: "Product", nodeId: "PRD-4", targetReportingGroupId: "RPG-0002", effectiveFrom: 1, effectiveTo: null, status: "Active" },
  { mappingRuleId: "MAP-5", sourceSystemId: "SRC-1", scopeType: "Product", nodeId: "PRD-5", targetReportingGroupId: "RPG-0003", effectiveFrom: 1, effectiveTo: null, status: "Active" },
];
const products = [
  { productId: "PRD-1", sourceSystemId: "SRC-1", mainNodeId: "MAIN-A", subNodeId: "SUB-A" },
  { productId: "PRD-2", sourceSystemId: "SRC-1", mainNodeId: "MAIN-B", subNodeId: "SUB-B" },
  { productId: "PRD-3", sourceSystemId: "SRC-1", mainNodeId: "MAIN-C", subNodeId: "SUB-C" },
  { productId: "PRD-4", sourceSystemId: "SRC-1", mainNodeId: "MAIN-D", subNodeId: "SUB-D" },
  { productId: "PRD-5", sourceSystemId: "SRC-1", mainNodeId: "MAIN-E", subNodeId: "SUB-E" },
];
const resolutions = [
  resolution("PRD-1", "RPG-0001", "Add-ons", "SourceMainCategory", "Inherited", "Mapped", "MAP-1", "MAIN-A", "SUB-A", { mainRuleIds: "MAP-1", mainTargetIds: "RPG-0001" }),
  resolution("PRD-2", "RPG-0002", "Drinks", "Product", "Explicit", "Mapped", "MAP-2", "MAIN-B", "SUB-B", { productRuleIds: "MAP-2", productTargetIds: "RPG-0002" }),
  resolution("PRD-3", "", "", "Unmapped", "Unmapped", "Unmapped", "", "MAIN-C", "SUB-C"),
  resolution("PRD-4", "", "", "Product", "Explicit conflict", "Conflict", "MAP-3, MAP-4", "MAIN-D", "SUB-D", { productRuleIds: "MAP-3, MAP-4", productTargetIds: "RPG-0001, RPG-0002" }),
  resolution("PRD-5", "RPG-0003", "Retired", "Product", "Explicit", "Inactive Target", "MAP-5", "MAIN-E", "SUB-E", { productRuleIds: "MAP-5", productTargetIds: "RPG-0003" }),
];
const materialized = resolutions.map(row => ({ ...row, asOfDate }));
const facts = [
  fact("F-1", "IMP-CURRENT", "PRD-1", "CAT-A", "In-house", 2, 100),
  fact("F-2", "IMP-CURRENT", "PRD-2", "CAT-B", "In-house", 1, 50),
  fact("F-3", "IMP-CURRENT", "PRD-3", "CAT-C", "In-house", 3, 30),
  fact("F-4", "IMP-CURRENT", "PRD-4", "CAT-B", "Takeaway", 4, 40),
  fact("F-5", "IMP-CURRENT", "PRD-5", "CAT-D", "Takeaway", 5, 20),
  fact("F-6", "IMP-COMPARE", "PRD-1", "CAT-A", "In-house", 1, 80),
  fact("F-7", "IMP-CURRENT", "PRD-2", "CAT-B", "In-house", 10, 999, "Superseded"),
];

const fingerprint = computeMappingFingerprint({ asOfDate, groups, rules, products, resolutions });
const bridge = buildMetricBridge({
  facts,
  effectiveMappings: materialized,
  mappingAsOfDate: asOfDate,
  mappingFingerprint: fingerprint,
  metricRefreshAt: 46245.5,
});
const currentScope = {
  publicationState: "Active Finalized",
  importId: "IMP-CURRENT",
  channel: "All channels",
};

test("mapping fingerprint is stable across source row ordering", () => {
  const reordered = computeMappingFingerprint({
    asOfDate,
    groups: [...groups].reverse(),
    rules: [...rules].reverse(),
    products: [...products].reverse(),
    resolutions: [...resolutions].reverse(),
  });
  assert.equal(fingerprint, expected.mapping_fingerprint);
  assert.equal(reordered, fingerprint);

  const reorderedConflictMembers = resolutions.map(row => row.productId === "PRD-4"
    ? {
      ...row,
      productRuleIds: "MAP-4, MAP-3",
      productTargetIds: "RPG-0002, RPG-0001",
      winningRuleId: "MAP-4, MAP-3",
    }
    : row);
  assert.equal(computeMappingFingerprint({
    asOfDate,
    groups,
    rules,
    products,
    resolutions: reorderedConflictMembers,
  }), fingerprint);
});

test("fresh Effective Mapping validates and stale mapping is rejected", () => {
  assert.deepEqual(validateEffectiveMappingFreshness({
    expected: resolutions,
    materialized,
    expectedAsOfDate: asOfDate,
  }), []);

  const reorderedConflictMembers = materialized.map(row => row.productId === "PRD-4"
    ? {
      ...row,
      productRuleIds: "MAP-4, MAP-3",
      productTargetIds: "RPG-0002, RPG-0001",
      winningRuleId: "MAP-4, MAP-3",
    }
    : row);
  assert.deepEqual(validateEffectiveMappingFreshness({
    expected: resolutions,
    materialized: reorderedConflictMembers,
    expectedAsOfDate: asOfDate,
  }), []);

  const stale = materialized.map(row => ({ ...row }));
  stale[1].effectiveReportingGroupId = "RPG-0001";
  assert.match(validateEffectiveMappingFreshness({
    expected: resolutions,
    materialized: stale,
    expectedAsOfDate: asOfDate,
  }).join(" "), /stale for product PRD-2/);

  const priorDate = materialized.map(row => ({ ...row, asOfDate: asOfDate - 1 }));
  assert.match(validateEffectiveMappingFreshness({
    expected: resolutions,
    materialized: priorDate,
    expectedAsOfDate: asOfDate,
  }).join(" "), /expected 46245/);
});

test("derived bridge accounts for every source fact and preserves lineage", () => {
  assert.equal(bridge.length, expected.all_facts.fact_count);
  assert.deepEqual(bridge.map(row => row.salesFactId), facts.map(row => row.salesFactId));
  assert.ok(bridge.every(row => row.mappingAsOfDate === asOfDate));
  assert.ok(bridge.every(row => row.mappingFingerprint === fingerprint));
  assert.deepEqual(new Set(bridge.map(row => row.resolutionStatus)), new Set([
    "Mapped", "Unmapped", "Conflict", "Inactive Target",
  ]));
});

test("derived bridge rejects missing or unsupported Effective Mapping membership", () => {
  assert.throws(() => buildMetricBridge({
    facts,
    effectiveMappings: materialized.filter(row => row.productId !== "PRD-3"),
    mappingAsOfDate: asOfDate,
    mappingFingerprint: fingerprint,
    metricRefreshAt: 46245.5,
  }), /without Effective Mapping/);

  const invalid = materialized.map(row => row.productId === "PRD-3"
    ? { ...row, resolutionStatus: "Unknown" }
    : row);
  assert.throws(() => buildMetricBridge({
    facts,
    effectiveMappings: invalid,
    mappingAsOfDate: asOfDate,
    mappingFingerprint: fingerprint,
    metricRefreshAt: 46245.5,
  }), /unsupported ResolutionStatus Unknown/);
});

test("Sales NOK, Quantity and fact count reconcile for all facts and active scope", () => {
  const rows = reconcileFactsAndBridge({
    facts,
    bridge,
    scopes: [
      { importId: "All imports", channel: "All channels" },
      currentScope,
    ],
  });
  assert.equal(rows[0].result, "PASS");
  assert.deepEqual(rows[0].source, {
    factCount: expected.all_facts.fact_count,
    salesNok: expected.all_facts.sales_nok,
    quantity: expected.all_facts.quantity,
  });
  assert.equal(rows[1].result, "PASS");
  assert.deepEqual(rows[1].source, {
    factCount: expected.current_active_scope.fact_count,
    salesNok: expected.current_active_scope.sales_nok,
    quantity: expected.current_active_scope.quantity,
  });
  assert.deepEqual(rows[1].coverage, Object.fromEntries(
    Object.entries(expected.current_active_scope.states).map(([state, values]) => [state, {
      factCount: values.fact_count,
      salesNok: values.sales_nok,
      quantity: values.quantity,
    }]),
  ));
});

test("Reporting Group aggregation has independently expected fixture results", () => {
  const totals = aggregateReportingGroups(bridge, currentScope);
  assert.deepEqual(totals, [
    { reportingGroupId: "RPG-0001", reportingGroupName: "Add-ons", factCount: 1, salesNok: 100, quantity: 2 },
    { reportingGroupId: "RPG-0002", reportingGroupName: "Drinks", factCount: 1, salesNok: 50, quantity: 1 },
  ]);
});

test("current mapping changes reclassify analysis without changing fact totals", () => {
  const remapped = materialized.map(row => row.productId === "PRD-2"
    ? { ...row, effectiveReportingGroupId: "RPG-0001", effectiveReportingGroupName: "Add-ons" }
    : row);
  const remappedBridge = buildMetricBridge({
    facts,
    effectiveMappings: remapped,
    mappingAsOfDate: asOfDate,
    mappingFingerprint: "MAP-REMAPPED",
    metricRefreshAt: 46246,
  });
  assert.deepEqual(reconcileFactsAndBridge({ facts, bridge: remappedBridge, scopes: [currentScope] })[0].source,
    reconcileFactsAndBridge({ facts, bridge, scopes: [currentScope] })[0].source);
  assert.deepEqual(aggregateReportingGroups(remappedBridge, currentScope), [
    { reportingGroupId: "RPG-0001", reportingGroupName: "Add-ons", factCount: 2, salesNok: 150, quantity: 3 },
  ]);
});

test("legacy equivalence is explicit and side-by-side variance is visible", () => {
  const definitions = [
    { definitionId: "LEGACY-EQ-001", reportingGroupId: "RPG-0001", legacyReportingCategoryId: "CAT-A", comparisonStatus: "Equivalent", active: "Yes" },
    { definitionId: "LEGACY-EQ-002", reportingGroupId: "RPG-0002", legacyReportingCategoryId: "CAT-B", comparisonStatus: "Equivalent", active: "Yes" },
  ];
  assert.deepEqual(validateEquivalenceDefinitions(definitions, {
    groupIds: groups.map(row => row.reportingGroupId),
    categoryIds: ["CAT-A", "CAT-B", "CAT-C", "CAT-D"],
  }), []);
  const comparisons = compareLegacyDefinitions({ bridge, definitions, scopes: [currentScope] });
  assert.equal(comparisons[0].result, expected.legacy_comparisons["LEGACY-EQ-001"].result);
  assert.equal(comparisons[0].legacySalesNok, expected.legacy_comparisons["LEGACY-EQ-001"].legacy_sales_nok);
  assert.equal(comparisons[0].rpgSalesNok, expected.legacy_comparisons["LEGACY-EQ-001"].rpg_sales_nok);
  assert.equal(comparisons[1].result, expected.legacy_comparisons["LEGACY-EQ-002"].result);
  assert.equal(comparisons[1].legacySalesNok, expected.legacy_comparisons["LEGACY-EQ-002"].legacy_sales_nok);
  assert.equal(comparisons[1].rpgSalesNok, expected.legacy_comparisons["LEGACY-EQ-002"].rpg_sales_nok);
  assert.equal(comparisons[1].salesVariance, expected.legacy_comparisons["LEGACY-EQ-002"].sales_variance);
});

test("legacy equivalence validation never infers IDs from names", () => {
  const errors = validateEquivalenceDefinitions([
    { definitionId: "LEGACY-EQ-003", reportingGroupId: "Add-ons", legacyReportingCategoryId: "Add-ons", comparisonStatus: "Equivalent", active: "Yes" },
  ], {
    groupIds: groups.map(row => row.reportingGroupId),
    categoryIds: ["CAT-A", "CAT-B"],
  });
  assert.match(errors.join(" "), /unknown ReportingGroupID Add-ons/);
  assert.match(errors.join(" "), /unknown legacy category Add-ons/);
});

test("legacy/RPG crosswalk exposes conflict and inactive-target membership", () => {
  const rows = buildLegacyRpgCrosswalk(bridge, [currentScope]);
  assert.ok(rows.some(row => row.legacyReportingCategoryId === "CAT-B" && row.resolutionStatus === "Conflict" && row.salesNok === 40));
  assert.ok(rows.some(row => row.legacyReportingCategoryId === "CAT-D" && row.resolutionStatus === "Inactive Target" && row.salesNok === 20));
});

test("Phase 2A Office Script preserves legacy metric surfaces and guards freshness before output", () => {
  const script = fs.readFileSync(
    new URL("../office-scripts/Build_0_3_0_Phase2A.ts", import.meta.url),
    "utf8",
  );
  assert.match(script, /PUL-0302A-001: Effective Mapping is stale or invalid/);
  assert.ok(
    script.indexOf("validateEffectiveMappingFreshness(") < script.indexOf("writeMetricContract(workbook)"),
    "freshness validation must precede Phase 2A output writes",
  );
  assert.doesNotMatch(script, /resetOutputSheet\(\s*workbook,\s*"(?:_Metric_Calc|Performance|Reports)"/);
  assert.match(script, /snapshotProtectedSurfaces\(workbook, kpiTable\)/);
  assert.match(script, /_Metric_Calc, Performance, Reports, or KPI Registry changed during Phase 2A/);
  assert.match(script, /"MappingAsOfDate", "MappingFingerprint", "MetricRefreshAt"/);
  assert.match(script, /"Mapped", "Unmapped", "Conflict", "Inactive Target"/);
});

test("Phase 2A Office Script avoids unsupported Map and Set iterator constructs", () => {
  const script = fs.readFileSync(
    new URL("../office-scripts/Build_0_3_0_Phase2A.ts", import.meta.url),
    "utf8",
  );
  const collectionNames = "classificationById|groupById|actualByProduct|mappingByProduct|byDefinition|channelsByImport|byGroup|map|active|imports|channelSet|groupIds|knownCategories|allowedStatus|legacyIds|seen";
  assert.doesNotMatch(script, /Array\.from\s*\(/);
  assert.doesNotMatch(script, /\.(?:entries|keys|values)\s*\(/);
  assert.doesNotMatch(script, new RegExp(`for\\s*\\([^)]*\\bof\\s+(?:${collectionNames})\\b`));
  assert.doesNotMatch(script, new RegExp(`\\.\\.\\.(?:${collectionNames})\\b`));
  assert.match(script, /byGroup\.forEach\s*\(/);
  assert.match(script, /map\.forEach\s*\(/);
  assert.match(script, /active\.forEach\s*\(/);
});

test("Phase 2A Office Script uses bounded indexed output ranges with actionable diagnostics", () => {
  const script = fs.readFileSync(
    new URL("../office-scripts/Build_0_3_0_Phase2A.ts", import.meta.url),
    "utf8",
  );
  assert.doesNotMatch(script, /\.getRange\s*\(/, "string-address Worksheet.getRange calls must not return");
  assert.match(script, /function checkedRangeByIndexes\s*\(/);
  assert.match(script, /PUL-0302A-020: Invalid worksheet range/);
  assert.match(script, /PUL-0302A-021: Worksheet range acquisition failed at/);
  assert.match(script, /sheet\.addTable\(tableRange, true\)/);
  assert.match(script, /validateOutputRows\(rows, headers\.length, sheetName, tableName\)/);

  const bridgeStartRow = 4;
  const checkpointFactRows = 18086;
  assert.equal(bridgeStartRow + checkpointFactRows, 18090);
  assert.ok(18090 <= 1048576, "checkpoint bridge range must fit on an Excel worksheet");
});

test("Phase 2A protected-surface reads are batched outside iteration", () => {
  const script = fs.readFileSync(
    new URL("../office-scripts/Build_0_3_0_Phase2A.ts", import.meta.url),
    "utf8",
  );
  const snapshotBody = script.match(
    /function snapshotProtectedSurfaces[\s\S]*?\n}\n\nfunction updateProtectedSnapshot/,
  )?.[0];
  assert.ok(snapshotBody, "snapshotProtectedSurfaces source must be present");
  assert.doesNotMatch(snapshotBody, /for\s*\(/);
  assert.match(snapshotBody, /metricCalcUsed\.getValues\(\)/);
  assert.match(snapshotBody, /performanceUsed\.getValues\(\)/);
  assert.match(snapshotBody, /reportsUsed\.getValues\(\)/);

  const resetBody = script.match(
    /function resetOutputSheet[\s\S]*?\n}\n\nfunction requiredSheet/,
  )?.[0];
  assert.ok(resetBody, "resetOutputSheet source must be present");
  assert.doesNotMatch(resetBody, /workbook\.getTable\s*\(/);
});

function resolution(productId, groupId, groupName, source, state, status, winningRuleId, mainNodeId, subNodeId, extra = {}) {
  return {
    productId,
    mainRuleIds: "",
    mainTargetIds: "",
    subRuleIds: "",
    subTargetIds: "",
    productRuleIds: "",
    productTargetIds: "",
    effectiveReportingGroupId: groupId,
    effectiveReportingGroupName: groupName,
    resolutionSource: source,
    resolutionState: state,
    resolutionStatus: status,
    winningRuleId,
    mainNodeId,
    subNodeId,
    ...extra,
  };
}

function fact(salesFactId, importId, productId, categoryId, channel, quantity, salesAmount, publicationState = "Active Finalized") {
  return {
    salesFactId,
    importId,
    restaurantId: "RST-1",
    productId,
    legacyReportingCategoryId: categoryId,
    periodStart: 1,
    periodEnd: 2,
    reportingChannel: channel,
    quantity,
    salesAmount,
    publicationState,
  };
}
