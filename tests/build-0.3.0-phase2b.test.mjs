import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";
import {
  ALL_CHANNELS_SCOPE,
  REPORTING_GROUP_SALES_SHARE_METRIC_ID,
  buildReportingGroupMetricResults,
  deterministicMetricResultId,
  deterministicRestaurantScopeFingerprint,
  summarizePerformanceRestaurantScope,
  validateMetricBridgeForCutover,
  validateReportingGroupMetricResults,
} from "../src/metrics/reporting-group-sales-share.mjs";

const expected = JSON.parse(fs.readFileSync(
  new URL("./expected-build-0.3.0-phase2b.json", import.meta.url),
  "utf8",
));

const mappingAsOfDate = 46246;
const mappingFingerprint = "MAP-1234567890abcdef";
const imports = [
  { importId: "IMP-COMPARE", periodStart: 1, periodEnd: 365 },
  { importId: "IMP-CURRENT", periodStart: 366, periodEnd: 372 },
];
const groups = [
  { reportingGroupId: "RPG-0002", reportingGroupName: "Drinks", sortOrder: 20 },
  { reportingGroupId: "RPG-0001", reportingGroupName: "Add-ons", sortOrder: 10 },
];
const restaurants = [
  { restaurantId: "RST-2", restaurantName: "Second" },
  { restaurantId: "RST-1", restaurantName: "First" },
];
const facts = [
  fact("F-1", "IMP-CURRENT", "RST-1", "PRD-1", 2, 100),
  fact("F-2", "IMP-CURRENT", "RST-1", "PRD-2", 1, 50),
  fact("F-3", "IMP-CURRENT", "RST-1", "PRD-3", 3, 30),
  fact("F-4", "IMP-CURRENT", "RST-2", "PRD-4", 4, 40),
  fact("F-5", "IMP-CURRENT", "RST-2", "PRD-5", 5, 20),
  fact("F-6", "IMP-COMPARE", "RST-1", "PRD-1", 1, 80),
  fact("F-7", "IMP-CURRENT", "RST-1", "PRD-2", 10, 999, "Superseded"),
];
const bridge = [
  bridgeFact(facts[0], "Mapped", "RPG-0001", "Inherited"),
  bridgeFact(facts[1], "Mapped", "RPG-0002", "Explicit"),
  bridgeFact(facts[2], "Unmapped", "", "Explicit exclusion"),
  bridgeFact(facts[3], "Conflict", "", "Explicit conflict"),
  bridgeFact(facts[4], "Inactive Target", "RPG-0003", "Explicit"),
  bridgeFact(facts[5], "Mapped", "RPG-0001", "Inherited"),
  bridgeFact(facts[6], "Mapped", "RPG-0002", "Explicit"),
];

const results = buildReportingGroupMetricResults({
  bridge,
  activeImports: imports,
  activeReportingGroups: groups,
  reportingRestaurants: restaurants,
  mappingAsOfDate,
  mappingFingerprint,
  calculatedAt: 46246.5,
});

test("Phase 2B bridge validation is one-to-one, state-complete, and freshness-aware", () => {
  assert.deepEqual(validateMetricBridgeForCutover({
    facts,
    bridge,
    expectedAsOfDate: mappingAsOfDate,
    expectedMappingFingerprint: mappingFingerprint,
  }), []);

  const staleFingerprint = bridge.map(row => ({ ...row }));
  staleFingerprint[0].mappingFingerprint = "MAP-stale";
  assert.match(validateMetricBridgeForCutover({
    facts,
    bridge: staleFingerprint,
    expectedAsOfDate: mappingAsOfDate,
    expectedMappingFingerprint: mappingFingerprint,
  }).join(" "), /mapping fingerprint MAP-stale/);

  const staleDate = bridge.map(row => ({ ...row, mappingAsOfDate: mappingAsOfDate - 1 }));
  assert.match(validateMetricBridgeForCutover({
    facts,
    bridge: staleDate,
    expectedAsOfDate: mappingAsOfDate,
    expectedMappingFingerprint: mappingFingerprint,
  }).join(" "), /MappingAsOfDate/);

  const missing = bridge.slice(1);
  assert.match(validateMetricBridgeForCutover({
    facts,
    bridge: missing,
    expectedAsOfDate: mappingAsOfDate,
    expectedMappingFingerprint: mappingFingerprint,
  }).join(" "), /row count|missing SalesFactID F-1/);

  const duplicate = [...bridge, { ...bridge[0] }];
  assert.match(validateMetricBridgeForCutover({
    facts,
    bridge: duplicate,
    expectedAsOfDate: mappingAsOfDate,
    expectedMappingFingerprint: mappingFingerprint,
  }).join(" "), /duplicate SalesFactID F-1|duplicate bridge fact/);
});

test("central result grain is deterministic across datasets, groups, and scopes", () => {
  assert.equal(results.length, 2 * 2 * 3);
  assert.equal(new Set(results.map(row => row.metricResultId)).size, results.length);
  assert.equal(new Set(results.map(row => [
    row.metricId, row.importId, row.reportingGroupId, row.scopeType, row.restaurantId,
  ].join("|"))).size, results.length);
  assert.ok(results.every(row => row.metricId === REPORTING_GROUP_SALES_SHARE_METRIC_ID));
  assert.ok(results.every(row => row.channelScope === ALL_CHANNELS_SCOPE));

  const reordered = buildReportingGroupMetricResults({
    bridge: [...bridge].reverse(),
    activeImports: [...imports].reverse(),
    activeReportingGroups: [...groups].reverse(),
    reportingRestaurants: [...restaurants].reverse(),
    mappingAsOfDate,
    mappingFingerprint,
    calculatedAt: 46246.5,
  });
  assert.deepEqual(reordered, results);
  assert.deepEqual(validateReportingGroupMetricResults({
    bridge,
    results,
    activeImports: imports,
    activeReportingGroups: groups,
    reportingRestaurants: restaurants,
  }), []);
});

test("numerator uses only Mapped selected RPG while every state remains in denominator", () => {
  const companyAddOns = result("IMP-CURRENT", "RPG-0001", "Company", "");
  assert.equal(companyAddOns.numeratorSalesNok, 100);
  assert.equal(companyAddOns.denominatorSalesNok, 240);
  assert.equal(companyAddOns.metricValue, 100 / 240);

  const companyDrinks = result("IMP-CURRENT", "RPG-0002", "Company", "");
  assert.equal(companyDrinks.numeratorSalesNok, 50);
  assert.equal(companyDrinks.denominatorSalesNok, 240);

  const firstRestaurant = result("IMP-CURRENT", "RPG-0001", "Restaurant", "RST-1");
  assert.equal(firstRestaurant.numeratorSalesNok, 100);
  assert.equal(firstRestaurant.denominatorSalesNok, 180);

  const secondRestaurant = result("IMP-CURRENT", "RPG-0001", "Restaurant", "RST-2");
  assert.equal(secondRestaurant.numeratorSalesNok, 0);
  assert.equal(secondRestaurant.denominatorSalesNok, 60);

  const comparison = result("IMP-COMPARE", "RPG-0001", "Company", "");
  assert.equal(comparison.numeratorSalesNok, 80);
  assert.equal(comparison.denominatorSalesNok, 80);
});

test("deselected restaurant is excluded from Company numerator and denominator", () => {
  const enabledRestaurants = restaurants.filter(row => row.restaurantId === "RST-1");
  const scopedResults = buildReportingGroupMetricResults({
    bridge,
    activeImports: imports,
    activeReportingGroups: groups,
    reportingRestaurants: enabledRestaurants,
    mappingAsOfDate,
    mappingFingerprint,
    calculatedAt: 46246.5,
  });
  const companyAddOns = scopedResults.find(row =>
    row.importId === "IMP-CURRENT" &&
    row.reportingGroupId === "RPG-0001" &&
    row.scopeType === "Company"
  );
  assert.ok(companyAddOns);
  assert.equal(companyAddOns.numeratorSalesNok, 100);
  assert.equal(companyAddOns.denominatorSalesNok, 180);
  assert.equal(companyAddOns.metricValue, 100 / 180);

  const companyDrinks = scopedResults.find(row =>
    row.importId === "IMP-CURRENT" &&
    row.reportingGroupId === "RPG-0002" &&
    row.scopeType === "Company"
  );
  assert.ok(companyDrinks);
  assert.equal(companyDrinks.numeratorSalesNok, 50);
  assert.equal(companyDrinks.denominatorSalesNok, 180);
  assert.ok(scopedResults.every(row => row.restaurantId !== "RST-2"));
  assert.equal(scopedResults.length, 2 * 2 * (1 + 1));

  const expectedScopeFingerprint = deterministicRestaurantScopeFingerprint(enabledRestaurants);
  assert.ok(scopedResults.filter(row => row.scopeType === "Company")
    .every(row => row.restaurantScopeFingerprint === expectedScopeFingerprint));
  assert.ok(scopedResults.filter(row => row.scopeType === "Restaurant")
    .every(row => row.restaurantScopeFingerprint === ""));
});

test("enabled and excluded Performance restaurant scopes reconcile to the complete bridge", () => {
  const coverage = summarizePerformanceRestaurantScope({
    bridge,
    reportingRestaurants: restaurants.filter(row => row.restaurantId === "RST-1"),
  });
  assert.equal(coverage.enabled.factCount, 5);
  assert.equal(coverage.excluded.factCount, 2);
  assert.equal(coverage.complete.factCount, bridge.length);
  assert.equal(coverage.complete.salesNok, sum(bridge, "salesAmount"));
  assert.equal(coverage.complete.quantity, sum(bridge, "quantity"));
  assert.equal(coverage.enabled.salesNok + coverage.excluded.salesNok, coverage.complete.salesNok);
  assert.equal(coverage.enabled.quantity + coverage.excluded.quantity, coverage.complete.quantity);
});

test("explicit exclusion remains Unmapped and contributes to company and restaurant denominators", () => {
  assert.equal(bridge[2].resolutionState, "Explicit exclusion");
  assert.equal(bridge[2].resolutionStatus, "Unmapped");
  assert.equal(result("IMP-CURRENT", "RPG-0001", "Restaurant", "RST-1").denominatorSalesNok, 180);
  assert.equal(result("IMP-CURRENT", "RPG-0002", "Restaurant", "RST-1").denominatorSalesNok, 180);
});

test("deterministic result IDs depend on stable IDs and scope rather than display names", () => {
  const input = {
    metricId: "KPI-0001",
    importId: "IMP-CURRENT",
    reportingGroupId: "RPG-0001",
    scopeType: "Company",
    restaurantId: "",
    restaurantScopeFingerprint: deterministicRestaurantScopeFingerprint(restaurants),
  };
  assert.match(deterministicMetricResultId(input), /^MRR-[0-9a-f]{16}$/);
  assert.equal(deterministicMetricResultId(input), deterministicMetricResultId({ ...input }));
  assert.notEqual(deterministicMetricResultId(input), deterministicMetricResultId({
    ...input,
    reportingGroupId: "RPG-0002",
  }));
  assert.notEqual(deterministicMetricResultId(input), deterministicMetricResultId({
    ...input,
    restaurantScopeFingerprint: deterministicRestaurantScopeFingerprint(restaurants.slice(0, 1)),
  }));
});

test("accepted source and mapping-state totals reconcile exactly", () => {
  const states = Object.values(expected.states);
  assert.equal(sum(states, "fact_count"), expected.source.fact_count);
  assert.ok(almostEqual(sum(states, "sales_nok"), expected.source.sales_nok));
  assert.ok(almostEqual(sum(states, "quantity"), expected.source.quantity));
  assert.equal(
    expected.result_grain.kpis * expected.result_grain.datasets *
      expected.result_grain.active_reporting_groups *
      (expected.result_grain.company_scopes + expected.result_grain.restaurant_scopes),
    expected.result_grain.expected_rows,
  );
});

test("all nine accepted Reporting Group numerators are independently reproduced by dataset", () => {
  const fixtureBridge = [];
  const fixtureImports = [];
  for (const [importId, dataSet] of Object.entries(expected.datasets)) {
    fixtureImports.push({ importId, periodStart: importId.includes("2025") ? 1 : 366, periodEnd: importId.includes("2025") ? 365 : 372 });
    for (const [reportingGroupId, group] of Object.entries(dataSet.reporting_groups)) {
      fixtureBridge.push(bridgeFact(
        fact(`${importId}-${reportingGroupId}`, importId, "RST-FIXTURE", reportingGroupId, group.quantity, group.sales_nok),
        "Mapped",
        reportingGroupId,
        "Fixture",
      ));
    }
    fixtureBridge.push(bridgeFact(
      fact(`${importId}-UNMAPPED`, importId, "RST-FIXTURE", "PRD-UNMAPPED", dataSet.unmapped.quantity, dataSet.unmapped.sales_nok),
      "Unmapped",
      "",
      "Unmapped",
    ));
  }
  const fixtureGroups = Object.entries(expected.datasets["IMP-2025-BASELINE"].reporting_groups)
    .map(([reportingGroupId, group], index) => ({
      reportingGroupId,
      reportingGroupName: group.name,
      sortOrder: index + 1,
    }));
  const fixtureRestaurants = [{ restaurantId: "RST-FIXTURE", restaurantName: "Fixture" }];
  for (let index = 2; index <= 16; index += 1) {
    fixtureRestaurants.push({
      restaurantId: `RST-${String(index).padStart(2, "0")}`,
      restaurantName: `Restaurant ${String(index).padStart(2, "0")}`,
    });
  }
  const fixtureResults = buildReportingGroupMetricResults({
    bridge: fixtureBridge,
    activeImports: fixtureImports,
    activeReportingGroups: fixtureGroups,
    reportingRestaurants: fixtureRestaurants,
    mappingAsOfDate,
    mappingFingerprint,
    calculatedAt: 46246.5,
  });
  assert.equal(fixtureResults.length, expected.result_grain.expected_rows);
  const fifteenRestaurantResults = buildReportingGroupMetricResults({
    bridge: fixtureBridge,
    activeImports: fixtureImports,
    activeReportingGroups: fixtureGroups,
    reportingRestaurants: fixtureRestaurants.slice(0, 15),
    mappingAsOfDate,
    mappingFingerprint,
    calculatedAt: 46246.5,
  });
  assert.equal(fifteenRestaurantResults.length, 2 * 9 * (1 + 15));

  for (const [importId, dataSet] of Object.entries(expected.datasets)) {
    for (const [reportingGroupId, group] of Object.entries(dataSet.reporting_groups)) {
      const row = fixtureResults.find(value =>
        value.importId === importId && value.reportingGroupId === reportingGroupId,
      );
      assert.ok(row, `${importId}/${reportingGroupId} result missing`);
      assert.ok(almostEqual(row.numeratorSalesNok, group.sales_nok));
      assert.ok(almostEqual(row.denominatorSalesNok, dataSet.sales_nok));
      assert.ok(almostEqual(row.metricValue, group.sales_nok / dataSet.sales_nok));
    }
  }
});

test("accepted Add-ons Performance targets are intentional RPG results", () => {
  const current = expected.datasets[expected.default_add_ons.current_import_id];
  const comparison = expected.datasets[expected.default_add_ons.comparison_import_id];
  assert.equal(current.reporting_groups["RPG-0001"].sales_nok, expected.default_add_ons.current_sales_nok);
  assert.ok(almostEqual(
    current.reporting_groups["RPG-0001"].sales_nok / current.sales_nok,
    expected.default_add_ons.current_share,
  ));
  assert.equal(comparison.reporting_groups["RPG-0001"].sales_nok, expected.default_add_ons.comparison_sales_nok);
  assert.ok(almostEqual(
    comparison.reporting_groups["RPG-0001"].sales_nok / comparison.sales_nok,
    expected.default_add_ons.comparison_share,
  ));
  assert.ok(almostEqual(
    expected.default_add_ons.current_share - expected.default_add_ons.comparison_share,
    expected.default_add_ons.change_percentage_points,
  ));
});

test("Phase 2B Office Script validates before mutation and centralizes Performance results", () => {
  const script = fs.readFileSync(
    new URL("../office-scripts/Build_0_3_0_Phase2B.ts", import.meta.url),
    "utf8",
  );
  assert.match(script, /PUL-0302B-001: Phase 2A metric bridge is stale or invalid/);
  assert.ok(script.indexOf("validateMetricBridgeForCutover(") < script.indexOf("writeMetricCalc("));
  assert.ok(script.indexOf("validateCentralResults(") < script.indexOf("writeMetricCalc("));
  assert.match(script, /tblMetricRPGResults/);
  assert.match(script, /"KPI-0001"/);
  assert.match(script, /"Reporting Group Sales Share"/);
  assert.doesNotMatch(script, /tblSalesFacts\[ReportingCategoryID\]/);
  assert.doesNotMatch(script, /tblSalesFacts\[SalesAmount\]/);
  assert.match(script, /snapshotProtectedTables[\s\S]*tblReportingCategories/);
  assert.match(script, /metricResultFormula\("MetricValue"/);
  assert.match(script, /metricResultFormula\("NumeratorSalesNOK"/);
  assert.match(script, /SUMIFS\(tblMetricRPGResults\[\$\{valueColumn\}\]/);
  assert.match(script, /"Reporting Group", "ReportingGroupID"/);
  assert.match(script, /"ScopeType", "RestaurantID", "RestaurantScopeFingerprint", "ChannelScope"/);
  assert.match(script, /groups\.filter\(group => group\.active === "Yes"\)/);
  assert.equal(
    (script.match(/if \(!restaurantIds\.has\(row\.restaurantId\)\) continue;/g) ?? []).length,
    2,
    "enabled-restaurant filter must guard both aggregation and independent validation",
  );
  assert.match(script, /deterministicRestaurantScopeFingerprint\(restaurants\)/);
  assert.match(script, /updateTableSnapshot\(state, "tblRestaurants"/);
  assert.match(script, /QA-0302B-14/);
  assert.match(script, /QA-0302B-15/);
  assert.match(script, /QA-0302B-16/);
  assert.doesNotMatch(script, /readEquivalenceDefinitions|compareLegacyDefinitions/);
});

test("Phase 2B preserves all-channel UI behavior and Reports linkage", () => {
  const script = fs.readFileSync(
    new URL("../office-scripts/Build_0_3_0_Phase2B.ts", import.meta.url),
    "utf8",
  );
  assert.doesNotMatch(script, /getRange\("B11"\)/);
  assert.doesNotMatch(script, /getRange\("G11"\)/);
  assert.doesNotMatch(script, /\["Channel",/);
  assert.match(script, /"All channels"/);
  assert.match(script, /reports\.getRange\("B7"\)\.setFormula\("=Performance!B7"\)/);
  assert.match(script, /reports\.getRange\("B15"\)\.setFormula\("=Performance!B16"\)/);
  assert.match(script, /reports\.getRange\("B18"\)\.setFormula\("=Performance!B19"\)/);
});

test("Phase 2B Office Script avoids unsupported iterator constructs and read calls inside loops", () => {
  const script = fs.readFileSync(
    new URL("../office-scripts/Build_0_3_0_Phase2B.ts", import.meta.url),
    "utf8",
  );
  assert.doesNotMatch(script, /Array\.from\s*\(/);
  assert.doesNotMatch(script, /\.(?:entries|keys|values)\s*\(/);
  assert.doesNotMatch(script, /\.\.\.(?:aggregates|mappedByGroup|factsById|bridgeById)\b/);
  const loops = [...script.matchAll(/for\s*\([^)]*\)\s*\{[\s\S]*?\n\s*\}/g)].map(match => match[0]);
  assert.ok(loops.every(loop => !/getRangeBetweenHeaderAndTotal\(\)\.getValues\(\)/.test(loop)));
  assert.match(script, /groupFormulas[\s\S]*?setFormulas\(groupFormulas\)/);
  assert.match(script, /restaurantFormulas[\s\S]*?setFormulas\(restaurantFormulas\)/);
});

function result(importId, reportingGroupId, scopeType, restaurantId) {
  const row = results.find(value =>
    value.importId === importId &&
    value.reportingGroupId === reportingGroupId &&
    value.scopeType === scopeType &&
    value.restaurantId === restaurantId,
  );
  assert.ok(row, `Missing ${importId}/${reportingGroupId}/${scopeType}/${restaurantId}`);
  return row;
}

function fact(salesFactId, importId, restaurantId, productId, quantity, salesAmount, publicationState = "Active Finalized") {
  return {
    salesFactId,
    importId,
    restaurantId,
    productId,
    reportingChannel: "In-house",
    periodStart: 1,
    periodEnd: 2,
    quantity,
    salesAmount,
    publicationState,
  };
}

function bridgeFact(source, resolutionStatus, effectiveReportingGroupId, resolutionState) {
  return {
    ...source,
    effectiveReportingGroupId,
    resolutionState,
    resolutionStatus,
    mappingAsOfDate,
    mappingFingerprint,
  };
}

function sum(rows, field) {
  return rows.reduce((total, row) => total + row[field], 0);
}

function almostEqual(left, right) {
  return Math.abs(left - right) <= 0.000001;
}
