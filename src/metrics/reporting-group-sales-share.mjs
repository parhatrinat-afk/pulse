/**
 * Deterministic Build 0.3.0 Phase 2B Reporting Group Sales Share primitives.
 *
 * The Phase 2A bridge is the only mapping-aware input. These functions do not
 * resolve mapping rules or mutate facts. They validate the bridge contract and
 * materialize centralized KPI results for company and restaurant scopes.
 */

export const REPORTING_GROUP_SALES_SHARE_METRIC_ID = "KPI-0001";
export const ALL_CHANNELS_SCOPE = "All channels";
export const ACTIVE_FINALIZED = "Active Finalized";

const MAPPING_STATES = new Set([
  "Mapped",
  "Unmapped",
  "Conflict",
  "Inactive Target",
]);

export function validateMetricBridgeForCutover({
  facts,
  bridge,
  expectedAsOfDate,
  expectedMappingFingerprint,
}) {
  const errors = [];
  const factsById = uniqueById(facts, "source fact", errors);
  const bridgeById = uniqueById(bridge, "bridge fact", errors);

  if (facts.length !== bridge.length) {
    errors.push(`Bridge row count is ${bridge.length}; expected ${facts.length}.`);
  }

  const sourceTotals = totals(facts);
  const bridgeTotals = totals(bridge);
  if (sourceTotals.factCount !== bridgeTotals.factCount) {
    errors.push("Bridge fact count does not reconcile to source facts.");
  }
  if (!almostEqual(sourceTotals.salesNok, bridgeTotals.salesNok)) {
    errors.push("Bridge Sales NOK does not reconcile to source facts.");
  }
  if (!almostEqual(sourceTotals.quantity, bridgeTotals.quantity)) {
    errors.push("Bridge Quantity does not reconcile to source facts.");
  }

  const coverage = emptyCoverage();
  for (const row of bridge) {
    const salesFactId = text(row.salesFactId);
    const state = text(row.resolutionStatus);
    if (!MAPPING_STATES.has(state)) {
      errors.push(`Bridge fact ${salesFactId || "(blank)"} has unsupported state ${state || "(blank)"}.`);
      continue;
    }
    if (state === "Mapped" && !text(row.effectiveReportingGroupId)) {
      errors.push(`Mapped bridge fact ${salesFactId || "(blank)"} has no ReportingGroupID.`);
    }
    coverage[state].factCount += 1;
    coverage[state].salesNok += number(row.salesAmount);
    coverage[state].quantity += number(row.quantity);

    if (number(row.mappingAsOfDate) !== number(expectedAsOfDate)) {
      errors.push(
        `Bridge fact ${salesFactId || "(blank)"} has MappingAsOfDate ${row.mappingAsOfDate}; expected ${expectedAsOfDate}.`,
      );
    }
    if (text(row.mappingFingerprint) !== text(expectedMappingFingerprint)) {
      errors.push(
        `Bridge fact ${salesFactId || "(blank)"} has mapping fingerprint ${text(row.mappingFingerprint) || "(blank)"}; expected ${expectedMappingFingerprint}.`,
      );
    }

    const source = factsById.get(salesFactId);
    if (!source) {
      errors.push(`Bridge contains unknown SalesFactID ${salesFactId || "(blank)"}.`);
      continue;
    }
    if (factSignature(source) !== factSignature(row)) {
      errors.push(`Bridge lineage or source measures differ for SalesFactID ${salesFactId}.`);
    }
  }

  for (const [salesFactId] of factsById) {
    if (!bridgeById.has(salesFactId)) {
      errors.push(`Bridge is missing SalesFactID ${salesFactId}.`);
    }
  }

  const coverageTotals = Object.values(coverage).reduce((result, row) => ({
    factCount: result.factCount + row.factCount,
    salesNok: result.salesNok + row.salesNok,
    quantity: result.quantity + row.quantity,
  }), { factCount: 0, salesNok: 0, quantity: 0 });
  if (coverageTotals.factCount !== bridgeTotals.factCount ||
      !almostEqual(coverageTotals.salesNok, bridgeTotals.salesNok) ||
      !almostEqual(coverageTotals.quantity, bridgeTotals.quantity)) {
    errors.push("Mapped, Unmapped, Conflict, and Inactive Target coverage does not reconcile to the bridge.");
  }

  return errors;
}

export function buildReportingGroupMetricResults({
  bridge,
  activeImports,
  activeReportingGroups,
  reportingRestaurants,
  mappingAsOfDate,
  mappingFingerprint,
  calculatedAt,
  metricId = REPORTING_GROUP_SALES_SHARE_METRIC_ID,
}) {
  const imports = uniqueAndSort(activeImports, "importId", compareImports);
  const groups = uniqueAndSort(activeReportingGroups, "reportingGroupId", compareGroups);
  const restaurants = uniqueAndSort(reportingRestaurants, "restaurantId", compareRestaurants);
  if (!imports.length) throw new Error("No Active Finalized dataset is available for Phase 2B.");
  if (!groups.length) throw new Error("No active Reporting Group is available for Phase 2B.");

  const importIds = new Set(imports.map(row => text(row.importId)));
  const restaurantIds = new Set(restaurants.map(row => text(row.restaurantId)));
  const companyRestaurantScopeFingerprint = deterministicRestaurantScopeFingerprint(restaurants);
  const aggregates = new Map();

  for (const row of bridge) {
    if (text(row.publicationState) !== ACTIVE_FINALIZED) continue;
    const importId = text(row.importId);
    if (!importIds.has(importId)) continue;
    const restaurantId = text(row.restaurantId);
    if (!restaurantIds.has(restaurantId)) continue;
    addToAggregate(aggregates, scopeKey(importId, "Company", ""), row);
    addToAggregate(aggregates, scopeKey(importId, "Restaurant", restaurantId), row);
  }

  const results = [];
  for (const dataSet of imports) {
    const importId = text(dataSet.importId);
    for (const group of groups) {
      const reportingGroupId = text(group.reportingGroupId);
      results.push(metricResult({
        metricId,
        importId,
        reportingGroupId,
        scopeType: "Company",
        restaurantId: "",
        restaurantScopeFingerprint: companyRestaurantScopeFingerprint,
        aggregate: aggregates.get(scopeKey(importId, "Company", "")),
        mappingAsOfDate,
        mappingFingerprint,
        calculatedAt,
      }));
      for (const restaurant of restaurants) {
        const restaurantId = text(restaurant.restaurantId);
        results.push(metricResult({
          metricId,
          importId,
          reportingGroupId,
          scopeType: "Restaurant",
          restaurantId,
          restaurantScopeFingerprint: "",
          aggregate: aggregates.get(scopeKey(importId, "Restaurant", restaurantId)),
          mappingAsOfDate,
          mappingFingerprint,
          calculatedAt,
        }));
      }
    }
  }
  return results;
}

export function validateReportingGroupMetricResults({
  bridge,
  results,
  activeImports,
  activeReportingGroups,
  reportingRestaurants,
}) {
  const errors = [];
  const expected = buildReportingGroupMetricResults({
    bridge,
    activeImports,
    activeReportingGroups,
    reportingRestaurants,
    mappingAsOfDate: results[0]?.mappingAsOfDate ?? 0,
    mappingFingerprint: results[0]?.mappingFingerprint ?? "",
    calculatedAt: results[0]?.calculatedAt ?? 0,
    metricId: results[0]?.metricId ?? REPORTING_GROUP_SALES_SHARE_METRIC_ID,
  });
  const actualByKey = new Map();
  for (const result of results) {
    const key = resultKey(result);
    if (actualByKey.has(key)) errors.push(`Duplicate centralized metric result ${key}.`);
    actualByKey.set(key, result);
  }
  if (results.length !== expected.length) {
    errors.push(`Central result row count is ${results.length}; expected ${expected.length}.`);
  }
  for (const expectedRow of expected) {
    const actual = actualByKey.get(resultKey(expectedRow));
    if (!actual) {
      errors.push(`Missing centralized metric result ${resultKey(expectedRow)}.`);
      continue;
    }
    if (text(actual.metricResultId) !== text(expectedRow.metricResultId)) {
      errors.push(`Metric result ID is not deterministic for ${resultKey(expectedRow)}.`);
    }
    if (text(actual.restaurantScopeFingerprint) !== text(expectedRow.restaurantScopeFingerprint)) {
      errors.push(`Restaurant scope fingerprint differs for ${resultKey(expectedRow)}.`);
    }
    if (!almostEqual(number(actual.numeratorSalesNok), expectedRow.numeratorSalesNok) ||
        !almostEqual(number(actual.denominatorSalesNok), expectedRow.denominatorSalesNok) ||
        !almostEqual(number(actual.metricValue), expectedRow.metricValue)) {
      errors.push(`Centralized metric values differ for ${resultKey(expectedRow)}.`);
    }
  }
  return errors;
}

export function deterministicMetricResultId({
  metricId,
  importId,
  reportingGroupId,
  scopeType,
  restaurantId,
  restaurantScopeFingerprint = "",
  channelScope = ALL_CHANNELS_SCOPE,
}) {
  return `MRR-${hashText([
    metricId,
    importId,
    reportingGroupId,
    scopeType,
    restaurantId,
    scopeType === "Company" ? restaurantScopeFingerprint : "",
    channelScope,
    ACTIVE_FINALIZED,
  ].map(text).join("|"))}`;
}

export function deterministicRestaurantScopeFingerprint(reportingRestaurants) {
  const ids = reportingRestaurants.map(row => text(row.restaurantId));
  ids.sort();
  const serialized = ids.map(id => `${id.length}:${id}`).join("|");
  return `RSC-${hashText(`ENABLED-RESTAURANTS|${serialized}`)}`;
}

export function summarizePerformanceRestaurantScope({ bridge, reportingRestaurants }) {
  const enabledIds = new Set(reportingRestaurants.map(row => text(row.restaurantId)));
  const enabled = emptyTotals();
  const excluded = emptyTotals();
  for (const row of bridge) {
    const target = enabledIds.has(text(row.restaurantId)) ? enabled : excluded;
    target.factCount += 1;
    target.salesNok += number(row.salesAmount);
    target.quantity += number(row.quantity);
  }
  return {
    enabled,
    excluded,
    complete: {
      factCount: enabled.factCount + excluded.factCount,
      salesNok: enabled.salesNok + excluded.salesNok,
      quantity: enabled.quantity + excluded.quantity,
    },
  };
}

function addToAggregate(aggregates, key, row) {
  const aggregate = aggregates.get(key) ?? { denominator: 0, mappedByGroup: new Map() };
  const sales = number(row.salesAmount);
  aggregate.denominator += sales;
  if (text(row.resolutionStatus) === "Mapped") {
    const groupId = text(row.effectiveReportingGroupId);
    aggregate.mappedByGroup.set(groupId, (aggregate.mappedByGroup.get(groupId) ?? 0) + sales);
  }
  aggregates.set(key, aggregate);
}

function metricResult({
  metricId,
  importId,
  reportingGroupId,
  scopeType,
  restaurantId,
  restaurantScopeFingerprint,
  aggregate,
  mappingAsOfDate,
  mappingFingerprint,
  calculatedAt,
}) {
  const denominatorSalesNok = aggregate?.denominator ?? 0;
  const numeratorSalesNok = aggregate?.mappedByGroup.get(reportingGroupId) ?? 0;
  return {
    metricResultId: deterministicMetricResultId({
      metricId,
      importId,
      reportingGroupId,
      scopeType,
      restaurantId,
      restaurantScopeFingerprint,
    }),
    metricId,
    importId,
    reportingGroupId,
    scopeType,
    restaurantId,
    restaurantScopeFingerprint,
    channelScope: ALL_CHANNELS_SCOPE,
    publicationState: ACTIVE_FINALIZED,
    numeratorSalesNok,
    denominatorSalesNok,
    metricValue: denominatorSalesNok === 0 ? 0 : numeratorSalesNok / denominatorSalesNok,
    mappingAsOfDate,
    mappingFingerprint,
    calculatedAt,
  };
}

function uniqueById(rows, label, errors) {
  const byId = new Map();
  for (const row of rows) {
    const id = text(row.salesFactId);
    if (!id) {
      errors.push(`${label} contains a blank SalesFactID.`);
      continue;
    }
    if (byId.has(id)) errors.push(`${label} contains duplicate SalesFactID ${id}.`);
    byId.set(id, row);
  }
  return byId;
}

function uniqueAndSort(rows, idField, compare) {
  const seen = new Set();
  const output = [];
  for (const row of rows) {
    const id = text(row[idField]);
    if (!id) throw new Error(`Phase 2B input contains a blank ${idField}.`);
    if (seen.has(id)) throw new Error(`Phase 2B input contains duplicate ${idField} ${id}.`);
    seen.add(id);
    output.push({ ...row });
  }
  output.sort(compare);
  return output;
}

function compareImports(left, right) {
  return number(left.periodStart) - number(right.periodStart) ||
    number(left.periodEnd) - number(right.periodEnd) ||
    text(left.importId).localeCompare(text(right.importId));
}

function compareGroups(left, right) {
  return number(left.sortOrder) - number(right.sortOrder) ||
    text(left.reportingGroupName).localeCompare(text(right.reportingGroupName)) ||
    text(left.reportingGroupId).localeCompare(text(right.reportingGroupId));
}

function compareRestaurants(left, right) {
  return text(left.restaurantName).localeCompare(text(right.restaurantName)) ||
    text(left.restaurantId).localeCompare(text(right.restaurantId));
}

function scopeKey(importId, scopeType, restaurantId) {
  return [importId, scopeType, restaurantId].map(text).join("|");
}

function resultKey(row) {
  return [
    row.metricId,
    row.importId,
    row.reportingGroupId,
    row.scopeType,
    row.restaurantId,
    row.restaurantScopeFingerprint,
    row.channelScope,
    row.publicationState,
  ].map(text).join("|");
}

function factSignature(row) {
  return [
    row.salesFactId,
    row.importId,
    row.restaurantId,
    row.productId,
    row.legacyReportingCategoryId,
    row.reportingChannel,
    row.periodStart,
    row.periodEnd,
    row.publicationState,
    number(row.quantity),
    number(row.salesAmount),
  ].map(text).join("|");
}

function totals(rows) {
  return rows.reduce((result, row) => ({
    factCount: result.factCount + 1,
    salesNok: result.salesNok + number(row.salesAmount),
    quantity: result.quantity + number(row.quantity),
  }), { factCount: 0, salesNok: 0, quantity: 0 });
}

function emptyCoverage() {
  return {
    Mapped: { factCount: 0, salesNok: 0, quantity: 0 },
    Unmapped: { factCount: 0, salesNok: 0, quantity: 0 },
    Conflict: { factCount: 0, salesNok: 0, quantity: 0 },
    "Inactive Target": { factCount: 0, salesNok: 0, quantity: 0 },
  };
}

function emptyTotals() {
  return { factCount: 0, salesNok: 0, quantity: 0 };
}

function hashText(value) {
  let left = 0;
  let right = 0;
  const input = `${value}\n`;
  for (let index = 0; index < input.length; index += 1) {
    const code = input.charCodeAt(index);
    left = (left * 131 + code) % 2147483647;
    right = (right * 137 + code) % 2147483629;
  }
  return left.toString(16).padStart(8, "0") + right.toString(16).padStart(8, "0");
}

function almostEqual(left, right) {
  return Math.abs(number(left) - number(right)) <= 0.000001;
}

function text(value) {
  return String(value ?? "").trim();
}

function number(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}
