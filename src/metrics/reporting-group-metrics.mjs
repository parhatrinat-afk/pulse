/**
 * Deterministic Build 0.3.0 Phase 2A Reporting Group metric primitives.
 *
 * These functions operate on derived analysis rows. They never mutate source
 * facts or mapping rules. Build 0.3.0 intentionally applies the current
 * effective mapping state to historical facts for analysis.
 */

export const MAPPING_STATES = Object.freeze([
  "Mapped",
  "Unmapped",
  "Conflict",
  "Inactive Target",
]);

const MAPPING_STATE_SET = new Set(MAPPING_STATES);
const FINGERPRINT_VERSION = "PULSE-MAPPING-SEMANTIC-V2";

export function computeMappingFingerprint({
  asOfDate,
  groups,
  rules,
  products,
  resolutions,
}) {
  const records = [record("V", [FINGERPRINT_VERSION, asOfDate])];

  for (const group of groups) {
    records.push(record("G", [
      group.reportingGroupId ?? group.id,
      group.reportingGroupName ?? group.name,
      group.active,
      group.sortOrder,
    ]));
  }

  for (const rule of rules) {
    records.push(record("R", [
      rule.mappingRuleId ?? rule.id,
      rule.sourceSystemId,
      rule.scopeType,
      rule.nodeId,
      normalizeRuleAction(rule),
      rule.targetReportingGroupId ?? rule.targetGroupId,
      rule.effectiveFrom,
      rule.effectiveTo,
      rule.status,
    ]));
  }

  for (const product of products) {
    records.push(record("P", [
      product.productId,
      product.sourceSystemId,
      product.mainNodeId,
      product.subNodeId,
    ]));
  }

  for (const resolution of resolutions) {
    records.push(record("E", resolutionSignature(resolution)));
  }

  records.sort();
  return hashRecords(records);
}

function normalizeRuleAction(rule) {
  return String(rule.ruleAction ?? rule.action ?? "").trim() || "Map";
}

export function validateEffectiveMappingFreshness({
  expected,
  materialized,
  expectedAsOfDate,
}) {
  const errors = [];
  const expectedByProduct = uniqueByProduct(expected, "expected", errors);
  const actualByProduct = uniqueByProduct(materialized, "materialized", errors);

  if (expected.length !== materialized.length) {
    errors.push(
      `Effective Mapping row count is ${materialized.length}; expected ${expected.length}.`,
    );
  }

  for (const row of materialized) {
    if (Number(row.asOfDate) !== Number(expectedAsOfDate)) {
      errors.push(
        `Effective Mapping ${row.productId || "(blank ProductID)"} has AsOfDate ${row.asOfDate}; expected ${expectedAsOfDate}.`,
      );
    }
  }

  for (const [productId, expectedRow] of expectedByProduct) {
    const actualRow = actualByProduct.get(productId);
    if (!actualRow) {
      errors.push(`Effective Mapping is missing product ${productId}.`);
      continue;
    }
    const expectedSignature = record("E", resolutionSignature(expectedRow));
    const actualSignature = record("E", resolutionSignature(actualRow));
    if (actualSignature !== expectedSignature) {
      errors.push(`Effective Mapping is stale for product ${productId}.`);
    }
  }

  return errors;
}

export function buildMetricBridge({
  facts,
  effectiveMappings,
  mappingAsOfDate,
  mappingFingerprint,
  metricRefreshAt,
}) {
  const mappingByProduct = new Map();
  for (const mapping of effectiveMappings) {
    const productId = String(mapping.productId ?? "").trim();
    if (!productId) throw new Error("Effective Mapping contains a blank ProductID.");
    if (mappingByProduct.has(productId)) {
      throw new Error(`Effective Mapping contains duplicate ProductID ${productId}.`);
    }
    if (!MAPPING_STATE_SET.has(mapping.resolutionStatus)) {
      throw new Error(
        `Effective Mapping product ${productId} has unsupported ResolutionStatus ${mapping.resolutionStatus}.`,
      );
    }
    mappingByProduct.set(productId, mapping);
  }

  const bridge = [];
  const factIds = new Set();
  for (const fact of facts) {
    const salesFactId = String(fact.salesFactId ?? "").trim();
    const productId = String(fact.productId ?? "").trim();
    if (!salesFactId) throw new Error("Source facts contain a blank SalesFactID.");
    if (factIds.has(salesFactId)) {
      throw new Error(`Source facts contain duplicate SalesFactID ${salesFactId}.`);
    }
    factIds.add(salesFactId);
    if (!productId) throw new Error(`Sales fact ${salesFactId} has a blank ProductID.`);
    const mapping = mappingByProduct.get(productId);
    if (!mapping) {
      throw new Error(`Sales fact ${salesFactId} references product ${productId} without Effective Mapping.`);
    }

    bridge.push({
      salesFactId,
      importId: fact.importId,
      restaurantId: fact.restaurantId,
      productId,
      reportingChannel: fact.reportingChannel,
      periodStart: fact.periodStart,
      periodEnd: fact.periodEnd,
      publicationState: fact.publicationState,
      quantity: number(fact.quantity),
      salesAmount: number(fact.salesAmount),
      legacyReportingCategoryId: fact.legacyReportingCategoryId,
      effectiveReportingGroupId: mapping.effectiveReportingGroupId ?? "",
      effectiveReportingGroupName: mapping.effectiveReportingGroupName ?? "",
      resolutionSource: mapping.resolutionSource,
      resolutionState: mapping.resolutionState,
      resolutionStatus: mapping.resolutionStatus,
      winningRuleId: mapping.winningRuleId ?? "",
      mappingAsOfDate,
      mappingFingerprint,
      metricRefreshAt,
    });
  }
  return bridge;
}

export function reconcileFactsAndBridge({ facts, bridge, scopes }) {
  return scopes.map(scope => {
    const sourceRows = facts.filter(fact => inScope(fact, scope));
    const bridgeRows = bridge.filter(row => inScope(row, scope));
    const source = totals(sourceRows);
    const derived = totals(bridgeRows);
    const coverage = stateCoverage(bridgeRows);
    const stateTotals = MAPPING_STATES.reduce(
      (result, state) => addTotals(result, coverage[state]),
      emptyTotals(),
    );
    const result =
      source.factCount === derived.factCount &&
      source.factCount === stateTotals.factCount &&
      almostEqual(source.salesNok, derived.salesNok) &&
      almostEqual(source.salesNok, stateTotals.salesNok) &&
      almostEqual(source.quantity, derived.quantity) &&
      almostEqual(source.quantity, stateTotals.quantity)
        ? "PASS"
        : "FAIL";

    return {
      ...scope,
      source,
      derived,
      coverage,
      stateTotals,
      factCountVariance: derived.factCount - source.factCount,
      salesVariance: derived.salesNok - source.salesNok,
      quantityVariance: derived.quantity - source.quantity,
      result,
    };
  });
}

export function aggregateReportingGroups(bridge, scope) {
  const byGroup = new Map();
  for (const row of bridge) {
    if (!inScope(row, scope) || row.resolutionStatus !== "Mapped") continue;
    const groupId = row.effectiveReportingGroupId;
    const current = byGroup.get(groupId) ?? {
      reportingGroupId: groupId,
      reportingGroupName: row.effectiveReportingGroupName,
      factCount: 0,
      salesNok: 0,
      quantity: 0,
    };
    current.factCount += 1;
    current.salesNok += number(row.salesAmount);
    current.quantity += number(row.quantity);
    byGroup.set(groupId, current);
  }
  const rows = [];
  byGroup.forEach(value => rows.push(value));
  rows.sort((left, right) => left.reportingGroupId.localeCompare(right.reportingGroupId));
  return rows;
}

export function buildLegacyRpgCrosswalk(bridge, scopes) {
  const output = [];
  for (const scope of scopes) {
    const aggregates = new Map();
    for (const row of bridge) {
      if (!inScope(row, scope)) continue;
      const groupId = row.effectiveReportingGroupId || "";
      const key = [
        row.legacyReportingCategoryId || "",
        groupId,
        row.resolutionStatus,
      ].join("\u001f");
      const current = aggregates.get(key) ?? {
        legacyReportingCategoryId: row.legacyReportingCategoryId || "",
        effectiveReportingGroupId: groupId,
        effectiveReportingGroupName: row.effectiveReportingGroupName || "",
        resolutionStatus: row.resolutionStatus,
        factCount: 0,
        salesNok: 0,
        quantity: 0,
      };
      current.factCount += 1;
      current.salesNok += number(row.salesAmount);
      current.quantity += number(row.quantity);
      aggregates.set(key, current);
    }
    aggregates.forEach(value => output.push({ ...scope, ...value }));
  }
  output.sort((left, right) =>
    scopeKey(left).localeCompare(scopeKey(right)) ||
    left.legacyReportingCategoryId.localeCompare(right.legacyReportingCategoryId) ||
    left.effectiveReportingGroupId.localeCompare(right.effectiveReportingGroupId) ||
    left.resolutionStatus.localeCompare(right.resolutionStatus),
  );
  return output;
}

export function validateEquivalenceDefinitions(definitions, { groupIds, categoryIds }) {
  const errors = [];
  const allowedStatuses = new Set(["Equivalent", "Partial", "Not Comparable"]);
  const allowedActive = new Set(["Yes", "No"]);
  const knownGroups = new Set(groupIds);
  const knownCategories = new Set(categoryIds);
  const byDefinition = new Map();

  for (const definition of definitions) {
    const definitionId = String(definition.definitionId ?? "").trim();
    if (!definitionId) {
      errors.push("Legacy equivalence row has a blank DefinitionID.");
      continue;
    }
    if (!knownGroups.has(definition.reportingGroupId)) {
      errors.push(`${definitionId} references unknown ReportingGroupID ${definition.reportingGroupId}.`);
    }
    if (!knownCategories.has(definition.legacyReportingCategoryId)) {
      errors.push(`${definitionId} references unknown legacy category ${definition.legacyReportingCategoryId}.`);
    }
    if (!allowedStatuses.has(definition.comparisonStatus)) {
      errors.push(`${definitionId} has unsupported ComparisonStatus ${definition.comparisonStatus}.`);
    }
    if (!allowedActive.has(definition.active)) {
      errors.push(`${definitionId} has unsupported Active value ${definition.active}.`);
    }

    const state = byDefinition.get(definitionId) ?? {
      reportingGroupId: definition.reportingGroupId,
      comparisonStatus: definition.comparisonStatus,
      categories: new Set(),
    };
    if (state.reportingGroupId !== definition.reportingGroupId) {
      errors.push(`${definitionId} targets more than one ReportingGroupID.`);
    }
    if (state.comparisonStatus !== definition.comparisonStatus) {
      errors.push(`${definitionId} has inconsistent ComparisonStatus values.`);
    }
    if (state.categories.has(definition.legacyReportingCategoryId)) {
      errors.push(`${definitionId} repeats legacy category ${definition.legacyReportingCategoryId}.`);
    }
    state.categories.add(definition.legacyReportingCategoryId);
    byDefinition.set(definitionId, state);
  }

  return errors;
}

export function compareLegacyDefinitions({ bridge, definitions, scopes }) {
  const activeDefinitions = new Map();
  for (const row of definitions) {
    if (row.active !== "Yes") continue;
    const current = activeDefinitions.get(row.definitionId) ?? {
      definitionId: row.definitionId,
      reportingGroupId: row.reportingGroupId,
      comparisonStatus: row.comparisonStatus,
      legacyCategoryIds: [],
    };
    current.legacyCategoryIds.push(row.legacyReportingCategoryId);
    activeDefinitions.set(row.definitionId, current);
  }

  const comparisons = [];
  activeDefinitions.forEach(definition => {
    definition.legacyCategoryIds.sort();
    const legacyIds = new Set(definition.legacyCategoryIds);
    for (const scope of scopes) {
      const scopeRows = bridge.filter(row => inScope(row, scope));
      const legacyRows = scopeRows.filter(row => legacyIds.has(row.legacyReportingCategoryId));
      const rpgRows = scopeRows.filter(row =>
        row.resolutionStatus === "Mapped" &&
        row.effectiveReportingGroupId === definition.reportingGroupId,
      );
      const denominator = totals(scopeRows);
      const legacy = totals(legacyRows);
      const rpg = totals(rpgRows);
      const factCountVariance = rpg.factCount - legacy.factCount;
      const salesVariance = rpg.salesNok - legacy.salesNok;
      const quantityVariance = rpg.quantity - legacy.quantity;
      const equivalent =
        factCountVariance === 0 &&
        almostEqual(salesVariance, 0) &&
        almostEqual(quantityVariance, 0);
      const result = definition.comparisonStatus === "Equivalent"
        ? (equivalent ? "PASS" : "VARIANCE")
        : "INFO";
      comparisons.push({
        ...scope,
        definitionId: definition.definitionId,
        comparisonStatus: definition.comparisonStatus,
        reportingGroupId: definition.reportingGroupId,
        legacyCategoryIds: definition.legacyCategoryIds.join(", "),
        scopeFactCount: denominator.factCount,
        scopeSalesNok: denominator.salesNok,
        scopeQuantity: denominator.quantity,
        legacyFactCount: legacy.factCount,
        rpgFactCount: rpg.factCount,
        factCountVariance,
        legacySalesNok: legacy.salesNok,
        rpgSalesNok: rpg.salesNok,
        salesVariance,
        legacyQuantity: legacy.quantity,
        rpgQuantity: rpg.quantity,
        quantityVariance,
        legacySalesShare: denominator.salesNok ? legacy.salesNok / denominator.salesNok : 0,
        rpgSalesShare: denominator.salesNok ? rpg.salesNok / denominator.salesNok : 0,
        salesShareVariance: denominator.salesNok ? salesVariance / denominator.salesNok : 0,
        result,
      });
    }
  });
  comparisons.sort((left, right) =>
    left.definitionId.localeCompare(right.definitionId) ||
    scopeKey(left).localeCompare(scopeKey(right)),
  );
  return comparisons;
}

export function stateCoverage(rows) {
  const result = {};
  for (const state of MAPPING_STATES) result[state] = emptyTotals();
  for (const row of rows) {
    if (!MAPPING_STATE_SET.has(row.resolutionStatus)) {
      throw new Error(`Unsupported mapping state ${row.resolutionStatus}.`);
    }
    const bucket = result[row.resolutionStatus];
    bucket.factCount += 1;
    bucket.salesNok += number(row.salesAmount);
    bucket.quantity += number(row.quantity);
  }
  return result;
}

export function totals(rows) {
  const result = emptyTotals();
  for (const row of rows) {
    result.factCount += 1;
    result.salesNok += number(row.salesAmount);
    result.quantity += number(row.quantity);
  }
  return result;
}

export function almostEqual(left, right) {
  return Math.abs(number(left) - number(right)) <=
    Math.max(0.000001, Math.abs(number(right)) * 1e-12);
}

function uniqueByProduct(rows, label, errors) {
  const result = new Map();
  for (const row of rows) {
    const productId = String(row.productId ?? "").trim();
    if (!productId) {
      errors.push(`${label} Effective Mapping contains a blank ProductID.`);
      continue;
    }
    if (result.has(productId)) {
      errors.push(`${label} Effective Mapping contains duplicate ProductID ${productId}.`);
      continue;
    }
    result.set(productId, row);
  }
  return result;
}

function resolutionSignature(row) {
  return [
    row.productId,
    normalizeDelimited(row.mainRuleIds ?? row.mainCategoryRuleId),
    normalizeDelimited(row.mainTargetIds ?? row.mainCategoryTargetId),
    normalizeDelimited(row.subRuleIds ?? row.subcategoryRuleId),
    normalizeDelimited(row.subTargetIds ?? row.subcategoryTargetId),
    normalizeDelimited(row.productRuleIds ?? row.productRuleId),
    normalizeDelimited(row.productTargetIds ?? row.productTargetId),
    row.effectiveReportingGroupId ?? row.effectiveGroupId,
    row.effectiveReportingGroupName ?? row.effectiveGroupName,
    row.resolutionSource,
    row.resolutionState,
    row.resolutionStatus,
    normalizeDelimited(row.winningRuleId),
    row.mainNodeId,
    row.subNodeId ?? row.subcategoryNodeId,
  ];
}

function normalizeDelimited(value) {
  const items = String(value ?? "")
    .split(",")
    .map(item => item.trim())
    .filter(Boolean)
    .sort();
  return items.join(", ");
}

function record(kind, values) {
  return `${kind}|${values.map(value => {
    const normalized = normalize(value);
    return `${normalized.length}:${normalized}`;
  }).join("|")}`;
}

function normalize(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") return Number.isFinite(value) ? String(value) : "";
  return String(value).trim();
}

function hashRecords(records) {
  let left = 0;
  let right = 0;
  for (const item of records) {
    const value = `${item}\n`;
    for (let index = 0; index < value.length; index++) {
      const code = value.charCodeAt(index);
      left = (left * 131 + code) % 2147483647;
      right = (right * 137 + code) % 2147483629;
    }
  }
  return `MAP-${left.toString(16).padStart(8, "0")}${right.toString(16).padStart(8, "0")}`;
}

function inScope(row, scope) {
  if (scope.publicationState && row.publicationState !== scope.publicationState) return false;
  if (scope.importId && scope.importId !== "All imports" && row.importId !== scope.importId) return false;
  if (scope.channel && scope.channel !== "All channels" && row.reportingChannel !== scope.channel) return false;
  if (scope.restaurantId && row.restaurantId !== scope.restaurantId) return false;
  return true;
}

function scopeKey(scope) {
  return [scope.publicationState ?? "", scope.importId ?? "", scope.channel ?? "", scope.restaurantId ?? ""].join("\u001f");
}

function number(value) {
  const converted = Number(value);
  return Number.isFinite(converted) ? converted : 0;
}

function emptyTotals() {
  return { factCount: 0, salesNok: 0, quantity: 0 };
}

function addTotals(left, right) {
  return {
    factCount: left.factCount + right.factCount,
    salesNok: left.salesNok + right.salesNok,
    quantity: left.quantity + right.quantity,
  };
}
