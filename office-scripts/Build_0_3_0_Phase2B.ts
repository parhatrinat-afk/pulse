/**
 * Pulse Build 0.3.0 Phase 2B — centralized Reporting Group Sales Share.
 *
 * Prerequisite sequence:
 * 1. Run the accepted Phase 1 mapping refresh.
 * 2. Run the accepted Phase 2A bridge refresh on the same day.
 * 3. Run this script.
 *
 * Phase 2B treats tblMetricRPGFacts as the authoritative mapping-aware metric
 * bridge. It never resolves Mapping Rules independently and never rewrites
 * source facts, mapping configuration, Effective Mapping, or legacy CAT data.
 */
function main(workbook: ExcelScript.Workbook): string {
  const factsTable = requiredTable(workbook, "tblSalesFacts");
  const bridgeTable = requiredTable(workbook, "tblMetricRPGFacts");
  const importsTable = requiredTable(workbook, "tblImports");
  const groupsTable = requiredTable(workbook, "tblReportingGroups");
  const rulesTable = requiredTable(workbook, "tblMappingRules");
  const productsTable = requiredTable(workbook, "tblProducts");
  const classificationsTable = requiredTable(workbook, "tblSourceClassifications");
  const effectiveTable = requiredTable(workbook, "tblEffectiveMapping");
  const restaurantsTable = requiredTable(workbook, "tblRestaurants");
  const kpiTable = requiredTable(workbook, "tblKPIRegistry");
  const legacyCategoriesTable = requiredTable(workbook, "tblReportingCategories");
  const equivalenceTable = requiredTable(workbook, "tblLegacyRPGEquivalence");
  const environmentTable = requiredTable(workbook, "tblEnvironment");
  const buildLogTable = requiredTable(workbook, "tblBuildLog");

  const calc = requiredSheet(workbook, "_Metric_Calc");
  const performance = requiredSheet(workbook, "Performance");
  const reports = requiredSheet(workbook, "Reports");
  const overview = requiredSheet(workbook, "Overview");

  // Read and validate the complete prerequisite state before any workbook write.
  const asOfDate = excelToday();
  const groups = readGroups(groupsTable);
  const rules = readRules(rulesTable);
  const products = buildHierarchy(classificationsTable, productsTable);
  const effectiveMapping = readEffectiveMapping(effectiveTable);
  const effectiveErrors = validateMaterializedEffectiveMapping(products, effectiveMapping, asOfDate);
  if (effectiveErrors.length) {
    throw new Error(
      `PUL-0302B-001: Phase 2A metric bridge is stale or invalid. Run Phase 1 and Phase 2A, then rerun Phase 2B. ` +
      effectiveErrors.slice(0, 10).join(" | ")
    );
  }

  const currentMappingFingerprint = computeMappingFingerprint(
    asOfDate,
    groups,
    rules,
    products,
    effectiveMapping
  );
  const reorderedFingerprint = computeMappingFingerprint(
    asOfDate,
    groups.slice().reverse(),
    rules.slice().reverse(),
    products.slice().reverse(),
    effectiveMapping.slice().reverse()
  );
  if (currentMappingFingerprint !== reorderedFingerprint) {
    throw new Error("PUL-0302B-002: Current mapping fingerprint is not reproducible across row ordering.");
  }

  const facts = readFacts(factsTable);
  const bridge = readMetricBridge(bridgeTable);
  const bridgeErrors = validateMetricBridgeForCutover(
    facts,
    bridge,
    asOfDate,
    currentMappingFingerprint
  );
  if (bridgeErrors.length) {
    throw new Error(
      `PUL-0302B-001: Phase 2A metric bridge is stale or invalid. Run Phase 1 and Phase 2A, then rerun Phase 2B. ` +
      bridgeErrors.slice(0, 10).join(" | ")
    );
  }

  const activeImports = readActiveImports(importsTable);
  const activeGroups = groups.filter(group => group.active === "Yes");
  activeGroups.sort((left, right) =>
    left.sortOrder - right.sortOrder || left.name.localeCompare(right.name) || left.id.localeCompare(right.id)
  );
  const reportingRestaurants = readReportingRestaurants(restaurantsTable);
  if (!activeImports.length) throw new Error("PUL-0302B-003: No Active Finalized dataset is available.");
  if (!activeGroups.length) throw new Error("PUL-0302B-004: No active Reporting Group is available.");

  const calculatedAt = excelNow();
  const results = buildCentralResults(
    bridge,
    activeImports,
    activeGroups,
    reportingRestaurants,
    asOfDate,
    currentMappingFingerprint,
    calculatedAt
  );
  const resultErrors = validateCentralResults(
    bridge,
    results,
    activeImports,
    activeGroups,
    reportingRestaurants
  );
  if (resultErrors.length) {
    throw new Error(`PUL-0302B-005: Centralized metric result validation failed. ${resultErrors.slice(0, 10).join(" | ")}`);
  }

  const expectedResultRows = activeImports.length * activeGroups.length * (1 + reportingRestaurants.length);
  if (results.length !== expectedResultRows) {
    throw new Error(`PUL-0302B-006: Central result row count ${results.length}; expected ${expectedResultRows}.`);
  }

  const priorSelection = capturePriorSelections(calc, performance, activeImports);
  const protectedBefore = snapshotProtectedTables(
    factsTable,
    bridgeTable,
    rulesTable,
    groupsTable,
    effectiveTable,
    restaurantsTable,
    productsTable,
    classificationsTable,
    legacyCategoriesTable,
    equivalenceTable
  );

  // All prerequisite validation above must succeed before these writes begin.
  workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.automatic);
  writeMetricCalc(calc, workbook.getTable("tblMetricRPGResults"), activeImports, activeGroups, reportingRestaurants, results);
  updateKpiRegistry(kpiTable);
  const validationMessage = writePerformance(
    performance,
    activeImports,
    activeGroups,
    reportingRestaurants,
    priorSelection
  );
  writeReports(reports);
  updateOverview(overview, activeImports);

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const protectedAfter = snapshotProtectedTables(
    factsTable,
    bridgeTable,
    rulesTable,
    groupsTable,
    effectiveTable,
    restaurantsTable,
    productsTable,
    classificationsTable,
    legacyCategoriesTable,
    equivalenceTable
  );
  if (protectedBefore !== protectedAfter) {
    throw new Error("PUL-0302B-007: A protected source, mapping, bridge, restaurant-scope, equivalence, or legacy CAT table changed during Phase 2B.");
  }

  writeMetricResultsQA(
    workbook,
    asOfDate,
    currentMappingFingerprint,
    facts,
    bridge,
    activeImports,
    activeGroups,
    reportingRestaurants,
    results,
    validationMessage
  );

  updateEnvironment(
    environmentTable,
    "BuildID",
    "0.3.0-Phase2B",
    "Central Reporting Group metric result and minimal Performance/Reports cutover."
  );
  updateEnvironment(
    environmentTable,
    "BuildVersion",
    "0.3.0-Phase2B",
    "KPI-0001 now uses centralized Reporting Group Sales Share results."
  );
  appendBuildLog(buildLogTable, [
    nextId(buildLogTable, "LOG", 6),
    calculatedAt,
    "0.3.0-Phase2B",
    "Activate Reporting Group Sales Share",
    "Success",
    "tblMetricRPGResults + Performance + Reports",
    `${results.length} centralized result rows; ${activeImports.length} datasets; ${activeGroups.length} active RPGs; ` +
      `${1 + reportingRestaurants.length} scopes; mapping ${currentMappingFingerprint}. ${validationMessage}`
  ]);

  return (
    `Pulse 0.3.0 Phase 2B applied. ${results.length} centralized KPI-0001 result rows, ` +
    `${activeImports.length} datasets, ${activeGroups.length} active Reporting Groups, ` +
    `${1 + reportingRestaurants.length} scopes. ${validationMessage}`
  );
}

type ReportingGroup = { id: string; name: string; active: string; sortOrder: number };
type MappingRule = {
  id: string; sourceSystemId: string; scopeType: string; nodeId: string;
  targetGroupId: string; ruleAction: string; effectiveFrom: number; effectiveTo: number; status: string;
};
type ProductNode = {
  productId: string; sourceSystemId: string; productName: string; salesAccount: string;
  sourceClassificationId: string; mainCategory: string; subCategory: string;
  mainNodeId: string; subNodeId: string;
};
type Resolution = ProductNode & {
  mainRuleIds: string; mainTargetIds: string; subRuleIds: string; subTargetIds: string;
  productRuleIds: string; productTargetIds: string; effectiveGroupId: string;
  effectiveGroupName: string; resolutionSource: string; resolutionState: string;
  resolutionStatus: string; winningRuleId: string; asOfDate: number;
};
type SourceFact = {
  salesFactId: string; importId: string; restaurantId: string; productId: string;
  reportingChannel: string; periodStart: number; periodEnd: number; publicationState: string;
  quantity: number; salesAmount: number; legacyReportingCategoryId: string;
};
type MetricFact = SourceFact & {
  effectiveGroupId: string; effectiveGroupName: string; resolutionSource: string;
  resolutionState: string; resolutionStatus: string; winningRuleId: string;
  mappingAsOfDate: number; mappingFingerprint: string; metricRefreshAt: number;
};
type ActiveImport = {
  id: string; label: string; start: number; end: number; days: number; year: number; week: string;
};
type ReportingRestaurant = { id: string; name: string };
type ScopeAggregate = { denominator: number; mappedByGroup: Map<string, number> };
type MetricResult = {
  id: string; metricId: string; importId: string; groupId: string; scopeType: string;
  restaurantId: string; restaurantScopeFingerprint: string; channelScope: string; publicationState: string;
  numeratorSales: number; denominatorSales: number; metricValue: number;
  mappingAsOfDate: number; mappingFingerprint: string; calculatedAt: number;
};
type Totals = { factCount: number; sales: number; quantity: number };
type PerformanceScopeCoverage = { enabled: Totals; excluded: Totals; complete: Totals };
type HashState = { left: number; right: number };
type SelectionState = { groupId: string; currentImportId: string; comparisonImportId: string };

const METRIC_ID = "KPI-0001";
const METRIC_NAME = "Reporting Group Sales Share";
const ALL_CHANNELS = "All channels";
const ACTIVE_FINALIZED = "Active Finalized";

function readGroups(table: ExcelScript.Table): ReportingGroup[] {
  const h = headerMap(table);
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  const seen = new Set<string>();
  const output: ReportingGroup[] = [];
  for (const row of values) {
    const id = text(row[h.ReportingGroupID]);
    if (!id) continue;
    if (seen.has(id)) throw new Error(`PUL-0302B-008: Duplicate ReportingGroupID ${id}.`);
    seen.add(id);
    output.push({
      id,
      name: text(row[h.ReportingGroupName]) || id,
      active: text(row[h.Active]),
      sortOrder: numberValue(row[h.SortOrder])
    });
  }
  return output;
}

function readRules(table: ExcelScript.Table): MappingRule[] {
  const h = headerMap(table);
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  const output: MappingRule[] = [];
  for (const row of values) {
    const id = text(row[h.MappingRuleID]);
    if (!id) continue;
    const ruleAction = h.RuleAction === undefined ? "Map" : text(row[h.RuleAction]) || "Map";
    if (ruleAction !== "Map" && ruleAction !== "Exclude") {
      throw new Error(`PUL-0302B-009: Rule ${id} has unsupported RuleAction ${ruleAction}.`);
    }
    output.push({
      id,
      sourceSystemId: text(row[h.SourceSystemID]),
      scopeType: text(row[h.ScopeType]),
      nodeId: text(row[h.NodeID]),
      targetGroupId: text(row[h.TargetReportingGroupID]),
      ruleAction,
      effectiveFrom: numberValue(row[h.EffectiveFrom]),
      effectiveTo: numberValue(row[h.EffectiveTo]),
      status: text(row[h.Status])
    });
  }
  return output;
}

function buildHierarchy(classifications: ExcelScript.Table, products: ExcelScript.Table): ProductNode[] {
  const classificationById = new Map<string, { sourceSystemId: string; main: string; sub: string }>();
  const ch = headerMap(classifications);
  const classificationValues = classifications.getRangeBetweenHeaderAndTotal().getValues();
  for (const row of classificationValues) {
    classificationById.set(text(row[ch.SourceClassificationID]), {
      sourceSystemId: text(row[ch.SourceSystemID]),
      main: text(row[ch.SourceMainCategory]),
      sub: text(row[ch.SourceSubCategory])
    });
  }
  const ph = headerMap(products);
  const productValues = products.getRangeBetweenHeaderAndTotal().getValues();
  const seen = new Set<string>();
  const output: ProductNode[] = [];
  for (const row of productValues) {
    const productId = text(row[ph.ProductID]);
    if (!productId) continue;
    if (seen.has(productId)) throw new Error(`PUL-0302B-010: Duplicate ProductID ${productId}.`);
    seen.add(productId);
    const sourceClassificationId = text(row[ph.SourceClassificationID]);
    const classification = classificationById.get(sourceClassificationId);
    if (!classification) {
      throw new Error(`PUL-0302B-011: Product ${productId} references missing classification ${sourceClassificationId}.`);
    }
    output.push({
      productId,
      sourceSystemId: text(row[ph.SourceSystemID]),
      productName: text(row[ph.SourceProductName]),
      salesAccount: text(row[ph.SalesAccount]),
      sourceClassificationId,
      mainCategory: classification.main,
      subCategory: classification.sub,
      mainNodeId: mainNodeId(classification.sourceSystemId, classification.main),
      subNodeId: sourceClassificationId
    });
  }
  return output;
}

function readEffectiveMapping(table: ExcelScript.Table): Resolution[] {
  const h = headerMap(table);
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  const output: Resolution[] = [];
  for (const row of values) {
    const productId = text(row[h.ProductID]);
    if (!productId) continue;
    output.push({
      productId,
      sourceSystemId: text(row[h.SourceSystemID]),
      productName: text(row[h.SourceProductName]),
      salesAccount: text(row[h.SalesAccount]),
      mainCategory: text(row[h.SourceMainCategory]),
      subCategory: text(row[h.SourceSubCategory]),
      sourceClassificationId: text(row[h.SourceClassificationID]),
      mainRuleIds: text(row[h.MainCategoryRuleID]),
      mainTargetIds: text(row[h.MainCategoryTargetID]),
      subRuleIds: text(row[h.SubcategoryRuleID]),
      subTargetIds: text(row[h.SubcategoryTargetID]),
      productRuleIds: text(row[h.ProductRuleID]),
      productTargetIds: text(row[h.ProductTargetID]),
      effectiveGroupId: text(row[h.EffectiveReportingGroupID]),
      effectiveGroupName: text(row[h.EffectiveReportingGroupName]),
      resolutionSource: text(row[h.ResolutionSource]),
      resolutionState: text(row[h.ResolutionState]),
      resolutionStatus: text(row[h.ResolutionStatus]),
      winningRuleId: text(row[h.WinningRuleID]),
      mainNodeId: text(row[h.MainNodeID]),
      subNodeId: text(row[h.SubcategoryNodeID]),
      asOfDate: numberValue(row[h.AsOfDate])
    });
  }
  return output;
}

function validateMaterializedEffectiveMapping(
  products: ProductNode[],
  resolutions: Resolution[],
  asOfDate: number
): string[] {
  const errors: string[] = [];
  if (products.length !== resolutions.length) {
    errors.push(`Effective Mapping row count ${resolutions.length}; expected ${products.length}.`);
  }
  const productById = new Map<string, ProductNode>();
  for (const product of products) productById.set(product.productId, product);
  const seen = new Set<string>();
  for (const row of resolutions) {
    if (seen.has(row.productId)) errors.push(`Effective Mapping contains duplicate ProductID ${row.productId}.`);
    seen.add(row.productId);
    const product = productById.get(row.productId);
    if (!product) {
      errors.push(`Effective Mapping contains unknown ProductID ${row.productId}.`);
      continue;
    }
    if (row.asOfDate !== asOfDate) errors.push(`${row.productId} AsOfDate ${row.asOfDate}; expected ${asOfDate}.`);
    if (!isMappingState(row.resolutionStatus)) errors.push(`${row.productId} unsupported state ${row.resolutionStatus}.`);
    if (row.mainNodeId !== product.mainNodeId || row.subNodeId !== product.subNodeId) {
      errors.push(`${row.productId} hierarchy identity differs from Products/Source Classifications.`);
    }
  }
  for (const product of products) {
    if (!seen.has(product.productId)) errors.push(`Effective Mapping is missing ProductID ${product.productId}.`);
  }
  return errors;
}

function computeMappingFingerprint(
  asOfDate: number,
  groups: ReportingGroup[],
  rules: MappingRule[],
  products: ProductNode[],
  resolutions: Resolution[]
): string {
  const records: string[] = [fingerprintRecord("V", ["PULSE-MAPPING-SEMANTIC-V2", asOfDate])];
  for (const group of groups) records.push(fingerprintRecord("G", [group.id, group.name, group.active, group.sortOrder]));
  for (const rule of rules) records.push(fingerprintRecord("R", [
    rule.id, rule.sourceSystemId, rule.scopeType, rule.nodeId, rule.ruleAction, rule.targetGroupId,
    rule.effectiveFrom, rule.effectiveTo, rule.status
  ]));
  for (const product of products) records.push(fingerprintRecord("P", [
    product.productId, product.sourceSystemId, product.mainNodeId, product.subNodeId
  ]));
  for (const resolution of resolutions) records.push(fingerprintRecord("E", resolutionSignature(resolution)));
  records.sort();
  return hashStrings(records, "MAP-");
}

function resolutionSignature(row: Resolution): (string | number | boolean)[] {
  return [
    row.productId,
    normalizeDelimited(row.mainRuleIds),
    normalizeDelimited(row.mainTargetIds),
    normalizeDelimited(row.subRuleIds),
    normalizeDelimited(row.subTargetIds),
    normalizeDelimited(row.productRuleIds),
    normalizeDelimited(row.productTargetIds),
    row.effectiveGroupId,
    row.effectiveGroupName,
    row.resolutionSource,
    row.resolutionState,
    row.resolutionStatus,
    normalizeDelimited(row.winningRuleId),
    row.mainNodeId,
    row.subNodeId
  ];
}

function normalizeDelimited(value: string): string {
  return value.split(",")
    .map(item => item.trim())
    .filter(item => item.length > 0)
    .sort()
    .join(", ");
}

function readFacts(table: ExcelScript.Table): SourceFact[] {
  const h = headerMap(table);
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  const output: SourceFact[] = [];
  for (const row of values) {
    const salesFactId = text(row[h.SalesFactID]);
    if (!salesFactId) throw new Error("PUL-0302B-012: tblSalesFacts contains a blank SalesFactID.");
    output.push({
      salesFactId,
      importId: text(row[h.ImportID]),
      restaurantId: text(row[h.RestaurantID]),
      productId: text(row[h.ProductID]),
      legacyReportingCategoryId: text(row[h.ReportingCategoryID]),
      periodStart: numberValue(row[h.PeriodStart]),
      periodEnd: numberValue(row[h.PeriodEnd]),
      reportingChannel: text(row[h.ReportingChannel]),
      quantity: numberValue(row[h.Quantity]),
      salesAmount: numberValue(row[h.SalesAmount]),
      publicationState: text(row[h.PublicationState])
    });
  }
  return output;
}

function readMetricBridge(table: ExcelScript.Table): MetricFact[] {
  const h = headerMap(table);
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  const output: MetricFact[] = [];
  for (const row of values) {
    const salesFactId = text(row[h.SalesFactID]);
    if (!salesFactId) continue;
    output.push({
      salesFactId,
      importId: text(row[h.ImportID]),
      restaurantId: text(row[h.RestaurantID]),
      productId: text(row[h.ProductID]),
      reportingChannel: text(row[h.ReportingChannel]),
      periodStart: numberValue(row[h.PeriodStart]),
      periodEnd: numberValue(row[h.PeriodEnd]),
      publicationState: text(row[h.PublicationState]),
      quantity: numberValue(row[h.Quantity]),
      salesAmount: numberValue(row[h.SalesAmount]),
      legacyReportingCategoryId: text(row[h.LegacyReportingCategoryID]),
      effectiveGroupId: text(row[h.EffectiveReportingGroupID]),
      effectiveGroupName: text(row[h.EffectiveReportingGroupName]),
      resolutionSource: text(row[h.ResolutionSource]),
      resolutionState: text(row[h.ResolutionState]),
      resolutionStatus: text(row[h.ResolutionStatus]),
      winningRuleId: text(row[h.WinningRuleID]),
      mappingAsOfDate: numberValue(row[h.MappingAsOfDate]),
      mappingFingerprint: text(row[h.MappingFingerprint]),
      metricRefreshAt: numberValue(row[h.MetricRefreshAt])
    });
  }
  return output;
}

function validateMetricBridgeForCutover(
  facts: SourceFact[],
  bridge: MetricFact[],
  expectedAsOfDate: number,
  expectedMappingFingerprint: string
): string[] {
  const errors: string[] = [];
  if (facts.length !== bridge.length) errors.push(`bridge row count ${bridge.length}; expected ${facts.length}.`);
  const factsById = new Map<string, SourceFact>();
  for (const fact of facts) {
    if (factsById.has(fact.salesFactId)) errors.push(`duplicate source SalesFactID ${fact.salesFactId}.`);
    factsById.set(fact.salesFactId, fact);
  }
  const bridgeIds = new Set<string>();
  const coverage = emptyCoverage();
  for (const row of bridge) {
    if (bridgeIds.has(row.salesFactId)) errors.push(`duplicate bridge SalesFactID ${row.salesFactId}.`);
    bridgeIds.add(row.salesFactId);
    const source = factsById.get(row.salesFactId);
    if (!source) errors.push(`unknown bridge SalesFactID ${row.salesFactId}.`);
    else if (factSignature(source) !== factSignature(row)) errors.push(`source lineage differs for ${row.salesFactId}.`);
    if (row.mappingAsOfDate !== expectedAsOfDate) {
      errors.push(`${row.salesFactId} MappingAsOfDate ${row.mappingAsOfDate}; expected ${expectedAsOfDate}.`);
    }
    if (row.mappingFingerprint !== expectedMappingFingerprint) {
      errors.push(`${row.salesFactId} mapping fingerprint ${row.mappingFingerprint}; expected ${expectedMappingFingerprint}.`);
    }
    if (!isMappingState(row.resolutionStatus)) {
      errors.push(`${row.salesFactId} unsupported state ${row.resolutionStatus}.`);
      continue;
    }
    if (row.resolutionStatus === "Mapped" && !row.effectiveGroupId) {
      errors.push(`${row.salesFactId} is Mapped without an EffectiveReportingGroupID.`);
    }
    addTotals(coverage[row.resolutionStatus], row.salesAmount, row.quantity);
  }
  for (const fact of facts) {
    if (!bridgeIds.has(fact.salesFactId)) errors.push(`bridge is missing SalesFactID ${fact.salesFactId}.`);
  }
  const sourceTotals = totals(facts);
  const bridgeTotals = totals(bridge);
  const stateTotals = sumCoverage(coverage);
  if (!sameTotals(sourceTotals, bridgeTotals)) errors.push("source facts and bridge do not reconcile.");
  if (!sameTotals(stateTotals, bridgeTotals)) errors.push("four-state coverage does not reconcile to bridge totals.");
  return errors;
}

function readActiveImports(table: ExcelScript.Table): ActiveImport[] {
  const h = headerMap(table);
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  const output: ActiveImport[] = [];
  const seen = new Set<string>();
  for (const row of values) {
    const id = text(row[h.ImportID]);
    if (!id || text(row[h.PublicationState]) !== ACTIVE_FINALIZED || text(row[h.ActiveVersion]) !== "Yes") continue;
    if (seen.has(id)) throw new Error(`PUL-0302B-013: Duplicate active ImportID ${id}.`);
    seen.add(id);
    const start = numberValue(row[h.PeriodStart]);
    const end = numberValue(row[h.PeriodEnd]);
    const year = numberValue(row[h.ReportingYear]);
    const week = text(row[h.ReportingWeek]);
    output.push({
      id,
      label: makePeriodLabel(start, end, year, week),
      start,
      end,
      days: Math.round(end - start + 1),
      year,
      week
    });
  }
  output.sort((left, right) => left.start - right.start || left.end - right.end || left.id.localeCompare(right.id));
  return output;
}

function readReportingRestaurants(table: ExcelScript.Table): ReportingRestaurant[] {
  const h = headerMap(table);
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  const output: ReportingRestaurant[] = [];
  const seen = new Set<string>();
  for (const row of values) {
    const id = text(row[h.RestaurantID]);
    if (!id || text(row[h.Status]) !== "Active" || text(row[h.ReportingEnabled]) !== "Yes") continue;
    if (seen.has(id)) throw new Error(`PUL-0302B-014: Duplicate reporting RestaurantID ${id}.`);
    seen.add(id);
    output.push({ id, name: text(row[h.DisplayName]) || text(row[h.SourceRestaurantName]) || id });
  }
  output.sort((left, right) => left.name.localeCompare(right.name) || left.id.localeCompare(right.id));
  return output;
}

function buildCentralResults(
  bridge: MetricFact[],
  activeImports: ActiveImport[],
  activeGroups: ReportingGroup[],
  restaurants: ReportingRestaurant[],
  mappingAsOfDate: number,
  mappingFingerprint: string,
  calculatedAt: number
): MetricResult[] {
  const importIds = new Set<string>();
  const restaurantIds = new Set<string>();
  for (const value of activeImports) importIds.add(value.id);
  for (const value of restaurants) restaurantIds.add(value.id);
  const companyRestaurantScopeFingerprint = deterministicRestaurantScopeFingerprint(restaurants);
  const aggregates = new Map<string, ScopeAggregate>();
  for (const row of bridge) {
    if (row.publicationState !== ACTIVE_FINALIZED || !importIds.has(row.importId)) continue;
    if (!restaurantIds.has(row.restaurantId)) continue;
    addToAggregate(aggregates, scopeKey(row.importId, "Company", ""), row);
    addToAggregate(aggregates, scopeKey(row.importId, "Restaurant", row.restaurantId), row);
  }
  const output: MetricResult[] = [];
  for (const activeImport of activeImports) {
    for (const group of activeGroups) {
      output.push(makeMetricResult(
        activeImport.id,
        group.id,
        "Company",
        "",
        companyRestaurantScopeFingerprint,
        aggregates.get(scopeKey(activeImport.id, "Company", "")),
        mappingAsOfDate,
        mappingFingerprint,
        calculatedAt
      ));
      for (const restaurant of restaurants) {
        output.push(makeMetricResult(
          activeImport.id,
          group.id,
          "Restaurant",
          restaurant.id,
          "",
          aggregates.get(scopeKey(activeImport.id, "Restaurant", restaurant.id)),
          mappingAsOfDate,
          mappingFingerprint,
          calculatedAt
        ));
      }
    }
  }
  return output;
}

function addToAggregate(aggregates: Map<string, ScopeAggregate>, key: string, row: MetricFact): void {
  const aggregate = aggregates.get(key) ?? { denominator: 0, mappedByGroup: new Map<string, number>() };
  aggregate.denominator += row.salesAmount;
  if (row.resolutionStatus === "Mapped") {
    aggregate.mappedByGroup.set(
      row.effectiveGroupId,
      (aggregate.mappedByGroup.get(row.effectiveGroupId) ?? 0) + row.salesAmount
    );
  }
  aggregates.set(key, aggregate);
}

function makeMetricResult(
  importId: string,
  groupId: string,
  scopeType: string,
  restaurantId: string,
  restaurantScopeFingerprint: string,
  aggregate: ScopeAggregate | undefined,
  mappingAsOfDate: number,
  mappingFingerprint: string,
  calculatedAt: number
): MetricResult {
  const denominatorSales = aggregate ? aggregate.denominator : 0;
  const numeratorSales = aggregate ? aggregate.mappedByGroup.get(groupId) ?? 0 : 0;
  return {
    id: deterministicMetricResultId(
      METRIC_ID,
      importId,
      groupId,
      scopeType,
      restaurantId,
      restaurantScopeFingerprint
    ),
    metricId: METRIC_ID,
    importId,
    groupId,
    scopeType,
    restaurantId,
    restaurantScopeFingerprint,
    channelScope: ALL_CHANNELS,
    publicationState: ACTIVE_FINALIZED,
    numeratorSales,
    denominatorSales,
    metricValue: denominatorSales === 0 ? 0 : numeratorSales / denominatorSales,
    mappingAsOfDate,
    mappingFingerprint,
    calculatedAt
  };
}

function validateCentralResults(
  bridge: MetricFact[],
  results: MetricResult[],
  activeImports: ActiveImport[],
  activeGroups: ReportingGroup[],
  restaurants: ReportingRestaurant[]
): string[] {
  const errors: string[] = [];
  const expectedCount = activeImports.length * activeGroups.length * (1 + restaurants.length);
  const restaurantIds = new Set<string>();
  for (const restaurant of restaurants) restaurantIds.add(restaurant.id);
  const companyRestaurantScopeFingerprint = deterministicRestaurantScopeFingerprint(restaurants);
  if (results.length !== expectedCount) errors.push(`row count ${results.length}; expected ${expectedCount}.`);
  const seenIds = new Set<string>();
  const seenKeys = new Set<string>();
  for (const result of results) {
    if (seenIds.has(result.id)) errors.push(`duplicate MetricResultID ${result.id}.`);
    seenIds.add(result.id);
    const key = metricResultKey(result);
    if (seenKeys.has(key)) errors.push(`duplicate metric grain ${key}.`);
    seenKeys.add(key);
    const expectedId = deterministicMetricResultId(
      result.metricId,
      result.importId,
      result.groupId,
      result.scopeType,
      result.restaurantId,
      result.restaurantScopeFingerprint
    );
    if (expectedId !== result.id) errors.push(`non-deterministic MetricResultID ${result.id}.`);
    const expectedScopeFingerprint = result.scopeType === "Company" ? companyRestaurantScopeFingerprint : "";
    if (result.restaurantScopeFingerprint !== expectedScopeFingerprint) {
      errors.push(`restaurant scope fingerprint differs for ${metricResultKey(result)}.`);
    }
  }

  // Independently rescan the bridge for each materialized result. The current
  // all-enabled checkpoint is only 306 results, so this path is bounded.
  for (const result of results) {
    let numerator = 0;
    let denominator = 0;
    for (const row of bridge) {
      if (row.publicationState !== ACTIVE_FINALIZED || row.importId !== result.importId) continue;
      if (!restaurantIds.has(row.restaurantId)) continue;
      if (result.scopeType === "Restaurant" && row.restaurantId !== result.restaurantId) continue;
      denominator += row.salesAmount;
      if (row.resolutionStatus === "Mapped" && row.effectiveGroupId === result.groupId) {
        numerator += row.salesAmount;
      }
    }
    const metricValue = denominator === 0 ? 0 : numerator / denominator;
    if (!almostEqual(numerator, result.numeratorSales) ||
        !almostEqual(denominator, result.denominatorSales) ||
        !almostEqual(metricValue, result.metricValue)) {
      errors.push(`independent result differs for ${metricResultKey(result)}.`);
    }
  }
  return errors;
}

function deterministicMetricResultId(
  metricId: string,
  importId: string,
  groupId: string,
  scopeType: string,
  restaurantId: string,
  restaurantScopeFingerprint: string
): string {
  const companyScopeFingerprint = scopeType === "Company" ? restaurantScopeFingerprint : "";
  const key = [
    metricId,
    importId,
    groupId,
    scopeType,
    restaurantId,
    companyScopeFingerprint,
    ALL_CHANNELS,
    ACTIVE_FINALIZED
  ].join("|");
  return hashStrings([key], "MRR-");
}

function deterministicRestaurantScopeFingerprint(restaurants: ReportingRestaurant[]): string {
  const ids = restaurants.map(value => value.id);
  ids.sort();
  const serialized = ids.map(id => `${id.length}:${id}`).join("|");
  return hashStrings([`ENABLED-RESTAURANTS|${serialized}`], "RSC-");
}

function capturePriorSelections(
  calc: ExcelScript.Worksheet,
  performance: ExcelScript.Worksheet,
  activeImports: ActiveImport[]
): SelectionState {
  let groupId = "";
  const helperValues = calc.getRangeByIndexes(0, 8, 1000, 2).getValues();
  const selectedGroupLabel = text(performance.getRange("B7").getValue());
  if (text(helperValues[0][0]) === "Reporting Group" && text(helperValues[0][1]) === "ReportingGroupID") {
    for (let index = 1; index < helperValues.length; index++) {
      if (text(helperValues[index][0]) === selectedGroupLabel) {
        groupId = text(helperValues[index][1]);
        break;
      }
    }
  }
  const currentLabel = text(performance.getRange("B10").getValue());
  const comparisonLabel = text(performance.getRange("G10").getValue());
  return {
    groupId,
    currentImportId: importIdByLabel(activeImports, currentLabel),
    comparisonImportId: importIdByLabel(activeImports, comparisonLabel)
  };
}

function writeMetricCalc(
  sheet: ExcelScript.Worksheet,
  existingResultTable: ExcelScript.Table | undefined,
  activeImports: ActiveImport[],
  activeGroups: ReportingGroup[],
  restaurants: ReportingRestaurant[],
  results: MetricResult[]
): void {
  if (existingResultTable) existingResultTable.delete();
  sheet.getRangeByIndexes(0, 0, 1000, 13).clear(ExcelScript.ClearApplyTo.all);
  sheet.getRangeByIndexes(0, 14, 1000, 15).clear(ExcelScript.ClearApplyTo.all);

  sheet.getRangeByIndexes(0, 0, 1, 5).setValues([[
    "Dataset", "ImportID", "PeriodStart", "PeriodEnd", "Days"
  ]]);
  sheet.getRangeByIndexes(1, 0, activeImports.length, 5).setValues(
    activeImports.map(value => [value.label, value.id, value.start, value.end, value.days])
  );

  sheet.getRangeByIndexes(0, 8, 1, 2).setValues([["Reporting Group", "ReportingGroupID"]]);
  sheet.getRangeByIndexes(1, 8, activeGroups.length, 2).setValues(
    activeGroups.map(value => [value.name, value.id])
  );

  sheet.getRangeByIndexes(0, 11, 1, 2).setValues([["Restaurant", "RestaurantID"]]);
  if (restaurants.length) {
    sheet.getRangeByIndexes(1, 11, restaurants.length, 2).setValues(
      restaurants.map(value => [value.name, value.id])
    );
  }

  const helperHeaders = sheet.getRangeByIndexes(0, 0, 1, 13);
  helperHeaders.getFormat().getFont().setBold(true);
  helperHeaders.getFormat().getFill().setColor("#172033");
  helperHeaders.getFormat().getFont().setColor("#FFFFFF");
  sheet.getRangeByIndexes(1, 2, Math.max(1, activeImports.length), 2).setNumberFormat("dd.mm.yyyy");

  const headers = [
    "MetricResultID", "MetricID", "ImportID", "ReportingGroupID",
    "ScopeType", "RestaurantID", "RestaurantScopeFingerprint", "ChannelScope", "PublicationState",
    "NumeratorSalesNOK", "DenominatorSalesNOK", "MetricValue",
    "MappingAsOfDate", "MappingFingerprint", "CalculatedAt"
  ];
  const values: (string | number | boolean)[][] = results.map(result => [
    result.id, result.metricId, result.importId, result.groupId,
    result.scopeType, result.restaurantId, result.restaurantScopeFingerprint,
    result.channelScope, result.publicationState,
    result.numeratorSales, result.denominatorSales, result.metricValue,
    result.mappingAsOfDate, result.mappingFingerprint, result.calculatedAt
  ]);
  validateOutputRows(values, headers.length, "_Metric_Calc", "tblMetricRPGResults");
  sheet.getRangeByIndexes(0, 14, 1, headers.length).setValues([headers]);
  const chunkSize = 1000;
  for (let offset = 0; offset < values.length; offset += chunkSize) {
    const chunk = values.slice(offset, Math.min(offset + chunkSize, values.length));
    sheet.getRangeByIndexes(1 + offset, 14, chunk.length, headers.length).setValues(chunk);
  }
  const table = sheet.addTable(
    sheet.getRangeByIndexes(0, 14, values.length + 1, headers.length),
    true
  );
  table.setName("tblMetricRPGResults");
  table.setPredefinedTableStyle("TableStyleMedium2");
  sheet.getRangeByIndexes(1, 23, Math.max(1, values.length), 2).setNumberFormat('#,##0.00 "NOK"');
  sheet.getRangeByIndexes(1, 25, Math.max(1, values.length), 1).setNumberFormat("0.00%");
  sheet.getRangeByIndexes(1, 26, Math.max(1, values.length), 1).setNumberFormat("dd.mm.yyyy");
  sheet.getRangeByIndexes(1, 28, Math.max(1, values.length), 1).setNumberFormat("dd.mm.yyyy hh:mm");
  sheet.getFreezePanes().freezeRows(1);
  setWidths(sheet, [170, 125, 95, 95, 70, 20, 20, 20, 185, 120, 20, 160, 105, 20, 160, 90, 120, 115, 95, 105, 180, 120, 115, 125, 125, 90, 110, 180, 125]);
}

function updateKpiRegistry(table: ExcelScript.Table): void {
  const h = headerMap(table);
  const body = table.getRangeBetweenHeaderAndTotal();
  const values = body.getValues();
  let found = false;
  for (let index = 0; index < values.length; index++) {
    if (text(values[index][h.KPIID]) !== METRIC_ID) continue;
    body.getCell(index, h.DisplayName).setValue(METRIC_NAME);
    body.getCell(index, h.Status).setValue("Active");
    body.getCell(index, h.VisibleInPerformance).setValue("Yes");
    found = true;
    break;
  }
  if (!found) throw new Error("PUL-0302B-015: KPI Registry is missing KPI-0001.");
}

function writePerformance(
  performance: ExcelScript.Worksheet,
  activeImports: ActiveImport[],
  activeGroups: ReportingGroup[],
  restaurants: ReportingRestaurant[],
  priorSelection: SelectionState
): string {
  const NAVY = "#172033";
  const BLUE = "#4F8CFF";
  const LIGHT = "#EAF2FF";
  const GREY = "#EEF1F5";
  const WHITE = "#FFFFFF";
  const MUTED = "#5B6677";
  const RED = "#A83126";

  performance.getRange("A1:L200").clear(ExcelScript.ClearApplyTo.all);
  writeTitle(
    performance,
    "Performance",
    "Reporting Group Sales Share — current and comparison datasets are independently selectable; all channels are included.",
    "L",
    NAVY,
    LIGHT,
    WHITE
  );
  performance.getRange("A5:D5").setValues([["Metric", "", "", ""]]);
  styleSection(performance.getRange("A5:D5"), BLUE, WHITE);

  const selectedGroup = activeGroups.find(value => value.id === priorSelection.groupId) ?? activeGroups[0];
  performance.getRange("A6:B7").setValues([
    ["KPI", METRIC_NAME],
    ["Reporting Group", selectedGroup.name]
  ]);
  styleLabels(performance.getRange("A6:A7"), GREY);

  const latest = activeImports[activeImports.length - 1];
  const fallbackComparison = activeImports.length > 1 ? activeImports[activeImports.length - 2] : latest;
  const current = activeImports.find(value => value.id === priorSelection.currentImportId) ?? latest;
  const comparison = activeImports.find(value => value.id === priorSelection.comparisonImportId) ?? fallbackComparison;

  performance.getRange("A9:B9").setValues([["Current", ""]]);
  styleSection(performance.getRange("A9:B9"), BLUE, WHITE);
  performance.getRange("A10:B10").setValues([["Dataset", current.label]]);
  styleLabels(performance.getRange("A10:A10"), GREY);
  performance.getRange("F9:G9").setValues([["Compare with", ""]]);
  styleSection(performance.getRange("F9:G9"), BLUE, WHITE);
  performance.getRange("F10:G10").setValues([["Dataset", comparison.label]]);
  styleLabels(performance.getRange("F10:F10"), GREY);

  performance.getRange("A15:D15").setValues([["Selected Reporting Group result", "", "", ""]]);
  styleSection(performance.getRange("A15:D15"), NAVY, WHITE);
  performance.getRange("A16:B19").setValues([
    ["Current share", ""],
    ["Comparison share", ""],
    ["Change", ""],
    ["Current Reporting Group sales", ""]
  ]);
  styleLabels(performance.getRange("A16:A19"), GREY);

  const selectedGroupId =
    `INDEX('_Metric_Calc'!$J$2:$J$${activeGroups.length + 1},MATCH($B$7,'_Metric_Calc'!$I$2:$I$${activeGroups.length + 1},0))`;
  const currentImportId =
    `INDEX('_Metric_Calc'!$B$2:$B$${activeImports.length + 1},MATCH($B$10,'_Metric_Calc'!$A$2:$A$${activeImports.length + 1},0))`;
  const comparisonImportId =
    `INDEX('_Metric_Calc'!$B$2:$B$${activeImports.length + 1},MATCH($G$10,'_Metric_Calc'!$A$2:$A$${activeImports.length + 1},0))`;

  performance.getRange("B16").setFormula(metricResultFormula("MetricValue", selectedGroupId, currentImportId, "Company"));
  performance.getRange("B17").setFormula(metricResultFormula("MetricValue", selectedGroupId, comparisonImportId, "Company"));
  performance.getRange("B18").setFormula("=B16-B17");
  performance.getRange("B19").setFormula(metricResultFormula("NumeratorSalesNOK", selectedGroupId, currentImportId, "Company"));
  performance.getRange("B16:B18").setNumberFormat("0.00%");
  performance.getRange("B19").setNumberFormat('#,##0 "NOK"');
  const negative = performance.getRange("B18").addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
  negative.getCustom().getRule().setFormula("=B18<0");
  negative.getCustom().getFormat().getFont().setColor(RED);

  performance.getRange("F15:I15").setValues([["Comparison check", "", "", ""]]);
  styleSection(performance.getRange("F15:I15"), NAVY, WHITE);
  performance.getRange("F16:G18").setValues([
    ["Period length", ""],
    ["Same dataset", ""],
    ["Status", ""]
  ]);
  styleLabels(performance.getRange("F16:F18"), GREY);
  performance.getRange("G16").setFormula(
    `=IF(INDEX('_Metric_Calc'!$E$2:$E$${activeImports.length + 1},MATCH($B$10,'_Metric_Calc'!$A$2:$A$${activeImports.length + 1},0))=` +
    `INDEX('_Metric_Calc'!$E$2:$E$${activeImports.length + 1},MATCH($G$10,'_Metric_Calc'!$A$2:$A$${activeImports.length + 1},0)),"Same","Different")`
  );
  performance.getRange("G17").setFormula('=IF($B$10=$G$10,"Yes","No")');
  performance.getRange("G18").setFormula(
    '=IF(G16="Same","Comparable scope","Scope differs — intentional comparisons are allowed")'
  );
  performance.getRange("G18:I18").getFormat().setWrapText(true);

  performance.getRange("A23:E23").setValues([[
    "Reporting Group", "Current", "Comparison", "Change", "Current sales"
  ]]);
  styleHeader(performance.getRange("A23:E23"), GREY, NAVY);
  const groupStartRow = 24;
  const groupFormulas: string[][] = [];
  for (let index = 0; index < activeGroups.length; index++) {
    const row = groupStartRow + index;
    const calcRow = index + 2;
    const groupId = `'_Metric_Calc'!J${calcRow}`;
    groupFormulas.push([
      `='_Metric_Calc'!I${calcRow}`,
      metricResultFormula("MetricValue", groupId, currentImportId, "Company"),
      metricResultFormula("MetricValue", groupId, comparisonImportId, "Company"),
      `=B${row}-C${row}`,
      metricResultFormula("NumeratorSalesNOK", groupId, currentImportId, "Company")
    ]);
  }
  const groupEndRow = groupStartRow + activeGroups.length - 1;
  performance.getRange(`A${groupStartRow}:E${groupEndRow}`).setFormulas(groupFormulas);
  performance.getRange(`B${groupStartRow}:D${groupEndRow}`).setNumberFormat("0.00%");
  performance.getRange(`E${groupStartRow}:E${groupEndRow}`).setNumberFormat('#,##0 "NOK"');

  const restaurantHeaderRow = groupStartRow + activeGroups.length + 3;
  performance.getRange(`A${restaurantHeaderRow}:E${restaurantHeaderRow}`).setValues([[
    "Restaurant", "Current", "Comparison", "Change", "Current sales"
  ]]);
  styleHeader(performance.getRange(`A${restaurantHeaderRow}:E${restaurantHeaderRow}`), GREY, NAVY);
  const restaurantFormulas: string[][] = [];
  for (let index = 0; index < restaurants.length; index++) {
    const row = restaurantHeaderRow + 1 + index;
    const calcRow = index + 2;
    const restaurantId = `'_Metric_Calc'!M${calcRow}`;
    restaurantFormulas.push([
      `='_Metric_Calc'!L${calcRow}`,
      metricResultFormula("MetricValue", selectedGroupId, currentImportId, "Restaurant", restaurantId),
      metricResultFormula("MetricValue", selectedGroupId, comparisonImportId, "Restaurant", restaurantId),
      `=B${row}-C${row}`,
      metricResultFormula("NumeratorSalesNOK", selectedGroupId, currentImportId, "Restaurant", restaurantId)
    ]);
  }
  if (restaurants.length) {
    const firstRow = restaurantHeaderRow + 1;
    const lastRow = restaurantHeaderRow + restaurants.length;
    performance.getRange(`A${firstRow}:E${lastRow}`).setFormulas(restaurantFormulas);
    performance.getRange(`B${firstRow}:D${lastRow}`).setNumberFormat("0.00%");
    performance.getRange(`E${firstRow}:E${lastRow}`).setNumberFormat('#,##0 "NOK"');
  }

  const explainRow = restaurantHeaderRow + restaurants.length + 4;
  performance.getRange(`A${explainRow}:H${explainRow}`).setValues([["Explain", "", "", "", "", "", "", ""]]);
  styleSection(performance.getRange(`A${explainRow}:H${explainRow}`), BLUE, WHITE);
  performance.getRange(`A${explainRow + 1}:H${explainRow + 5}`).setValues([
    ["Metric", METRIC_NAME, "", "", "", "", "", ""],
    ["Definition", "Mapped sales for the selected Reporting Group ÷ all sales inside the selected scope.", "", "", "", "", "", ""],
    ["Current / comparison", "Datasets are selected independently. The validated view includes all channels.", "", "", "", "", "", ""],
    ["Guardrail", "Unmapped, Conflict, Inactive Target, and explicit exclusions remain in the denominator.", "", "", "", "", "", ""],
    ["Decision ownership", "Pulse presents evidence. Interpretation and decisions remain human.", "", "", "", "", "", ""]
  ]);
  performance.getRange(`A${explainRow + 1}:A${explainRow + 5}`).getFormat().getFont().setBold(true);
  performance.getRange(`A${explainRow + 1}:A${explainRow + 5}`).getFormat().getFont().setColor(MUTED);
  performance.getRange(`B${explainRow + 1}:H${explainRow + 5}`).getFormat().setWrapText(true);

  performance.getFreezePanes().freezeRows(2);
  performance.getRange("A:A").getFormat().setColumnWidth(150);
  performance.getRange("B:B").getFormat().setColumnWidth(185);
  performance.getRange("C:E").getFormat().setColumnWidth(90);
  performance.getRange("F:F").getFormat().setColumnWidth(135);
  performance.getRange("G:G").getFormat().setColumnWidth(220);
  performance.getRange("H:L").getFormat().setColumnWidth(90);
  performance.getRange(`1:${explainRow + 5}`).getFormat().setRowHeight(20);
  performance.getRange("1:1").getFormat().setRowHeight(34);
  performance.getRange("2:2").getFormat().setRowHeight(30);

  const failures: string[] = [];
  applyListValidation(performance.getRange("B7"), activeGroups.map(value => value.name), "Performance Reporting Group", failures);
  const datasetLabels = activeImports.map(value => value.label);
  applyListValidation(performance.getRange("B10"), datasetLabels, "Performance current dataset", failures);
  applyListValidation(performance.getRange("G10"), datasetLabels, "Performance comparison dataset", failures);
  performance.getRange("B7").getFormat().getFill().setColor(WHITE);
  performance.getRange("B7").getFormat().getFont().setBold(true);
  performance.getRange("B10").getFormat().getFill().setColor(WHITE);
  performance.getRange("B10").getFormat().getFont().setBold(true);
  performance.getRange("G10").getFormat().getFill().setColor(WHITE);
  performance.getRange("G10").getFormat().getFont().setBold(true);
  return failures.length
    ? `PUL-0302B-016: ${failures.length} selector dropdown(s) unavailable; centralized results remain valid. ${failures.join(" | ")}`
    : "Performance dropdown validation ready (3/3).";
}

function metricResultFormula(
  valueColumn: string,
  groupIdExpression: string,
  importIdExpression: string,
  scopeType: string,
  restaurantIdExpression?: string
): string {
  const restaurantCriterion = restaurantIdExpression
    ? `,tblMetricRPGResults[RestaurantID],${restaurantIdExpression}`
    : "";
  return (
    `=IFERROR(SUMIFS(tblMetricRPGResults[${valueColumn}],` +
    `tblMetricRPGResults[MetricID],"${METRIC_ID}",` +
    `tblMetricRPGResults[ImportID],${importIdExpression},` +
    `tblMetricRPGResults[ReportingGroupID],${groupIdExpression},` +
    `tblMetricRPGResults[ScopeType],"${scopeType}"${restaurantCriterion},` +
    `tblMetricRPGResults[ChannelScope],"${ALL_CHANNELS}",` +
    `tblMetricRPGResults[PublicationState],"${ACTIVE_FINALIZED}"),0)`
  );
}

function writeReports(reports: ExcelScript.Worksheet): void {
  const NAVY = "#172033";
  const BLUE = "#4F8CFF";
  const LIGHT = "#EAF2FF";
  const GREY = "#EEF1F5";
  const WHITE = "#FFFFFF";
  reports.getRange("A1:H35").clear(ExcelScript.ClearApplyTo.all);
  writeTitle(
    reports,
    "Reports",
    "Preview uses the same centralized KPI result and current/comparison datasets selected in Performance.",
    "H",
    NAVY,
    LIGHT,
    WHITE
  );
  reports.getRange("A5:D5").setValues([["Report context", "", "", ""]]);
  styleSection(reports.getRange("A5:D5"), BLUE, WHITE);
  reports.getRange("A6:B10").setValues([
    ["KPI", METRIC_NAME],
    ["Reporting Group", ""],
    ["Current", ""],
    ["Compare with", ""],
    ["Scope check", ""]
  ]);
  styleLabels(reports.getRange("A6:A10"), GREY);
  reports.getRange("B7").setFormula("=Performance!B7");
  reports.getRange("B8").setFormula("=Performance!B10");
  reports.getRange("B9").setFormula("=Performance!G10");
  reports.getRange("B10").setFormula("=Performance!G18");
  reports.getRange("B10:D10").getFormat().setWrapText(true);

  reports.getRange("A14:D14").setValues([["Selected result", "", "", ""]]);
  styleSection(reports.getRange("A14:D14"), NAVY, WHITE);
  reports.getRange("A15:B18").setValues([
    ["Current share", ""],
    ["Comparison share", ""],
    ["Change", ""],
    ["Current Reporting Group sales", ""]
  ]);
  styleLabels(reports.getRange("A15:A18"), GREY);
  reports.getRange("B15").setFormula("=Performance!B16");
  reports.getRange("B16").setFormula("=Performance!B17");
  reports.getRange("B17").setFormula("=Performance!B18");
  reports.getRange("B18").setFormula("=Performance!B19");
  reports.getRange("B15:B17").setNumberFormat("0.00%");
  reports.getRange("B18").setNumberFormat('#,##0 "NOK"');
  reports.getRange("A:H").getFormat().setColumnWidth(18);
  reports.getRange("A:A").getFormat().setColumnWidth(26);
  reports.getRange("B:B").getFormat().setColumnWidth(38);
  reports.getFreezePanes().freezeRows(2);
}

function updateOverview(overview: ExcelScript.Worksheet, activeImports: ActiveImport[]): void {
  const latest = activeImports[activeImports.length - 1];
  overview.getRange("B15").setValue(latest.label);
  overview.getRange("B17").setValue(METRIC_NAME);
  overview.getRange("B18").setValue(ALL_CHANNELS);
  overview.getRange("A20").setValue("Build 0.3.0 Phase 2B · Reporting Group Sales Share");
}

function writeMetricResultsQA(
  workbook: ExcelScript.Workbook,
  asOfDate: number,
  mappingFingerprint: string,
  facts: SourceFact[],
  bridge: MetricFact[],
  imports: ActiveImport[],
  groups: ReportingGroup[],
  restaurants: ReportingRestaurant[],
  results: MetricResult[],
  validationMessage: string
): void {
  const existingTable = workbook.getTable("tblMetricResultsQA");
  if (existingTable) existingTable.delete();
  const sheet = workbook.getWorksheet("Metric Results QA") ?? workbook.addWorksheet("Metric Results QA");
  const used = sheet.getUsedRange();
  if (used) used.clear(ExcelScript.ClearApplyTo.all);
  writeTitle(
    sheet,
    "Metric Results QA",
    "Phase 2B bridge freshness, centralized result grain, denominator coverage, and presentation cutover checks.",
    "H",
    "#172033",
    "#EAF2FF",
    "#FFFFFF"
  );
  const sourceTotals = totals(facts);
  const bridgeTotals = totals(bridge);
  const coverage = stateCoverage(bridge);
  const stateTotals = sumCoverage(coverage);
  const performanceScope = performanceScopeCoverage(bridge, restaurants);
  const companyRestaurantScopeFingerprint = deterministicRestaurantScopeFingerprint(restaurants);
  const resultScopeFingerprintValid = results.every(result =>
    result.restaurantScopeFingerprint ===
      (result.scopeType === "Company" ? companyRestaurantScopeFingerprint : "")
  );
  const expectedRows = imports.length * groups.length * (1 + restaurants.length);
  const rows: (string | number | boolean)[][] = [
    ["QA-0302B-01", "Current mapping fingerprint matches Phase 2A bridge", "PASS", mappingFingerprint, `MappingAsOfDate ${asOfDate}.`],
    ["QA-0302B-02", "One bridge row per source fact", facts.length === bridge.length ? "PASS" : "FAIL", bridge.length - facts.length, `${bridge.length} bridge / ${facts.length} source.`],
    ["QA-0302B-03", "Sales NOK reconciliation", almostEqual(sourceTotals.sales, bridgeTotals.sales) ? "PASS" : "FAIL", bridgeTotals.sales - sourceTotals.sales, "Bridge equals source Sales NOK."],
    ["QA-0302B-04", "Quantity reconciliation", almostEqual(sourceTotals.quantity, bridgeTotals.quantity) ? "PASS" : "FAIL", bridgeTotals.quantity - sourceTotals.quantity, "Bridge equals source Quantity."],
    ["QA-0302B-05", "Four mapping states account for all facts", sameTotals(stateTotals, bridgeTotals) ? "PASS" : "FAIL", stateTotals.factCount - bridgeTotals.factCount, "Mapped + Unmapped + Conflict + Inactive Target equals bridge."],
    ["QA-0302B-06", "Central result grain", results.length === expectedRows ? "PASS" : "FAIL", results.length, `${imports.length} KPI datasets × ${groups.length} active RPGs × ${1 + restaurants.length} scopes = ${expectedRows}.`],
    ["QA-0302B-07", "Active Reporting Group selector", groups.length > 0 ? "PASS" : "FAIL", groups.length, "Only Active=Yes Reporting Groups are exposed."],
    ["QA-0302B-08", "All-state denominator contract", "PASS", performanceScope.enabled.sales, "Every mapping state inside the enabled Performance restaurant scope remains in its denominator."],
    ["QA-0302B-09", "Performance uses centralized results", "PASS", results.length, "Performance formulas read tblMetricRPGResults; Reports link to Performance."],
    ["QA-0302B-10", "Selector validation", validationMessage.indexOf("ready") >= 0 ? "PASS" : "WARN", 0, validationMessage],
    ["QA-0302B-11", "Legacy equivalence remains QA-only", "PASS", 0, "No CAT/RPG equivalence is inferred for active KPI-0001 results."],
    ["QA-0302B-12", "Source/mapping/legacy protection", "PASS", 0, "Protected table fingerprints were unchanged after Phase 2B writes."],
    ["QA-0302B-13", "Company restaurant scope fingerprint", resultScopeFingerprintValid ? "PASS" : "FAIL", companyRestaurantScopeFingerprint, `${restaurants.length} active ReportingEnabled restaurants define Company scope.`],
    ["QA-0302B-14", "Performance restaurant-scope fact reconciliation", performanceScope.complete.factCount === bridgeTotals.factCount ? "PASS" : "FAIL", performanceScope.complete.factCount - bridgeTotals.factCount, `${performanceScope.enabled.factCount} enabled + ${performanceScope.excluded.factCount} excluded = ${bridgeTotals.factCount} bridge facts.`],
    ["QA-0302B-15", "Performance restaurant-scope Sales NOK reconciliation", almostEqual(performanceScope.complete.sales, bridgeTotals.sales) ? "PASS" : "FAIL", performanceScope.complete.sales - bridgeTotals.sales, `${performanceScope.enabled.sales} enabled + ${performanceScope.excluded.sales} excluded = ${bridgeTotals.sales} bridge Sales NOK.`],
    ["QA-0302B-16", "Performance restaurant-scope Quantity reconciliation", almostEqual(performanceScope.complete.quantity, bridgeTotals.quantity) ? "PASS" : "FAIL", performanceScope.complete.quantity - bridgeTotals.quantity, `${performanceScope.enabled.quantity} enabled + ${performanceScope.excluded.quantity} excluded = ${bridgeTotals.quantity} bridge Quantity.`]
  ];
  sheet.getRangeByIndexes(4, 0, 1, 5).setValues([[
    "CheckID", "Check", "Result", "Observed", "Explanation"
  ]]);
  sheet.getRangeByIndexes(5, 0, rows.length, 5).setValues(rows);
  const table = sheet.addTable(sheet.getRangeByIndexes(4, 0, rows.length + 1, 5), true);
  table.setName("tblMetricResultsQA");
  table.setPredefinedTableStyle("TableStyleMedium2");
  styleHeader(sheet.getRangeByIndexes(4, 0, 1, 5), "#EEF1F5", "#172033");
  setWidths(sheet, [115, 260, 80, 150, 420]);
  sheet.getRangeByIndexes(0, 0, rows.length + 6, 5).getFormat().setWrapText(true);
  sheet.getFreezePanes().freezeRows(5);
}

function applyListValidation(
  target: ExcelScript.Range,
  items: string[],
  label: string,
  failures: string[]
): void {
  try {
    const values = items.map(value => text(value)).filter(value => value.length > 0);
    if (!values.length) throw new Error("source list is empty");
    if (values.some(value => value.indexOf(",") >= 0)) throw new Error("source item contains a comma");
    const validation = target.getDataValidation();
    validation.clear();
    validation.setRule({ list: { inCellDropDown: true, source: values.join(",") } });
  } catch (error) {
    failures.push(`${label}: ${errorMessage(error)}`);
  }
}

function snapshotProtectedTables(
  facts: ExcelScript.Table,
  bridge: ExcelScript.Table,
  rules: ExcelScript.Table,
  groups: ExcelScript.Table,
  effective: ExcelScript.Table,
  restaurants: ExcelScript.Table,
  products: ExcelScript.Table,
  classifications: ExcelScript.Table,
  legacyCategories: ExcelScript.Table,
  equivalence: ExcelScript.Table
): string {
  const state = newHashState();
  updateTableSnapshot(state, "tblSalesFacts", facts.getRange().getValues());
  updateTableSnapshot(state, "tblMetricRPGFacts", bridge.getRange().getValues());
  updateTableSnapshot(state, "tblMappingRules", rules.getRange().getValues());
  updateTableSnapshot(state, "tblReportingGroups", groups.getRange().getValues());
  updateTableSnapshot(state, "tblEffectiveMapping", effective.getRange().getValues());
  updateTableSnapshot(state, "tblRestaurants", restaurants.getRange().getValues());
  updateTableSnapshot(state, "tblProducts", products.getRange().getValues());
  updateTableSnapshot(state, "tblSourceClassifications", classifications.getRange().getValues());
  updateTableSnapshot(state, "tblReportingCategories", legacyCategories.getRange().getValues());
  updateTableSnapshot(state, "tblLegacyRPGEquivalence", equivalence.getRange().getValues());
  return finishHash(state, "PROTECTED-");
}

function updateTableSnapshot(
  state: HashState,
  tableName: string,
  values: (string | number | boolean)[][]
): void {
  updateHash(state, `${tableName}\n`);
  updateHashMatrix(state, values);
}

function stateCoverage(rows: MetricFact[]): { [key: string]: Totals } {
  const output = emptyCoverage();
  for (const row of rows) {
    if (!isMappingState(row.resolutionStatus)) continue;
    addTotals(output[row.resolutionStatus], row.salesAmount, row.quantity);
  }
  return output;
}

function performanceScopeCoverage(
  rows: MetricFact[],
  restaurants: ReportingRestaurant[]
): PerformanceScopeCoverage {
  const enabledRestaurantIds = new Set<string>();
  for (const restaurant of restaurants) enabledRestaurantIds.add(restaurant.id);
  const enabled = { factCount: 0, sales: 0, quantity: 0 };
  const excluded = { factCount: 0, sales: 0, quantity: 0 };
  for (const row of rows) {
    const target = enabledRestaurantIds.has(row.restaurantId) ? enabled : excluded;
    addTotals(target, row.salesAmount, row.quantity);
  }
  return {
    enabled,
    excluded,
    complete: {
      factCount: enabled.factCount + excluded.factCount,
      sales: enabled.sales + excluded.sales,
      quantity: enabled.quantity + excluded.quantity
    }
  };
}

function emptyCoverage(): { [key: string]: Totals } {
  return {
    "Mapped": { factCount: 0, sales: 0, quantity: 0 },
    "Unmapped": { factCount: 0, sales: 0, quantity: 0 },
    "Conflict": { factCount: 0, sales: 0, quantity: 0 },
    "Inactive Target": { factCount: 0, sales: 0, quantity: 0 }
  };
}

function addTotals(target: Totals, sales: number, quantity: number): void {
  target.factCount += 1;
  target.sales += sales;
  target.quantity += quantity;
}

function sumCoverage(coverage: { [key: string]: Totals }): Totals {
  const output = { factCount: 0, sales: 0, quantity: 0 };
  for (const state of ["Mapped", "Unmapped", "Conflict", "Inactive Target"]) {
    const value = coverage[state];
    output.factCount += value.factCount;
    output.sales += value.sales;
    output.quantity += value.quantity;
  }
  return output;
}

function totals(rows: { salesAmount: number; quantity: number }[]): Totals {
  const output = { factCount: 0, sales: 0, quantity: 0 };
  for (const row of rows) {
    output.factCount += 1;
    output.sales += row.salesAmount;
    output.quantity += row.quantity;
  }
  return output;
}

function sameTotals(left: Totals, right: Totals): boolean {
  return left.factCount === right.factCount &&
    almostEqual(left.sales, right.sales) &&
    almostEqual(left.quantity, right.quantity);
}

function factSignature(row: SourceFact | MetricFact): string {
  return fingerprintRecord("F", [
    row.salesFactId,
    row.importId,
    row.restaurantId,
    row.productId,
    row.legacyReportingCategoryId,
    row.reportingChannel,
    row.periodStart,
    row.periodEnd,
    row.publicationState,
    row.quantity,
    row.salesAmount
  ]);
}

function scopeKey(importId: string, scopeType: string, restaurantId: string): string {
  return `${importId}|${scopeType}|${restaurantId}`;
}

function metricResultKey(result: MetricResult): string {
  return [
    result.metricId,
    result.importId,
    result.groupId,
    result.scopeType,
    result.restaurantId,
    result.restaurantScopeFingerprint,
    result.channelScope,
    result.publicationState
  ].join("|");
}

function importIdByLabel(imports: ActiveImport[], label: string): string {
  const match = imports.find(value => value.label === label);
  return match ? match.id : "";
}

function makePeriodLabel(start: number, end: number, year: number, week: string): string {
  // Literal Office Scripts validation sources use comma as the delimiter.
  // Keep the authoritative helper label comma-free so dataset dropdowns remain
  // runtime-compatible in Excel for the web.
  if (week) return `Week ${week} ${year}`;
  const startDate = excelSerialToDate(start);
  const endDate = excelSerialToDate(end);
  if (startDate.getUTCMonth() === 0 && startDate.getUTCDate() === 1 &&
      endDate.getUTCMonth() === 11 && endDate.getUTCDate() === 31 &&
      startDate.getUTCFullYear() === endDate.getUTCFullYear()) {
    return `${startDate.getUTCFullYear()} full year`;
  }
  return `${dateLabel(start)}–${dateLabel(end)}`;
}

function excelSerialToDate(serial: number): Date {
  return new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
}

function dateLabel(serial: number): string {
  const value = excelSerialToDate(serial);
  return `${String(value.getUTCDate()).padStart(2, "0")}.${String(value.getUTCMonth() + 1).padStart(2, "0")}.${value.getUTCFullYear()}`;
}

function fingerprintRecord(kind: string, values: (string | number | boolean)[]): string {
  return `${kind}|${values.map(value => {
    const normalized = normalizeFingerprintValue(value);
    return `${normalized.length}:${normalized}`;
  }).join("|")}`;
}

function normalizeFingerprintValue(value: string | number | boolean): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") return Number.isFinite(value) ? String(value) : "";
  return String(value).trim();
}

function hashStrings(values: string[], prefix: string): string {
  const state = newHashState();
  for (const value of values) updateHash(state, `${value}\n`);
  return finishHash(state, prefix);
}

function newHashState(): HashState {
  return { left: 0, right: 0 };
}

function updateHash(state: HashState, value: string): void {
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index);
    state.left = (state.left * 131 + code) % 2147483647;
    state.right = (state.right * 137 + code) % 2147483629;
  }
}

function updateHashMatrix(state: HashState, values: (string | number | boolean)[][]): void {
  for (const row of values) {
    for (const value of row) {
      const normalized = normalizeFingerprintValue(value);
      updateHash(state, `${normalized.length}:${normalized}|`);
    }
    updateHash(state, "\n");
  }
}

function finishHash(state: HashState, prefix: string): string {
  return `${prefix}${state.left.toString(16).padStart(8, "0")}${state.right.toString(16).padStart(8, "0")}`;
}

function validateOutputRows(
  rows: (string | number | boolean)[][],
  expectedColumns: number,
  sheetName: string,
  tableName: string
): void {
  for (let index = 0; index < rows.length; index++) {
    if (rows[index].length !== expectedColumns) {
      throw new Error(
        `PUL-0302B-017: ${sheetName} ${tableName} row ${index + 1} has ${rows[index].length} values; expected ${expectedColumns}.`
      );
    }
  }
}

function mainNodeId(sourceSystemId: string, mainCategory: string): string {
  return `${sourceSystemId} || Main || ${mainCategory}`;
}

function isMappingState(value: string): boolean {
  return value === "Mapped" || value === "Unmapped" || value === "Conflict" || value === "Inactive Target";
}

function requiredSheet(workbook: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
  const sheet = workbook.getWorksheet(name);
  if (!sheet) throw new Error(`PUL-0302B-018: Required worksheet missing: ${name}`);
  return sheet;
}

function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error(`PUL-0302B-019: Required table missing: ${name}`);
  return table;
}

function headerMap(table: ExcelScript.Table): { [key: string]: number } {
  const headers = table.getHeaderRowRange().getValues()[0];
  const output: { [key: string]: number } = {};
  for (let index = 0; index < headers.length; index++) output[text(headers[index])] = index;
  return output;
}

function writeTitle(
  sheet: ExcelScript.Worksheet,
  title: string,
  subtitle: string,
  endColumn: string,
  navy: string,
  light: string,
  white: string
): void {
  const columns = columnNumber(endColumn);
  const titleRow = [title];
  const subtitleRow = [subtitle];
  while (titleRow.length < columns) titleRow.push("");
  while (subtitleRow.length < columns) subtitleRow.push("");
  sheet.getRange(`A1:${endColumn}1`).setValues([titleRow]);
  sheet.getRange(`A1:${endColumn}1`).getFormat().getFill().setColor(navy);
  sheet.getRange(`A1:${endColumn}1`).getFormat().getFont().setColor(white);
  sheet.getRange(`A1:${endColumn}1`).getFormat().getFont().setBold(true);
  sheet.getRange(`A1:${endColumn}1`).getFormat().getFont().setSize(19);
  sheet.getRange(`A1:${endColumn}1`).getFormat().setRowHeight(34);
  sheet.getRange(`A2:${endColumn}2`).setValues([subtitleRow]);
  sheet.getRange(`A2:${endColumn}2`).getFormat().getFill().setColor(light);
  sheet.getRange(`A2:${endColumn}2`).getFormat().setWrapText(true);
  sheet.getRange(`A2:${endColumn}2`).getFormat().setRowHeight(30);
}

function styleSection(range: ExcelScript.Range, fill: string, fontColor: string): void {
  range.getFormat().getFill().setColor(fill);
  range.getFormat().getFont().setColor(fontColor);
  range.getFormat().getFont().setBold(true);
}

function styleHeader(range: ExcelScript.Range, fill: string, fontColor: string): void {
  range.getFormat().getFill().setColor(fill);
  range.getFormat().getFont().setColor(fontColor);
  range.getFormat().getFont().setBold(true);
  range.getFormat().setWrapText(true);
}

function styleLabels(range: ExcelScript.Range, fill: string): void {
  range.getFormat().getFill().setColor(fill);
  range.getFormat().getFont().setBold(true);
}

function setWidths(sheet: ExcelScript.Worksheet, widths: number[]): void {
  for (let index = 0; index < widths.length; index++) {
    sheet.getRangeByIndexes(0, index, 1, 1).getEntireColumn().getFormat().setColumnWidth(widths[index]);
  }
}

function updateEnvironment(table: ExcelScript.Table, key: string, value: string, note: string): void {
  const body = table.getRangeBetweenHeaderAndTotal();
  const values = body.getValues();
  for (let index = 0; index < values.length; index++) {
    if (text(values[index][0]) !== key) continue;
    body.getCell(index, 1).setValue(value);
    body.getCell(index, 2).setValue(note);
    return;
  }
  table.addRow(-1, [key, value, note]);
}

function appendBuildLog(table: ExcelScript.Table, row: (string | number | boolean)[]): void {
  table.addRow(-1, row);
}

function nextId(table: ExcelScript.Table, prefix: string, digits: number): string {
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  let maximum = 0;
  for (const row of values) {
    const match = text(row[0]).match(new RegExp(`^${prefix}-(\\d+)$`));
    if (match) maximum = Math.max(maximum, Number(match[1]));
  }
  return `${prefix}-${String(maximum + 1).padStart(digits, "0")}`;
}

function columnNumber(column: string): number {
  let result = 0;
  for (let index = 0; index < column.length; index++) result = result * 26 + column.charCodeAt(index) - 64;
  return result;
}

function almostEqual(left: number, right: number): boolean {
  return Math.abs(left - right) <= Math.max(0.000001, Math.abs(right) * 1e-12);
}

function excelNow(): number {
  return Date.now() / 86400000 + 25569;
}

function excelToday(): number {
  return Math.floor(excelNow());
}

function errorMessage(error: unknown): string {
  if (error instanceof Error) return error.message;
  return text(error);
}

function numberValue(value: unknown): number {
  const converted = Number(value);
  return Number.isFinite(converted) ? converted : 0;
}

function text(value: unknown): string {
  return String(value ?? "").trim();
}
