/**
 * Pulse Build 0.3.0 Phase 2A — Reporting Group metric contract and bridge.
 *
 * Prerequisite sequence:
 * 1. Start from excel/Pulse_Build_0_2_0_QA.xlsx.xlsx.
 * 2. Run Build_0_3_0_Phase1.ts.
 * 3. Run this script while Effective Mapping is current for today.
 *
 * Build 0.3.0 mapping semantics are current-state analytical classification:
 * the current effective mapping is applied to historical facts without editing
 * those facts. Fact-date mapping/versioning is deliberately out of scope.
 *
 * Phase 2A is additive. It does not edit _Sales_Facts, _Metric_Calc,
 * Performance, Reports, KPI Registry, or the legacy KPI-0001 calculation path.
 */
function main(workbook: ExcelScript.Workbook): string {
  const factsTable = requiredTable(workbook, "tblSalesFacts");
  const productsTable = requiredTable(workbook, "tblProducts");
  const classificationsTable = requiredTable(workbook, "tblSourceClassifications");
  const groupsTable = requiredTable(workbook, "tblReportingGroups");
  const rulesTable = requiredTable(workbook, "tblMappingRules");
  const effectiveTable = requiredTable(workbook, "tblEffectiveMapping");
  const categoriesTable = requiredTable(workbook, "tblReportingCategories");
  const environmentTable = requiredTable(workbook, "tblEnvironment");
  const buildLogTable = requiredTable(workbook, "tblBuildLog");
  const kpiTable = requiredTable(workbook, "tblKPIRegistry");

  const factsBefore = snapshotTable(factsTable);
  const protectedBefore = snapshotProtectedSurfaces(workbook, kpiTable);
  const asOfDate = excelToday();
  const groups = readGroups(groupsTable);
  const rules = readRules(rulesTable);
  const products = buildHierarchy(classificationsTable, productsTable);
  const expectedResolutions = products.map(product => resolveProduct(product, rules, groups, asOfDate));
  const materializedResolutions = readEffectiveMapping(effectiveTable);

  const freshnessErrors = validateEffectiveMappingFreshness(
    expectedResolutions,
    materializedResolutions,
    asOfDate
  );
  if (freshnessErrors.length) {
    throw new Error(
      `PUL-0302A-001: Effective Mapping is stale or invalid. Run Build_0_3_0_Phase1.ts and rerun Phase 2A. ${freshnessErrors.slice(0, 10).join(" | ")}`
    );
  }

  const mappingFingerprint = computeMappingFingerprint(
    asOfDate,
    groups,
    rules,
    products,
    expectedResolutions
  );
  const reorderedFingerprint = computeMappingFingerprint(
    asOfDate,
    groups.slice().reverse(),
    rules.slice().reverse(),
    products.slice().reverse(),
    expectedResolutions.slice().reverse()
  );
  if (mappingFingerprint !== reorderedFingerprint) {
    throw new Error("PUL-0302A-002: Mapping fingerprint is not reproducible across row ordering.");
  }

  const facts = readFacts(factsTable);
  const refreshAt = excelNow();
  const bridge = buildMetricBridge(
    facts,
    materializedResolutions,
    asOfDate,
    mappingFingerprint,
    refreshAt
  );

  writeMetricContract(workbook);
  const equivalenceTable = ensureLegacyEquivalenceTable(workbook, groups);
  const definitions = readEquivalenceDefinitions(equivalenceTable);
  const definitionErrors = validateEquivalenceDefinitions(
    definitions,
    groups,
    readLegacyCategoryIds(categoriesTable)
  );

  writeMetricBridge(workbook, bridge);
  const materializedBridge = readMetricBridge(requiredTable(workbook, "tblMetricRPGFacts"));
  const bridgeErrors = validateMaterializedBridge(
    facts,
    materializedBridge,
    asOfDate,
    mappingFingerprint
  );
  if (bridgeErrors.length) {
    throw new Error(`PUL-0302A-017: Materialized metric bridge is invalid. ${bridgeErrors.slice(0, 10).join(" | ")}`);
  }
  const scopes = buildScopes(facts);
  const activeScopes = scopes.filter(scope => scope.publicationState === "Active Finalized");
  const reconciliations = reconcileFactsAndBridge(facts, materializedBridge, scopes);
  const groupTotals = aggregateReportingGroups(materializedBridge, activeScopes);
  const crosswalk = buildLegacyRpgCrosswalk(materializedBridge, activeScopes);
  const comparisons = definitionErrors.length
    ? []
    : compareLegacyDefinitions(materializedBridge, definitions, activeScopes);

  const factsAfter = snapshotTable(factsTable);
  if (!sameSnapshot(factsBefore, factsAfter)) {
    throw new Error("PUL-0302A-003: tblSalesFacts changed during Phase 2A.");
  }
  const protectedAfter = snapshotProtectedSurfaces(workbook, kpiTable);
  if (protectedBefore !== protectedAfter) {
    throw new Error("PUL-0302A-004: _Metric_Calc, Performance, Reports, or KPI Registry changed during Phase 2A.");
  }

  writeMetricQA(
    workbook,
    asOfDate,
    mappingFingerprint,
    refreshAt,
    factsBefore,
    bridge.length,
    reconciliations,
    groupTotals,
    crosswalk,
    comparisons,
    definitions.length,
    definitionErrors
  );

  updateEnvironment(
    environmentTable,
    "BuildID",
    "0.3.0-Phase2A",
    "Reporting Group metric contract, immutable fact bridge, reconciliation, and legacy comparison only."
  );
  updateEnvironment(
    environmentTable,
    "BuildVersion",
    "0.3.0-Phase2A",
    "Performance, Reports, _Metric_Calc, and KPI-0001 remain on the validated legacy path."
  );
  appendBuildLog(buildLogTable, [
    nextId(buildLogTable, "LOG", 6),
    refreshAt,
    "0.3.0-Phase2A",
    "Refresh Reporting Group metric bridge",
    "Success",
    "Metric bridge + reconciliation",
    `${bridge.length} facts derived; mapping ${mappingFingerprint}; ${definitionErrors.length} equivalence configuration error(s).`
  ]);

  const definitionMessage = definitionErrors.length
    ? `${definitionErrors.length} legacy equivalence configuration error(s) surfaced in Metric Migration QA.`
    : definitions.length
      ? `${comparisons.length} legacy comparison row(s) refreshed.`
      : "Legacy equivalence table ready; no human-authored definitions are active yet.";
  return `Pulse 0.3.0 Phase 2A refreshed. ${bridge.length} facts, mapping ${mappingFingerprint}. ${definitionMessage}`;
}

type ReportingGroup = { id: string; name: string; active: string; sortOrder: number };
type MappingRule = {
  id: string; sourceSystemId: string; scopeType: string; nodeId: string;
  targetGroupId: string; effectiveFrom: number; effectiveTo: number; status: string;
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
type MetricScope = {
  publicationScope: string; publicationState: string; importId: string; channel: string;
};
type Totals = { factCount: number; sales: number; quantity: number };
type Reconciliation = {
  scope: MetricScope; source: Totals; bridge: Totals; coverage: { [key: string]: Totals };
  stateTotals: Totals; factVariance: number; salesVariance: number; quantityVariance: number;
  stateFactVariance: number; stateSalesVariance: number; stateQuantityVariance: number; result: string;
};
type GroupTotal = MetricScope & {
  groupId: string; groupName: string; factCount: number; sales: number; quantity: number;
};
type CrosswalkRow = MetricScope & {
  legacyCategoryId: string; groupId: string; groupName: string; resolutionStatus: string;
  factCount: number; sales: number; quantity: number;
};
type EquivalenceDefinition = {
  definitionId: string; groupId: string; legacyCategoryId: string;
  comparisonStatus: string; active: string; notes: string; approvedBy: string; approvedAt: number;
};
type LegacyComparison = MetricScope & {
  definitionId: string; comparisonStatus: string; groupId: string; legacyCategoryIds: string;
  scopeTotals: Totals; legacy: Totals; rpg: Totals; factVariance: number;
  salesVariance: number; quantityVariance: number; legacySalesShare: number;
  rpgSalesShare: number; shareVariance: number; result: string;
};
type Snapshot = { rows: number; sales: number; quantity: number; fingerprint: string };
type HashState = { left: number; right: number };

const MAPPING_STATES = ["Mapped", "Unmapped", "Conflict", "Inactive Target"];

function readGroups(table: ExcelScript.Table): ReportingGroup[] {
  const h = headerMap(table);
  const groups: ReportingGroup[] = [];
  const seen = new Set<string>();
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  for (const row of values) {
    const id = text(row[h.ReportingGroupID]);
    if (!id) continue;
    if (seen.has(id)) throw new Error(`PUL-0302A-005: Duplicate ReportingGroupID ${id}.`);
    seen.add(id);
    groups.push({
      id,
      name: text(row[h.ReportingGroupName]),
      active: text(row[h.Active]),
      sortOrder: numberValue(row[h.SortOrder])
    });
  }
  return groups;
}

function readRules(table: ExcelScript.Table): MappingRule[] {
  const h = headerMap(table);
  return table.getRangeBetweenHeaderAndTotal().getValues()
    .filter(row => text(row[h.MappingRuleID]))
    .map(row => ({
      id: text(row[h.MappingRuleID]),
      sourceSystemId: text(row[h.SourceSystemID]),
      scopeType: text(row[h.ScopeType]),
      nodeId: text(row[h.NodeID]),
      targetGroupId: text(row[h.TargetReportingGroupID]),
      effectiveFrom: numberValue(row[h.EffectiveFrom]),
      effectiveTo: numberValue(row[h.EffectiveTo]),
      status: text(row[h.Status])
    }));
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
  const result: ProductNode[] = [];
  const seen = new Set<string>();
  const ph = headerMap(products);
  const productValues = products.getRangeBetweenHeaderAndTotal().getValues();
  for (const row of productValues) {
    const productId = text(row[ph.ProductID]);
    if (!productId) continue;
    if (seen.has(productId)) throw new Error(`PUL-0302A-006: Duplicate ProductID ${productId}.`);
    seen.add(productId);
    const classificationId = text(row[ph.SourceClassificationID]);
    const classification = classificationById.get(classificationId);
    if (!classification) {
      throw new Error(`PUL-0302A-007: Product ${productId} references missing classification ${classificationId}.`);
    }
    result.push({
      productId,
      sourceSystemId: text(row[ph.SourceSystemID]),
      productName: text(row[ph.SourceProductName]),
      salesAccount: text(row[ph.SalesAccount]),
      sourceClassificationId: classificationId,
      mainCategory: classification.main,
      subCategory: classification.sub,
      mainNodeId: mainNodeId(classification.sourceSystemId, classification.main),
      subNodeId: classificationId
    });
  }
  return result;
}

function resolveProduct(
  product: ProductNode,
  rules: MappingRule[],
  groups: ReportingGroup[],
  asOfDate: number
): Resolution {
  const groupById = new Map<string, ReportingGroup>();
  for (const group of groups) groupById.set(group.id, group);
  const scopes = [
    { type: "SourceMainCategory", nodeId: product.mainNodeId },
    { type: "SourceSubCategory", nodeId: product.subNodeId },
    { type: "Product", nodeId: product.productId }
  ];
  const applicable: { [key: string]: MappingRule[] } = {};
  for (const scope of scopes) {
    applicable[scope.type] = rules.filter(rule =>
      rule.status === "Active" &&
      rule.sourceSystemId === product.sourceSystemId &&
      rule.scopeType === scope.type &&
      rule.nodeId === scope.nodeId &&
      (!rule.effectiveFrom || rule.effectiveFrom <= asOfDate) &&
      (!rule.effectiveTo || rule.effectiveTo >= asOfDate)
    );
  }
  const base = {
    ...product,
    mainRuleIds: ids(applicable.SourceMainCategory),
    mainTargetIds: targets(applicable.SourceMainCategory),
    subRuleIds: ids(applicable.SourceSubCategory),
    subTargetIds: targets(applicable.SourceSubCategory),
    productRuleIds: ids(applicable.Product),
    productTargetIds: targets(applicable.Product),
    asOfDate
  };
  for (const scopeType of ["Product", "SourceSubCategory", "SourceMainCategory"]) {
    const candidates = applicable[scopeType];
    if (candidates.length > 1) {
      return {
        ...base,
        effectiveGroupId: "",
        effectiveGroupName: "",
        resolutionSource: scopeType,
        resolutionState: "Explicit conflict",
        resolutionStatus: "Conflict",
        winningRuleId: ids(candidates)
      };
    }
    if (candidates.length === 1) {
      const rule = candidates[0];
      const group = groupById.get(rule.targetGroupId);
      return {
        ...base,
        effectiveGroupId: rule.targetGroupId,
        effectiveGroupName: group ? group.name : "",
        resolutionSource: scopeType,
        resolutionState: scopeType === "Product" ? "Explicit" : "Inherited",
        resolutionStatus: !group || group.active !== "Yes" ? "Inactive Target" : "Mapped",
        winningRuleId: rule.id
      };
    }
  }
  return {
    ...base,
    effectiveGroupId: "",
    effectiveGroupName: "",
    resolutionSource: "Unmapped",
    resolutionState: "Unmapped",
    resolutionStatus: "Unmapped",
    winningRuleId: ""
  };
}

function readEffectiveMapping(table: ExcelScript.Table): Resolution[] {
  const h = headerMap(table);
  return table.getRangeBetweenHeaderAndTotal().getValues()
    .filter(row => text(row[h.ProductID]))
    .map(row => ({
      productId: text(row[h.ProductID]),
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
    }));
}

function validateEffectiveMappingFreshness(
  expected: Resolution[],
  actual: Resolution[],
  asOfDate: number
): string[] {
  const errors: string[] = [];
  if (expected.length !== actual.length) {
    errors.push(`row count ${actual.length}; expected ${expected.length}`);
  }
  const actualByProduct = new Map<string, Resolution>();
  for (const row of actual) {
    if (actualByProduct.has(row.productId)) errors.push(`duplicate ProductID ${row.productId}`);
    actualByProduct.set(row.productId, row);
    if (row.asOfDate !== asOfDate) errors.push(`${row.productId} AsOfDate ${row.asOfDate}; expected ${asOfDate}`);
    if (!isMappingState(row.resolutionStatus)) errors.push(`${row.productId} unsupported status ${row.resolutionStatus}`);
  }
  for (const expectedRow of expected) {
    const actualRow = actualByProduct.get(expectedRow.productId);
    if (!actualRow) {
      errors.push(`missing product ${expectedRow.productId}`);
      continue;
    }
    if (fingerprintRecord("E", resolutionSignature(expectedRow)) !==
        fingerprintRecord("E", resolutionSignature(actualRow))) {
      errors.push(`stale product ${expectedRow.productId}`);
    }
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
  const records: string[] = [fingerprintRecord("V", ["PULSE-MAPPING-SEMANTIC-V1", asOfDate])];
  for (const group of groups) records.push(fingerprintRecord("G", [group.id, group.name, group.active, group.sortOrder]));
  for (const rule of rules) records.push(fingerprintRecord("R", [
    rule.id, rule.sourceSystemId, rule.scopeType, rule.nodeId, rule.targetGroupId,
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
  const facts: SourceFact[] = [];
  const seen = new Set<string>();
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  for (const row of values) {
    const salesFactId = text(row[h.SalesFactID]);
    if (!salesFactId) throw new Error("PUL-0302A-008: tblSalesFacts contains a blank SalesFactID.");
    if (seen.has(salesFactId)) throw new Error(`PUL-0302A-009: Duplicate SalesFactID ${salesFactId}.`);
    seen.add(salesFactId);
    const productId = text(row[h.ProductID]);
    if (!productId) throw new Error(`PUL-0302A-010: Sales fact ${salesFactId} has a blank ProductID.`);
    facts.push({
      salesFactId,
      importId: text(row[h.ImportID]),
      restaurantId: text(row[h.RestaurantID]),
      productId,
      legacyReportingCategoryId: text(row[h.ReportingCategoryID]),
      periodStart: numberValue(row[h.PeriodStart]),
      periodEnd: numberValue(row[h.PeriodEnd]),
      reportingChannel: text(row[h.ReportingChannel]),
      quantity: numberValue(row[h.Quantity]),
      salesAmount: numberValue(row[h.SalesAmount]),
      publicationState: text(row[h.PublicationState])
    });
  }
  return facts;
}

function buildMetricBridge(
  facts: SourceFact[],
  mappings: Resolution[],
  mappingAsOfDate: number,
  mappingFingerprint: string,
  metricRefreshAt: number
): MetricFact[] {
  const mappingByProduct = new Map<string, Resolution>();
  for (const mapping of mappings) {
    if (mappingByProduct.has(mapping.productId)) {
      throw new Error(`PUL-0302A-011: Duplicate Effective Mapping product ${mapping.productId}.`);
    }
    mappingByProduct.set(mapping.productId, mapping);
  }
  const output: MetricFact[] = [];
  for (const fact of facts) {
    const mapping = mappingByProduct.get(fact.productId);
    if (!mapping) {
      throw new Error(`PUL-0302A-012: Sales fact ${fact.salesFactId} has no Effective Mapping for ${fact.productId}.`);
    }
    if (!isMappingState(mapping.resolutionStatus)) {
      throw new Error(`PUL-0302A-013: Product ${fact.productId} has unsupported mapping status ${mapping.resolutionStatus}.`);
    }
    output.push({
      ...fact,
      effectiveGroupId: mapping.effectiveGroupId,
      effectiveGroupName: mapping.effectiveGroupName,
      resolutionSource: mapping.resolutionSource,
      resolutionState: mapping.resolutionState,
      resolutionStatus: mapping.resolutionStatus,
      winningRuleId: mapping.winningRuleId,
      mappingAsOfDate,
      mappingFingerprint,
      metricRefreshAt
    });
  }
  return output;
}

function writeMetricContract(workbook: ExcelScript.Workbook): void {
  const sheet = resetOutputSheet(workbook, "Metric Contract", workbook.getTable("tblMetricContract"));
  writeTitle(
    sheet,
    "Metric Contract",
    "Reporting Group Metric Contract",
    "Phase 2A deterministic contract only. Performance, Reports, _Metric_Calc, and KPI-0001 remain unchanged; Attach Rate is out of scope.",
    9
  );
  const headers = [
    "MetricID", "DisplayName", "ValueType", "Numerator", "Denominator",
    "ScopeDimensions", "MappingMembership", "Phase2AStatus", "Notes"
  ];
  const rows: (string | number | boolean)[][] = [
    ["MEASURE-SALES-NOK", "Sales NOK", "Currency", "SUM(SalesAmount)", "Not applicable", "ImportID; optional ReportingChannel; optional RestaurantID; PublicationState", "All mapping states", "Implemented for bridge QA", "Source totals are never rewritten."],
    ["MEASURE-QUANTITY", "Quantity", "Number", "SUM(Quantity)", "Not applicable", "ImportID; optional ReportingChannel; optional RestaurantID; PublicationState", "All mapping states", "Implemented for bridge QA", "Signed source quantities are preserved."],
    ["MEASURE-FACT-COUNT", "Fact Count", "Count", "COUNT(SalesFactID)", "Not applicable", "ImportID; optional ReportingChannel; optional RestaurantID; PublicationState", "All mapping states", "Implemented for bridge QA", "One derived row per source fact."],
    ["METRIC-RPG-SALES-SHARE", "Reporting Group Sales Share", "Percentage", "Mapped Sales NOK for selected ReportingGroupID", "All Sales NOK in identical selected scope", "ImportID; optional ReportingChannel; optional RestaurantID; PublicationState", "Numerator = Mapped selected RPG; denominator = all states", "Contract + QA only", "Performance cutover is Phase 2B."],
    ["METRIC-RPG-QUANTITY-SHARE", "Reporting Group Quantity Share", "Percentage", "Mapped Quantity for selected ReportingGroupID", "All Quantity in identical selected scope", "ImportID; optional ReportingChannel; optional RestaurantID; PublicationState", "Numerator = Mapped selected RPG; denominator = all states", "Contract + QA only", "KPI exposure remains Draft/out of scope."]
  ];
  addOutputTable(sheet, "Metric Contract", 4, headers, rows, "tblMetricContract", "TableStyleMedium2");
  sheet.getFreezePanes().freezeRows(4);
  setWidths(sheet, [180, 190, 100, 270, 260, 310, 310, 165, 260]);
}

function ensureLegacyEquivalenceTable(
  workbook: ExcelScript.Workbook,
  groups: ReportingGroup[]
): ExcelScript.Table {
  const sheet = workbook.getWorksheet("Metric Equivalence") ?? workbook.addWorksheet("Metric Equivalence");
  let table = workbook.getTable("tblLegacyRPGEquivalence");
  if (!table) {
    const used = sheet.getUsedRange();
    if (used) used.clear(ExcelScript.ClearApplyTo.all);
    writeTitle(
      sheet,
      "Metric Equivalence",
      "Legacy CAT to Reporting Group Equivalence",
      "Human-authored ID-to-ID comparison definitions. Repeat DefinitionID for multiple CAT members. Pulse never infers equivalence from names.",
      8
    );
    checkedRangeByIndexes(sheet, "Metric Equivalence", 2, 0, 1, 8, "equivalence headers").setValues([[
      "DefinitionID", "ReportingGroupID", "LegacyReportingCategoryID", "ComparisonStatus",
      "Active", "Notes", "ApprovedBy", "ApprovedAt"
    ]]);
    table = sheet.addTable(
      checkedRangeByIndexes(sheet, "Metric Equivalence", 2, 0, 1, 8, "equivalence table"),
      true
    );
    table.setName("tblLegacyRPGEquivalence");
    table.setPredefinedTableStyle("TableStyleMedium4");
    checkedRangeByIndexes(sheet, "Metric Equivalence", 3, 3, 997, 1, "ComparisonStatus validation").getDataValidation().setRule({
      list: { inCellDropDown: true, source: "Equivalent,Partial,Not Comparable" }
    });
    checkedRangeByIndexes(sheet, "Metric Equivalence", 3, 4, 997, 1, "Active validation").getDataValidation().setRule({
      list: { inCellDropDown: true, source: "Yes,No" }
    });
    const groupIds = groups.map(group => group.id);
    if (groupIds.length && groupIds.join(",").length <= 255) {
      checkedRangeByIndexes(sheet, "Metric Equivalence", 3, 1, 997, 1, "ReportingGroupID validation").getDataValidation().setRule({
        list: { inCellDropDown: true, source: groupIds.join(",") }
      });
    }
    checkedRangeByIndexes(sheet, "Metric Equivalence", 3, 0, 997, 8, "editable definition area").getFormat().getFill().setColor("#FFFDF5");
    checkedRangeByIndexes(sheet, "Metric Equivalence", 3, 7, 997, 1, "ApprovedAt format").setNumberFormatLocal("dd.mm.yyyy hh:mm");
    sheet.getFreezePanes().freezeRows(3);
    setWidths(sheet, [135, 145, 175, 135, 70, 280, 130, 130]);
  }
  return table;
}

function readEquivalenceDefinitions(table: ExcelScript.Table): EquivalenceDefinition[] {
  const h = headerMap(table);
  return table.getRangeBetweenHeaderAndTotal().getValues()
    .filter(row => text(row[h.DefinitionID]) || text(row[h.ReportingGroupID]) || text(row[h.LegacyReportingCategoryID]))
    .map(row => ({
      definitionId: text(row[h.DefinitionID]),
      groupId: text(row[h.ReportingGroupID]),
      legacyCategoryId: text(row[h.LegacyReportingCategoryID]),
      comparisonStatus: text(row[h.ComparisonStatus]),
      active: text(row[h.Active]),
      notes: text(row[h.Notes]),
      approvedBy: text(row[h.ApprovedBy]),
      approvedAt: numberValue(row[h.ApprovedAt])
    }));
}

function validateEquivalenceDefinitions(
  definitions: EquivalenceDefinition[],
  groups: ReportingGroup[],
  categoryIds: string[]
): string[] {
  const errors: string[] = [];
  const groupIds = new Set<string>();
  for (const group of groups) groupIds.add(group.id);
  const knownCategories = new Set<string>();
  for (const id of categoryIds) knownCategories.add(id);
  const allowedStatus = new Set<string>();
  for (const value of ["Equivalent", "Partial", "Not Comparable"]) allowedStatus.add(value);
  const byDefinition = new Map<string, { groupId: string; status: string; categories: Set<string> }>();
  for (const row of definitions) {
    if (!row.definitionId) {
      errors.push("blank DefinitionID");
      continue;
    }
    if (!groupIds.has(row.groupId)) errors.push(`${row.definitionId} unknown ReportingGroupID ${row.groupId}`);
    if (!knownCategories.has(row.legacyCategoryId)) errors.push(`${row.definitionId} unknown legacy category ${row.legacyCategoryId}`);
    if (!allowedStatus.has(row.comparisonStatus)) errors.push(`${row.definitionId} invalid ComparisonStatus ${row.comparisonStatus}`);
    if (row.active !== "Yes" && row.active !== "No") errors.push(`${row.definitionId} invalid Active ${row.active}`);
    const current = byDefinition.get(row.definitionId) ?? {
      groupId: row.groupId,
      status: row.comparisonStatus,
      categories: new Set<string>()
    };
    if (current.groupId !== row.groupId) errors.push(`${row.definitionId} targets multiple ReportingGroupIDs`);
    if (current.status !== row.comparisonStatus) errors.push(`${row.definitionId} has inconsistent ComparisonStatus`);
    if (current.categories.has(row.legacyCategoryId)) errors.push(`${row.definitionId} repeats ${row.legacyCategoryId}`);
    current.categories.add(row.legacyCategoryId);
    byDefinition.set(row.definitionId, current);
  }
  return errors;
}

function readLegacyCategoryIds(table: ExcelScript.Table): string[] {
  const h = headerMap(table);
  return table.getRangeBetweenHeaderAndTotal().getValues()
    .map(row => text(row[h.ReportingCategoryID]))
    .filter(id => id.length > 0);
}

function writeMetricBridge(workbook: ExcelScript.Workbook, rows: MetricFact[]): void {
  const sheet = resetOutputSheet(workbook, "_Metric_RPG_Facts", workbook.getTable("tblMetricRPGFacts"));
  writeTitle(
    sheet,
    "_Metric_RPG_Facts",
    "Reporting Group Metric Facts",
    "Derived analysis bridge. One row per immutable sales fact joined to the current Effective Mapping state; never edit this table manually.",
    20
  );
  const headers = [
    "SalesFactID", "ImportID", "RestaurantID", "ProductID", "ReportingChannel",
    "PeriodStart", "PeriodEnd", "PublicationState", "Quantity", "SalesAmount",
    "LegacyReportingCategoryID", "EffectiveReportingGroupID", "EffectiveReportingGroupName",
    "ResolutionSource", "ResolutionState", "ResolutionStatus", "WinningRuleID",
    "MappingAsOfDate", "MappingFingerprint", "MetricRefreshAt"
  ];
  const values = rows.map(row => [
    row.salesFactId, row.importId, row.restaurantId, row.productId, row.reportingChannel,
    row.periodStart, row.periodEnd, row.publicationState, row.quantity, row.salesAmount,
    row.legacyReportingCategoryId, row.effectiveGroupId, row.effectiveGroupName,
    row.resolutionSource, row.resolutionState, row.resolutionStatus, row.winningRuleId,
    row.mappingAsOfDate, row.mappingFingerprint, row.metricRefreshAt
  ]);
  addOutputTable(sheet, "_Metric_RPG_Facts", 4, headers, values, "tblMetricRPGFacts", "TableStyleMedium2");
  if (rows.length) {
    checkedRangeByIndexes(sheet, "_Metric_RPG_Facts", 4, 5, rows.length, 2, "PeriodStart/PeriodEnd format").setNumberFormatLocal("dd.mm.yyyy");
    checkedRangeByIndexes(sheet, "_Metric_RPG_Facts", 4, 8, rows.length, 2, "Quantity/SalesAmount format").setNumberFormat("#,##0.00");
    checkedRangeByIndexes(sheet, "_Metric_RPG_Facts", 4, 17, rows.length, 1, "MappingAsOfDate format").setNumberFormatLocal("dd.mm.yyyy");
    checkedRangeByIndexes(sheet, "_Metric_RPG_Facts", 4, 19, rows.length, 1, "MetricRefreshAt format").setNumberFormatLocal("dd.mm.yyyy hh:mm");
  }
  sheet.getFreezePanes().freezeRows(4);
  setWidths(sheet, [120, 120, 100, 105, 105, 90, 90, 115, 90, 105, 135, 145, 175, 130, 110, 110, 120, 105, 190, 130]);
}

function readMetricBridge(table: ExcelScript.Table): MetricFact[] {
  const h = headerMap(table);
  return table.getRangeBetweenHeaderAndTotal().getValues()
    .filter(row => text(row[h.SalesFactID]))
    .map(row => ({
      salesFactId: text(row[h.SalesFactID]),
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
    }));
}

function validateMaterializedBridge(
  facts: SourceFact[],
  bridge: MetricFact[],
  asOfDate: number,
  mappingFingerprint: string
): string[] {
  const errors: string[] = [];
  if (bridge.length !== facts.length) errors.push(`row count ${bridge.length}; expected ${facts.length}`);
  const factIds = new Set<string>();
  for (const fact of facts) factIds.add(fact.salesFactId);
  const bridgeIds = new Set<string>();
  for (const row of bridge) {
    if (!factIds.has(row.salesFactId)) errors.push(`unknown SalesFactID ${row.salesFactId}`);
    if (bridgeIds.has(row.salesFactId)) errors.push(`duplicate SalesFactID ${row.salesFactId}`);
    bridgeIds.add(row.salesFactId);
    if (row.mappingAsOfDate !== asOfDate) errors.push(`${row.salesFactId} MappingAsOfDate ${row.mappingAsOfDate}; expected ${asOfDate}`);
    if (row.mappingFingerprint !== mappingFingerprint) errors.push(`${row.salesFactId} mapping fingerprint mismatch`);
    if (!isMappingState(row.resolutionStatus)) errors.push(`${row.salesFactId} unsupported status ${row.resolutionStatus}`);
  }
  return errors;
}

function buildScopes(facts: SourceFact[]): MetricScope[] {
  const scopes: MetricScope[] = [{
    publicationScope: "All facts",
    publicationState: "",
    importId: "All imports",
    channel: "All channels"
  }];
  const imports = new Set<string>();
  const channelsByImport = new Map<string, Set<string>>();
  for (const fact of facts) {
    if (fact.publicationState !== "Active Finalized") continue;
    imports.add(fact.importId);
    const channels = channelsByImport.get(fact.importId) ?? new Set<string>();
    if (fact.reportingChannel) channels.add(fact.reportingChannel);
    channelsByImport.set(fact.importId, channels);
  }
  const importIds: string[] = [];
  imports.forEach(id => importIds.push(id));
  importIds.sort();
  for (const importId of importIds) {
    scopes.push({
      publicationScope: "Active Finalized",
      publicationState: "Active Finalized",
      importId,
      channel: "All channels"
    });
    const channels: string[] = [];
    const channelSet = channelsByImport.get(importId);
    if (channelSet) channelSet.forEach(channel => channels.push(channel));
    channels.sort();
    for (const channel of channels) scopes.push({
      publicationScope: "Active Finalized",
      publicationState: "Active Finalized",
      importId,
      channel
    });
  }
  return scopes;
}

function reconcileFactsAndBridge(
  facts: SourceFact[],
  bridge: MetricFact[],
  scopes: MetricScope[]
): Reconciliation[] {
  return scopes.map(scope => {
    const source = totals(facts.filter(row => inScope(row, scope)));
    const derivedRows = bridge.filter(row => inScope(row, scope));
    const derived = totals(derivedRows);
    const coverage = stateCoverage(derivedRows);
    let stateTotals = emptyTotals();
    for (const state of MAPPING_STATES) stateTotals = addTotals(stateTotals, coverage[state]);
    const factVariance = derived.factCount - source.factCount;
    const salesVariance = derived.sales - source.sales;
    const quantityVariance = derived.quantity - source.quantity;
    const stateFactVariance = stateTotals.factCount - source.factCount;
    const stateSalesVariance = stateTotals.sales - source.sales;
    const stateQuantityVariance = stateTotals.quantity - source.quantity;
    const result = factVariance === 0 && stateFactVariance === 0 &&
      almostEqual(salesVariance, 0) && almostEqual(stateSalesVariance, 0) &&
      almostEqual(quantityVariance, 0) && almostEqual(stateQuantityVariance, 0)
      ? "PASS"
      : "FAIL";
    return {
      scope, source, bridge: derived, coverage, stateTotals,
      factVariance, salesVariance, quantityVariance,
      stateFactVariance, stateSalesVariance, stateQuantityVariance, result
    };
  });
}

function aggregateReportingGroups(bridge: MetricFact[], scopes: MetricScope[]): GroupTotal[] {
  const output: GroupTotal[] = [];
  for (const scope of scopes) {
    const byGroup = new Map<string, GroupTotal>();
    for (const row of bridge) {
      if (!inScope(row, scope) || row.resolutionStatus !== "Mapped") continue;
      const current = byGroup.get(row.effectiveGroupId) ?? {
        ...scope,
        groupId: row.effectiveGroupId,
        groupName: row.effectiveGroupName,
        factCount: 0,
        sales: 0,
        quantity: 0
      };
      current.factCount += 1;
      current.sales += row.salesAmount;
      current.quantity += row.quantity;
      byGroup.set(row.effectiveGroupId, current);
    }
    byGroup.forEach(value => output.push(value));
  }
  output.sort((a, b) => scopeKey(a).localeCompare(scopeKey(b)) || a.groupId.localeCompare(b.groupId));
  return output;
}

function buildLegacyRpgCrosswalk(bridge: MetricFact[], scopes: MetricScope[]): CrosswalkRow[] {
  const output: CrosswalkRow[] = [];
  for (const scope of scopes) {
    const map = new Map<string, CrosswalkRow>();
    for (const row of bridge) {
      if (!inScope(row, scope)) continue;
      const key = [row.legacyReportingCategoryId, row.effectiveGroupId, row.resolutionStatus].join("\u001f");
      const current = map.get(key) ?? {
        ...scope,
        legacyCategoryId: row.legacyReportingCategoryId,
        groupId: row.effectiveGroupId,
        groupName: row.effectiveGroupName,
        resolutionStatus: row.resolutionStatus,
        factCount: 0,
        sales: 0,
        quantity: 0
      };
      current.factCount += 1;
      current.sales += row.salesAmount;
      current.quantity += row.quantity;
      map.set(key, current);
    }
    map.forEach(value => output.push(value));
  }
  output.sort((a, b) =>
    scopeKey(a).localeCompare(scopeKey(b)) ||
    a.legacyCategoryId.localeCompare(b.legacyCategoryId) ||
    a.groupId.localeCompare(b.groupId) ||
    a.resolutionStatus.localeCompare(b.resolutionStatus)
  );
  return output;
}

function compareLegacyDefinitions(
  bridge: MetricFact[],
  definitions: EquivalenceDefinition[],
  scopes: MetricScope[]
): LegacyComparison[] {
  const active = new Map<string, {
    definitionId: string; groupId: string; status: string; categoryIds: string[];
  }>();
  for (const row of definitions) {
    if (row.active !== "Yes") continue;
    const current = active.get(row.definitionId) ?? {
      definitionId: row.definitionId,
      groupId: row.groupId,
      status: row.comparisonStatus,
      categoryIds: []
    };
    current.categoryIds.push(row.legacyCategoryId);
    active.set(row.definitionId, current);
  }
  const output: LegacyComparison[] = [];
  active.forEach(definition => {
    definition.categoryIds.sort();
    const legacyIds = new Set<string>();
    for (const id of definition.categoryIds) legacyIds.add(id);
    for (const scope of scopes) {
      const scopeRows = bridge.filter(row => inScope(row, scope));
      const legacy = totals(scopeRows.filter(row => legacyIds.has(row.legacyReportingCategoryId)));
      const rpg = totals(scopeRows.filter(row =>
        row.resolutionStatus === "Mapped" && row.effectiveGroupId === definition.groupId
      ));
      const denominator = totals(scopeRows);
      const factVariance = rpg.factCount - legacy.factCount;
      const salesVariance = rpg.sales - legacy.sales;
      const quantityVariance = rpg.quantity - legacy.quantity;
      const equal = factVariance === 0 && almostEqual(salesVariance, 0) && almostEqual(quantityVariance, 0);
      output.push({
        ...scope,
        definitionId: definition.definitionId,
        comparisonStatus: definition.status,
        groupId: definition.groupId,
        legacyCategoryIds: definition.categoryIds.join(", "),
        scopeTotals: denominator,
        legacy,
        rpg,
        factVariance,
        salesVariance,
        quantityVariance,
        legacySalesShare: denominator.sales ? legacy.sales / denominator.sales : 0,
        rpgSalesShare: denominator.sales ? rpg.sales / denominator.sales : 0,
        shareVariance: denominator.sales ? salesVariance / denominator.sales : 0,
        result: definition.status === "Equivalent" ? (equal ? "PASS" : "VARIANCE") : "INFO"
      });
    }
  });
  output.sort((a, b) => a.definitionId.localeCompare(b.definitionId) || scopeKey(a).localeCompare(scopeKey(b)));
  return output;
}

function writeMetricQA(
  workbook: ExcelScript.Workbook,
  asOfDate: number,
  mappingFingerprint: string,
  refreshAt: number,
  factSnapshot: Snapshot,
  bridgeRows: number,
  reconciliations: Reconciliation[],
  groupTotals: GroupTotal[],
  crosswalk: CrosswalkRow[],
  comparisons: LegacyComparison[],
  definitionCount: number,
  definitionErrors: string[]
): void {
  const sheet = resetOutputSheet(
    workbook,
    "Metric Migration QA",
    workbook.getTable("tblMetricQA"),
    workbook.getTable("tblMetricReconciliation"),
    workbook.getTable("tblMetricReportingGroupTotals"),
    workbook.getTable("tblLegacyRPGCrosswalk"),
    workbook.getTable("tblLegacyRPGComparison")
  );
  writeTitle(
    sheet,
    "Metric Migration QA",
    "Phase 2A Metric Migration QA",
    "Central fact, mapping-state, Reporting Group, and human-configured legacy-equivalence reconciliation. Performance remains unchanged.",
    28
  );
  checkedRangeByIndexes(sheet, "Metric Migration QA", 2, 0, 1, 6, "QA mapping summary").setValues([[
    "MappingAsOfDate", asOfDate, "MappingFingerprint", mappingFingerprint, "MetricRefreshAt", refreshAt
  ]]);
  checkedRangeByIndexes(sheet, "Metric Migration QA", 2, 1, 1, 1, "QA MappingAsOfDate format").setNumberFormatLocal("dd.mm.yyyy");
  checkedRangeByIndexes(sheet, "Metric Migration QA", 2, 5, 1, 1, "QA MetricRefreshAt format").setNumberFormatLocal("dd.mm.yyyy hh:mm");

  const factFailures = reconciliations.filter(row => row.factVariance !== 0 || row.stateFactVariance !== 0).length;
  const salesFailures = reconciliations.filter(row =>
    !almostEqual(row.salesVariance, 0) || !almostEqual(row.stateSalesVariance, 0)
  ).length;
  const quantityFailures = reconciliations.filter(row =>
    !almostEqual(row.quantityVariance, 0) || !almostEqual(row.stateQuantityVariance, 0)
  ).length;
  const invalidStates = reconciliations.some(row =>
    row.stateFactVariance !== 0 || !almostEqual(row.stateSalesVariance, 0) || !almostEqual(row.stateQuantityVariance, 0)
  );
  const equivalentVariances = comparisons.filter(row => row.comparisonStatus === "Equivalent" && row.result === "VARIANCE").length;
  const checks: (string | number | boolean)[][] = [
    ["QA-0302A-01", "Effective Mapping current and recomputed", "PASS", asOfDate, "Script fails before output if Phase 1 Effective Mapping differs from current rules, groups, products, or today's AsOfDate."],
    ["QA-0302A-02", "Mapping fingerprint reproducible", "PASS", mappingFingerprint, "Semantic v1 fingerprint is row-order independent and stored on every derived fact."],
    ["QA-0302A-03", "All source fact rows accounted for", bridgeRows === factSnapshot.rows ? "PASS" : "FAIL", bridgeRows - factSnapshot.rows, "One derived bridge row is required for every tblSalesFacts row."],
    ["QA-0302A-04", "Sales NOK reconciliation", salesFailures === 0 ? "PASS" : "FAIL", salesFailures, "Source, bridge, and four-state totals must agree in every generated scope."],
    ["QA-0302A-05", "Quantity reconciliation", quantityFailures === 0 ? "PASS" : "FAIL", quantityFailures, "Signed source quantity is retained and reconciled."],
    ["QA-0302A-06", "Fact-count reconciliation", factFailures === 0 ? "PASS" : "FAIL", factFailures, "Source, bridge, and state counts must agree exactly."],
    ["QA-0302A-07", "Mapping-state coverage", invalidStates ? "FAIL" : "PASS", invalidStates ? 1 : 0, "Every derived fact is Mapped, Unmapped, Conflict, or Inactive Target."],
    ["QA-0302A-08", "Legacy equivalence configuration", definitionErrors.length ? "FAIL" : definitionCount ? "PASS" : "WARN", definitionErrors.length, definitionErrors.length ? definitionErrors.join(" | ") : definitionCount ? "Human-authored ID definitions validated." : "No equivalence definitions entered yet; no equivalence was inferred."],
    ["QA-0302A-09", "Equivalent legacy/RPG variance visible", equivalentVariances ? "WARN" : "PASS", equivalentVariances, "Material differences remain visible in tblLegacyRPGComparison; they are never silently accepted."],
    ["QA-0302A-10", "Legacy Performance path preserved", "PASS", 0, "_Metric_Calc, Performance, Reports, and KPI Registry were fingerprinted before/after and remained unchanged."],
    ["QA-0302A-11", "Source facts unchanged", "PASS", factSnapshot.fingerprint, "tblSalesFacts content fingerprint remained unchanged through bridge creation."]
  ];
  addOutputTable(sheet, "Metric Migration QA", 5, ["CheckID", "Check", "Result", "Observed", "Explanation"], checks, "tblMetricQA", "TableStyleMedium2");

  const reconciliationRows = reconciliations.map(row => {
    const mapped = row.coverage["Mapped"];
    const unmapped = row.coverage["Unmapped"];
    const conflict = row.coverage["Conflict"];
    const inactive = row.coverage["Inactive Target"];
    return [
      row.scope.publicationScope, row.scope.importId, row.scope.channel,
      row.source.factCount, row.bridge.factCount, row.factVariance,
      row.source.sales, row.bridge.sales, row.salesVariance,
      row.source.quantity, row.bridge.quantity, row.quantityVariance,
      mapped.factCount, mapped.sales, mapped.quantity,
      unmapped.factCount, unmapped.sales, unmapped.quantity,
      conflict.factCount, conflict.sales, conflict.quantity,
      inactive.factCount, inactive.sales, inactive.quantity,
      row.stateFactVariance, row.stateSalesVariance, row.stateQuantityVariance, row.result
    ];
  });
  let row = 19;
  addOutputTable(sheet, "Metric Migration QA", row, [
    "PublicationScope", "ImportID", "ChannelScope",
    "SourceFactCount", "BridgeFactCount", "FactCountVariance",
    "SourceSalesNOK", "BridgeSalesNOK", "SalesVariance",
    "SourceQuantity", "BridgeQuantity", "QuantityVariance",
    "MappedFactCount", "MappedSalesNOK", "MappedQuantity",
    "UnmappedFactCount", "UnmappedSalesNOK", "UnmappedQuantity",
    "ConflictFactCount", "ConflictSalesNOK", "ConflictQuantity",
    "InactiveTargetFactCount", "InactiveTargetSalesNOK", "InactiveTargetQuantity",
    "StateFactCountVariance", "StateSalesVariance", "StateQuantityVariance", "Result"
  ], reconciliationRows, "tblMetricReconciliation", "TableStyleMedium2");
  row += Math.max(1, reconciliationRows.length) + 4;

  const groupRows = groupTotals.map(value => [
    value.publicationScope, value.importId, value.channel, value.groupId, value.groupName,
    value.factCount, value.sales, value.quantity
  ]);
  addOutputTable(sheet, "Metric Migration QA", row, [
    "PublicationScope", "ImportID", "ChannelScope", "ReportingGroupID",
    "ReportingGroupName", "FactCount", "SalesNOK", "Quantity"
  ], groupRows, "tblMetricReportingGroupTotals", "TableStyleMedium4");
  row += Math.max(1, groupRows.length) + 4;

  const comparisonRows = comparisons.map(value => [
    value.definitionId, value.comparisonStatus, value.publicationScope, value.importId, value.channel,
    value.groupId, value.legacyCategoryIds,
    value.scopeTotals.factCount, value.scopeTotals.sales, value.scopeTotals.quantity,
    value.legacy.factCount, value.rpg.factCount, value.factVariance,
    value.legacy.sales, value.rpg.sales, value.salesVariance,
    value.legacy.quantity, value.rpg.quantity, value.quantityVariance,
    value.legacySalesShare, value.rpgSalesShare, value.shareVariance, value.result
  ]);
  addOutputTable(sheet, "Metric Migration QA", row, [
    "DefinitionID", "ComparisonStatus", "PublicationScope", "ImportID", "ChannelScope",
    "ReportingGroupID", "LegacyReportingCategoryIDs",
    "ScopeFactCount", "ScopeSalesNOK", "ScopeQuantity",
    "LegacyFactCount", "RPGFactCount", "FactCountVariance",
    "LegacySalesNOK", "RPGSalesNOK", "SalesVariance",
    "LegacyQuantity", "RPGQuantity", "QuantityVariance",
    "LegacySalesShare", "RPGSalesShare", "SalesShareVariance", "Result"
  ], comparisonRows, "tblLegacyRPGComparison", "TableStyleMedium5");
  row += Math.max(1, comparisonRows.length) + 4;

  const crosswalkRows = crosswalk.map(value => [
    value.publicationScope, value.importId, value.channel, value.legacyCategoryId,
    value.groupId, value.groupName, value.resolutionStatus,
    value.factCount, value.sales, value.quantity
  ]);
  addOutputTable(sheet, "Metric Migration QA", row, [
    "PublicationScope", "ImportID", "ChannelScope", "LegacyReportingCategoryID",
    "EffectiveReportingGroupID", "EffectiveReportingGroupName", "ResolutionStatus",
    "FactCount", "SalesNOK", "Quantity"
  ], crosswalkRows, "tblLegacyRPGCrosswalk", "TableStyleMedium3");

  const used = sheet.getUsedRange();
  if (used) used.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  sheet.getFreezePanes().freezeRows(5);
  setWidths(sheet, [130, 135, 130, 145, 165, 130, 135, 100, 110, 110, 110, 110, 110, 110, 110, 110, 110, 110, 110, 110, 110, 120, 120, 120, 120, 120, 120, 90]);
}

function stateCoverage(rows: MetricFact[]): { [key: string]: Totals } {
  const result: { [key: string]: Totals } = {};
  for (const state of MAPPING_STATES) result[state] = emptyTotals();
  for (const row of rows) {
    if (!isMappingState(row.resolutionStatus)) {
      throw new Error(`PUL-0302A-014: Unsupported bridge ResolutionStatus ${row.resolutionStatus}.`);
    }
    const bucket = result[row.resolutionStatus];
    bucket.factCount += 1;
    bucket.sales += row.salesAmount;
    bucket.quantity += row.quantity;
  }
  return result;
}

function totals(rows: { salesAmount: number; quantity: number }[]): Totals {
  const result = emptyTotals();
  for (const row of rows) {
    result.factCount += 1;
    result.sales += row.salesAmount;
    result.quantity += row.quantity;
  }
  return result;
}

function inScope(row: SourceFact | MetricFact, scope: MetricScope): boolean {
  if (scope.publicationState && row.publicationState !== scope.publicationState) return false;
  if (scope.importId !== "All imports" && row.importId !== scope.importId) return false;
  if (scope.channel !== "All channels" && row.reportingChannel !== scope.channel) return false;
  return true;
}

function scopeKey(scope: MetricScope): string {
  return [scope.publicationScope, scope.importId, scope.channel].join("\u001f");
}

function emptyTotals(): Totals {
  return { factCount: 0, sales: 0, quantity: 0 };
}

function addTotals(left: Totals, right: Totals): Totals {
  return {
    factCount: left.factCount + right.factCount,
    sales: left.sales + right.sales,
    quantity: left.quantity + right.quantity
  };
}

function snapshotTable(table: ExcelScript.Table): Snapshot {
  const h = headerMap(table);
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  let sales = 0;
  let quantity = 0;
  if (h.SalesAmount !== undefined && h.Quantity !== undefined) {
    for (const row of rows) {
      sales += numberValue(row[h.SalesAmount]);
      quantity += numberValue(row[h.Quantity]);
    }
  }
  const state = newHashState();
  updateHashMatrix(state, [table.getHeaderRowRange().getValues()[0]]);
  updateHashMatrix(state, rows);
  return { rows: rows.length, sales, quantity, fingerprint: finishHash(state, "DATA-") };
}

function snapshotProtectedSurfaces(workbook: ExcelScript.Workbook, kpiTable: ExcelScript.Table): string {
  const state = newHashState();
  const metricCalcUsed = requiredSheet(workbook, "_Metric_Calc").getUsedRange();
  const performanceUsed = requiredSheet(workbook, "Performance").getUsedRange();
  const reportsUsed = requiredSheet(workbook, "Reports").getUsedRange();
  const metricCalcAddress = metricCalcUsed ? metricCalcUsed.getAddress() : "";
  const metricCalcValues = metricCalcUsed ? metricCalcUsed.getValues() : [];
  const metricCalcFormulas = metricCalcUsed ? metricCalcUsed.getFormulas() : [];
  const performanceAddress = performanceUsed ? performanceUsed.getAddress() : "";
  const performanceValues = performanceUsed ? performanceUsed.getValues() : [];
  const performanceFormulas = performanceUsed ? performanceUsed.getFormulas() : [];
  const reportsAddress = reportsUsed ? reportsUsed.getAddress() : "";
  const reportsValues = reportsUsed ? reportsUsed.getValues() : [];
  const reportsFormulas = reportsUsed ? reportsUsed.getFormulas() : [];
  updateProtectedSnapshot(state, "_Metric_Calc", metricCalcAddress, metricCalcValues, metricCalcFormulas);
  updateProtectedSnapshot(state, "Performance", performanceAddress, performanceValues, performanceFormulas);
  updateProtectedSnapshot(state, "Reports", reportsAddress, reportsValues, reportsFormulas);
  updateHash(state, "tblKPIRegistry");
  updateHashMatrix(state, [kpiTable.getHeaderRowRange().getValues()[0]]);
  updateHashMatrix(state, kpiTable.getRangeBetweenHeaderAndTotal().getValues());
  return finishHash(state, "PROTECTED-");
}

function updateProtectedSnapshot(
  state: HashState,
  name: string,
  address: string,
  values: (string | number | boolean)[][],
  formulas: (string | number | boolean)[][]
): void {
  updateHash(state, name);
  updateHash(state, address);
  updateHashMatrix(state, values);
  updateHashMatrix(state, formulas);
}

function sameSnapshot(left: Snapshot, right: Snapshot): boolean {
  return left.rows === right.rows &&
    almostEqual(left.sales, right.sales) &&
    almostEqual(left.quantity, right.quantity) &&
    left.fingerprint === right.fingerprint;
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

function updateHashMatrix(
  state: HashState,
  values: (string | number | boolean)[][]
): void {
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

function addOutputTable(
  sheet: ExcelScript.Worksheet,
  sheetName: string,
  startRow: number,
  headers: string[],
  rows: (string | number | boolean)[][],
  tableName: string,
  style: string
): ExcelScript.Table {
  if (!headers.length) throw new Error(`PUL-0302A-018: ${sheetName} ${tableName} has no headers.`);
  const dataRowCount = rows.length;
  validateOutputRows(rows, headers.length, sheetName, tableName);
  const headerRange = checkedRangeByIndexes(
    sheet,
    sheetName,
    startRow - 1,
    0,
    1,
    headers.length,
    `${tableName} header`
  );
  headerRange.setValues([headers]);
  const chunkSize = 2000;
  for (let offset = 0; offset < rows.length; offset += chunkSize) {
    const chunk = rows.slice(offset, Math.min(offset + chunkSize, rows.length));
    checkedRangeByIndexes(
      sheet,
      sheetName,
      startRow + offset,
      0,
      chunk.length,
      headers.length,
      `${tableName} data rows ${offset + 1}-${offset + chunk.length}`
    ).setValues(chunk);
  }
  const tableRange = checkedRangeByIndexes(
    sheet,
    sheetName,
    startRow - 1,
    0,
    dataRowCount + 1,
    headers.length,
    `${tableName} table`
  );
  const table = sheet.addTable(tableRange, true);
  table.setName(tableName);
  table.setPredefinedTableStyle(style);
  styleHeader(headerRange);
  return table;
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
        `PUL-0302A-019: ${sheetName} ${tableName} row ${index + 1} has ${rows[index].length} values; expected ${expectedColumns}.`
      );
    }
  }
}

function columnName(count: number): string {
  let value = count;
  let result = "";
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function checkedRangeByIndexes(
  sheet: ExcelScript.Worksheet,
  sheetName: string,
  startRow: number,
  startColumn: number,
  rowCount: number,
  columnCount: number,
  context: string
): ExcelScript.Range {
  const maxRows = 1048576;
  const maxColumns = 16384;
  const valid = Number.isInteger(startRow) && Number.isInteger(startColumn) &&
    Number.isInteger(rowCount) && Number.isInteger(columnCount) &&
    startRow >= 0 && startColumn >= 0 && rowCount > 0 && columnCount > 0 &&
    startRow + rowCount <= maxRows && startColumn + columnCount <= maxColumns;
  const address = valid
    ? indexedAddress(startRow, startColumn, rowCount, columnCount)
    : "[invalid geometry]";
  if (!valid) {
    throw new Error(
      `PUL-0302A-020: Invalid worksheet range ${sheetName}!${address} for ${context}; ` +
      `startRow=${startRow}, startColumn=${startColumn}, rowCount=${rowCount}, columnCount=${columnCount}.`
    );
  }
  try {
    return sheet.getRangeByIndexes(startRow, startColumn, rowCount, columnCount);
  } catch (error) {
    throw new Error(
      `PUL-0302A-021: Worksheet range acquisition failed at ${sheetName}!${address} for ${context}. ${errorMessage(error)}`
    );
  }
}

function indexedAddress(
  startRow: number,
  startColumn: number,
  rowCount: number,
  columnCount: number
): string {
  const firstColumn = columnName(startColumn + 1);
  const lastColumn = columnName(startColumn + columnCount);
  const firstRow = startRow + 1;
  const lastRow = startRow + rowCount;
  return `${firstColumn}${firstRow}:${lastColumn}${lastRow}`;
}

function errorMessage(error: unknown): string {
  if (error instanceof Error) return error.message;
  return text(error);
}

function isMappingState(value: string): boolean {
  return value === "Mapped" || value === "Unmapped" || value === "Conflict" || value === "Inactive Target";
}

function ids(rules: MappingRule[]): string {
  return rules.map(rule => rule.id).join(", ");
}

function targets(rules: MappingRule[]): string {
  return rules.map(rule => rule.targetGroupId).join(", ");
}

function mainNodeId(sourceSystemId: string, mainCategory: string): string {
  return `${sourceSystemId} || Main || ${mainCategory}`;
}

function resetOutputSheet(
  workbook: ExcelScript.Workbook,
  name: string,
  ...tables: (ExcelScript.Table | undefined)[]
): ExcelScript.Worksheet {
  for (const table of tables) {
    if (table) table.delete();
  }
  const sheet = workbook.getWorksheet(name) ?? workbook.addWorksheet(name);
  const used = sheet.getUsedRange();
  if (used) used.clear(ExcelScript.ClearApplyTo.all);
  return sheet;
}

function requiredSheet(workbook: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
  const sheet = workbook.getWorksheet(name);
  if (!sheet) throw new Error(`PUL-0302A-015: Required sheet missing: ${name}`);
  return sheet;
}

function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error(`PUL-0302A-016: Required table missing: ${name}`);
  return table;
}

function headerMap(table: ExcelScript.Table): { [key: string]: number } {
  const headers = table.getHeaderRowRange().getValues()[0];
  const result: { [key: string]: number } = {};
  for (let index = 0; index < headers.length; index++) result[text(headers[index])] = index;
  return result;
}

function writeTitle(
  sheet: ExcelScript.Worksheet,
  sheetName: string,
  title: string,
  subtitle: string,
  columnCount: number
): void {
  const titleBand = checkedRangeByIndexes(sheet, sheetName, 0, 0, 1, columnCount, "title band");
  const subtitleBand = checkedRangeByIndexes(sheet, sheetName, 1, 0, 1, columnCount, "subtitle band");
  titleBand.getFormat().getFill().setColor("#172033");
  titleBand.getFormat().getFont().setColor("#FFFFFF");
  titleBand.getFormat().getFont().setBold(true);
  titleBand.getFormat().getFont().setSize(18);
  titleBand.getCell(0, 0).setValue(title);
  subtitleBand.getFormat().getFill().setColor("#EAF2FF");
  subtitleBand.getFormat().setWrapText(true);
  subtitleBand.getCell(0, 0).setValue(subtitle);
  titleBand.getFormat().setRowHeight(32);
  subtitleBand.getFormat().setRowHeight(30);
}

function styleHeader(range: ExcelScript.Range): void {
  range.getFormat().getFill().setColor("#4F8CFF");
  range.getFormat().getFont().setColor("#FFFFFF");
  range.getFormat().getFont().setBold(true);
  range.getFormat().setWrapText(true);
}

function setWidths(sheet: ExcelScript.Worksheet, widths: number[]): void {
  for (let index = 0; index < widths.length; index++) {
    sheet.getRangeByIndexes(0, index, 1, 1).getEntireColumn().getFormat().setColumnWidth(widths[index]);
  }
}

function updateEnvironment(
  table: ExcelScript.Table,
  key: string,
  value: string,
  note: string
): void {
  const body = table.getRangeBetweenHeaderAndTotal();
  const rows = body.getValues();
  let matchingIndex = -1;
  for (let index = 0; index < rows.length; index++) {
    if (text(rows[index][0]) === key) {
      matchingIndex = index;
      break;
    }
  }
  if (matchingIndex >= 0) {
    body.getCell(matchingIndex, 1).setValue(value);
    body.getCell(matchingIndex, 2).setValue(note);
    return;
  }
  table.addRow(-1, [key, value, note]);
}

function appendBuildLog(
  table: ExcelScript.Table,
  row: (string | number | boolean)[]
): void {
  table.addRow(-1, row);
}

function nextId(table: ExcelScript.Table, prefix: string, digits: number): string {
  let max = 0;
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  for (const row of values) {
    const match = text(row[0]).match(new RegExp(`^${prefix}-(\\d+)$`));
    if (match) max = Math.max(max, Number(match[1]));
  }
  return `${prefix}-${String(max + 1).padStart(digits, "0")}`;
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

function numberValue(value: unknown): number {
  const converted = Number(value);
  return Number.isFinite(converted) ? converted : 0;
}

function text(value: unknown): string {
  return String(value ?? "").trim();
}
